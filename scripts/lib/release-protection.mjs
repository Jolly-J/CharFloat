import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
import JavaScriptObfuscator from 'javascript-obfuscator';

export const ROOT = fileURLToPath(new URL('../../', import.meta.url));
export const STAGE = path.join(ROOT, 'dist/release-app');
export const MANIFEST = 'protection-manifest.json';
export const POLICY = JSON.parse(fs.readFileSync(new URL('../release-files.json', import.meta.url), 'utf8'));
export const PROFILE = 'offline-addons-v1';
export const sha256 = data => crypto.createHash('sha256').update(data).digest('hex');
export const ADDONS = ['wps-addon/addon-core.js', 'office-addon/public/taskpane.js'];

// 不重命名协议/宿主属性，不干扰调试器，不改变动态脚本执行环境。
// 固定版本和 seed 使同一输入可重建；编码只是混淆，不是加密安全边界。
export const OPTIONS = Object.freeze({
  compact: true, target: 'browser-no-eval', seed: 210023,
  identifierNamesGenerator: 'hexadecimal', renameGlobals: false, renameProperties: false,
  controlFlowFlattening: false, deadCodeInjection: false, debugProtection: false,
  debugProtectionInterval: 0, selfDefending: false, disableConsoleOutput: false,
  simplify: false, splitStrings: false, transformObjectKeys: false,
  stringArray: true, stringArrayThreshold: 0.75, stringArrayEncoding: ['base64'],
  stringArrayCallsTransform: false, stringArrayWrappersCount: 1,
  stringArrayRotate: true, stringArrayShuffle: true, sourceMap: false,
});

export function protectAddon(source, name) {
  let fingerprint;
  if (name === ADDONS[0]) {
    const old = source.match(/ADDON_BUILD_FINGERPRINT:\s*([a-f0-9]{64})/)?.[1];
    if (!old) throw new Error('WPS 构建指纹缺失');
    // 保护配置改变也必须触发部署升级；头部与 register 上报值一起更新。
    fingerprint = sha256(source + PROFILE + JSON.stringify(OPTIONS));
    source = source.replaceAll(old, fingerprint);
  }
  const code = JavaScriptObfuscator.obfuscate(source, OPTIONS).getObfuscatedCode();
  new vm.Script(code, { filename: name });
  return `// Release protection: ${PROFILE}\n` +
    (fingerprint ? `// ADDON_BUILD_FINGERPRINT: ${fingerprint}\n` : '') + code + '\n';
}

export function walk(dir, prefix = '') {
  return fs.readdirSync(dir, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name)).flatMap(e => {
    const rel = prefix ? `${prefix}/${e.name}` : e.name;
    if (e.isSymbolicLink()) throw new Error(`发布资源不得为软链接: ${rel}`);
    return e.isDirectory() ? walk(path.join(dir, e.name), rel) : [rel];
  });
}

export function allowed(name) {
  return POLICY.files.includes(name) || new RegExp(POLICY.rendererAssetPattern).test(name) ||
    name === 'package.json' || name === MANIFEST;
}

/**
 * 校验发布文件集合。
 *
 * `allowPlainCli`：**仅 Windows 包**需要。V8 字节码同时绑定「版本」与「编译平台」，
 * macOS 上编的 `cli.jsc` 到 Windows 会 `cachedDataRejected`（真机已证），
 * 而 macOS 上无法生成 Windows 的字节码，因此在 Windows 包内附带明文兜底
 * `dist/bridge/cli-full.cjs`，由加载器在字节码加载失败时回退。
 * 开启后**只**放行这一个文件，其余 `*-full.cjs` 仍禁止发布。
 */
export function checkFiles(names, read, { manifestRequired = true, allowPlainCli = false } = {}) {
  const errors = [];
  const PLAIN_CLI = 'dist/bridge/cli-full.cjs';
  const isAllowedPlain = n => allowPlainCli && n === PLAIN_CLI;
  // 兜底文件是打包阶段（beforePack）在清单写好后加进去的，不参与清单比对
  const own = names.filter(n => !n.startsWith('node_modules/') && !isAllowedPlain(n));
  for (const name of names) {
    if (/[.]map$|(?:^|\/)AGENTS[.]md$|(?:^|\/)[^/]*-full[.]cjs$/.test(name) && !isAllowedPlain(name)) errors.push(`禁止发布: ${name}`);
    if (!name.startsWith('node_modules/') && !isAllowedPlain(name) && !allowed(name)) errors.push(`不在发布白名单: ${name}`);
  }
  for (const name of [...POLICY.files, 'package.json', ...(manifestRequired ? [MANIFEST] : [])]) {
    if (!names.includes(name)) errors.push(`缺少运行资源: ${name}`);
  }
  if (errors.length) return errors;
  for (const name of ADDONS) {
    const code = read(name).toString();
    if (!code.startsWith(`// Release protection: ${PROFILE}\n`)) errors.push(`加载项未经过发布保护: ${name}`);
    try { new vm.Script(code); } catch { errors.push(`加载项语法错误: ${name}`); }
  }
  const shim = read('dist/bridge/cli.cjs').toString();
  if (Buffer.byteLength(shim) > 4096 || !shim.includes("require('bytenode')") || !shim.includes('cli.jsc')) {
    errors.push('CLI 不是字节码加载器');
  }
  if (read('dist/bridge/cli.jsc').length < 1024) errors.push('CLI 字节码异常');
  const pkg = JSON.parse(read('package.json'));
  if (pkg.main !== 'dist/main/index.js') errors.push('桌面入口与白名单不一致');
  if (manifestRequired) {
    const manifest = JSON.parse(read(MANIFEST));
    if (manifest.profile !== PROFILE || manifest.version !== pkg.version) errors.push('发布清单版本不一致');
    const expected = own.filter(n => n !== MANIFEST && n !== 'package.json').sort();
    if (JSON.stringify(Object.keys(manifest.files).sort()) !== JSON.stringify(expected)) errors.push('发布清单文件集合不一致');
    for (const name of expected) if (manifest.files[name] !== sha256(read(name))) errors.push(`内容与构建清单不一致: ${name}`);
  }
  return errors;
}

export function verifyStage({ allowPlainCli = false } = {}) {
  const names = walk(STAGE);
  const read = name => fs.readFileSync(path.join(STAGE, name));
  const errors = checkFiles(names, read, { allowPlainCli });
  const manifest = JSON.parse(read(MANIFEST));
  // 仅用于本机构建门禁，不是防攻击者篡改的数字签名。
  for (const [name, hash] of Object.entries(manifest.inputs)) {
    const file = path.join(ROOT, name);
    if (!fs.existsSync(file) || sha256(fs.readFileSync(file)) !== hash) errors.push(`发布副本已过期，请重建: ${name}`);
  }
  if (errors.length) throw new Error(errors.join('\n'));
  return manifest;
}
