// 验证代码保护真的生效——**每次打包后跑一遍**。
//
// 检查四件事：
//   ① 包内没有 .map、没有内嵌的原始 .ts 源码
//   ② 明文包 cli-full.cjs 没有被打进包
//   ③ 核心 cli 走的是**字节码**（把明文包临时移走仍能跑，才证明不是悄悄回退）
//   ④ 功能仍正常（MCP 握手 + 工具清单）
//
// 用法：node scripts/check-cli-protection.mjs [app.asar 或 .app 路径]
//   不传路径时自动找 release/<版本>/mac/*.app
import { execFileSync, spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const root = process.cwd();
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const version = pkg.version;

let pass = 0, fail = 0;
const ck = (name, ok, detail = '') => {
  if (ok) { pass++; console.log(`  ✔ ${name}`); }
  else { fail++; console.log(`  ✖ ${name}${detail ? '   → ' + detail : ''}`); }
};

// ── 定位 app.asar
function findAsar() {
  const arg = process.argv[2];
  const candidates = [];
  if (arg) candidates.push(arg);
  const macDir = path.join(root, 'release', version, 'mac');
  if (fs.existsSync(macDir)) {
    for (const d of fs.readdirSync(macDir)) {
      candidates.push(path.join(macDir, d, 'Office Agent Bridge.app', 'Contents/Resources/app.asar'));
    }
  }
  candidates.push(path.join(root, 'release', version, 'win/win-unpacked/resources/app.asar'));
  return candidates.filter((c) => c && fs.existsSync(c));
}

console.log('\n═══ 代码保护校验 ═══\n');

// ── 本地产物自检（不需要包）
console.log('【本地产物】');
const shim = path.join(root, 'dist/bridge/cli.cjs');
const jsc = path.join(root, 'dist/bridge/cli.jsc');
ck('cli.jsc 存在（字节码已编译）', fs.existsSync(jsc));
if (fs.existsSync(shim)) {
  const size = fs.statSync(shim).size;
  ck('cli.cjs 是加载器而非明文包', size < 4096, `实际 ${size} 字节（明文包约 1.3 MB）`);
} else ck('cli.cjs 存在', false);

const asars = findAsar();
if (asars.length === 0) {
  console.log('\n  ⊘ 没找到打包产物，跳过包内检查（先跑 npm run dist / electron-builder）');
}
// 两个平台的包都要查 —— 只查 mac 会漏掉 Windows 包的差异
for (const asar of asars) {
  console.log(`\n【包内检查】${path.relative(root, asar)}`);
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'prot-'));
  try {
    execFileSync('npx', ['asar', 'extract', asar, tmp], { stdio: 'ignore' });
    const walk = (dir, out = []) => {
      for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
        const p = path.join(dir, e.name);
        if (e.isDirectory()) walk(p, out); else out.push(p);
      }
      return out;
    };
    const files = walk(tmp);
    const maps = files.filter((f) => f.endsWith('.map'));
    ck('包内没有 .map', maps.length === 0, `仍有 ${maps.length} 个`);

    const tsFiles = files.filter((f) => f.endsWith('.ts') && !f.includes('node_modules'));
    ck('包内没有原始 .ts 源码', tsFiles.length === 0, `仍有 ${tsFiles.length} 个`);

    const leaked = files.filter((f) => {
      if (!f.endsWith('.js') && !f.endsWith('.cjs')) return false;
      try { return fs.readFileSync(f, 'utf8').includes('mcpServers = { ...config.mcpServers'); }
      catch { return false; }
    });
    ck('包内没有可读的核心源码片段', leaked.length === 0, leaked.slice(0, 3).join(', '));

    const full = files.filter((f) => f.endsWith('cli-full.cjs'));
    ck('明文包 cli-full.cjs 未打进包', full.length === 0, full.join(', '));

    ck('包内有字节码 cli.jsc', files.some((f) => f.endsWith('cli.jsc')));

  // ⚠️ 只查 asar 内容**不够**：安装器把入口重定向到 `app.asar.unpacked`，
  // 那里必须同时有 cli.cjs 与 cli.jsc，且 bytenode 可解析。
  // 这条是"打包后打不开"的真实根因（2026-09-23 由 Windows 测试暴露，mac 同样受影响）。
  const unpacked = asar.replace(/app\.asar$/, 'app.asar.unpacked');
  if (fs.existsSync(unpacked)) {
    const uShim = path.join(unpacked, 'dist/bridge/cli.cjs');
    const uJsc = path.join(unpacked, 'dist/bridge/cli.jsc');
    ck('asar.unpacked 里有入口 cli.cjs', fs.existsSync(uShim));
    ck('asar.unpacked 里有字节码 cli.jsc', fs.existsSync(uJsc), '入口重定向到 unpacked，缺它则加载器找不到字节码');
    // 真正跑一次：以安装器实际使用的路径启动
    const electron = process.platform === 'darwin'
      ? path.join(root, 'node_modules/electron/dist/Electron.app/Contents/MacOS/Electron')
      : path.join(root, 'node_modules/electron/dist/electron');
    if (fs.existsSync(uShim) && fs.existsSync(electron)) {
      try {
        const out = execFileSync(electron, [uShim, '--status'], {
          env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' },
          encoding: 'utf8', timeout: 20000,
        });
        ck('以安装器路径启动可用（真实场景）', out.includes('"service"'), out.trim().slice(0, 120));
      } catch (e) {
        ck('以安装器路径启动可用（真实场景）', false, String(e.message).split('\n')[0].slice(0, 120));
      }
    }
  }

    // ── 关键一条：把明文包移走后，字节码仍能跑 → 证明真的在用字节码
    const localFull = path.join(root, 'dist/bridge/cli-full.cjs');
    const hold = path.join(os.tmpdir(), `cli-full-hold-${Date.now()}.cjs`);
    let moved = false;
    if (fs.existsSync(localFull)) { fs.renameSync(localFull, hold); moved = true; }
    try {
      const electron = process.platform === 'darwin'
        ? path.join(root, 'node_modules/electron/dist/Electron.app/Contents/MacOS/Electron')
        : process.platform === 'win32'
          ? path.join(root, 'node_modules/electron/dist/electron.exe')
          : path.join(root, 'node_modules/electron/dist/electron');
      if (!fs.existsSync(electron)) {
        console.log('  ⊘ 找不到 Electron 二进制，跳过"字节码实跑"检查');
      } else {
        const tools = execFileSync(electron, [path.join(root, 'scripts/probe-mcp-tools.mjs')], {
          env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }, encoding: 'utf8', timeout: 30000,
        }).trim();
        ck('明文包移走后仍能跑（确认真走字节码）', Number(tools) > 100, `读到 ${tools} 个工具`);
      }
    } catch (e) {
      ck('明文包移走后仍能跑（确认真走字节码）', false, String(e.message).slice(0, 120));
    } finally {
      if (moved) fs.renameSync(hold, localFull);
    }
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

console.log(`\n  合计 ${pass} 通过 / ${fail} 失败\n`);
process.exit(fail === 0 ? 0 : 1);
