/**
 * 加载项构建新鲜度与自动升级的回归测试。
 *
 * 覆盖的都是**曾经真实出问题**的点：
 *   - 版本号相同 ≠ 构建相同（`manifest.xml` 的 2.1.0 在多次构建间不变，
 *     只比版本号会把"装了旧构建"判成最新）；
 *   - 升级后不触发重载 → 磁盘新、进程旧（ISS-59）；
 *   - 部署失败还去重载 → 加载项仍是旧文件，制造"已修复"的假象。
 *
 * 测试用 `WPS_BRIDGE_RESOURCES` 与 `WPS_BRIDGE_ADDON_DIR` 指向临时目录，
 * **不触碰真实的 WPS 加载项目录**。
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const FP_A = 'a'.repeat(64);
const FP_B = 'b'.repeat(64);

function writeAddon(file: string, fingerprint: string | null) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const header = fingerprint
    ? `// 本文件由构建脚本生成\n// ADDON_BUILD_FINGERPRINT: ${fingerprint}\n(function () {\n  var ADDON_BUILD_FINGERPRINT = "${fingerprint}";\n})();\n`
    : `// 本文件由构建脚本生成（无指纹）\n(function () {}());\n`;
  fs.writeFileSync(file, header, 'utf8');
}

/** 每个用例独立的临时环境：包内资源目录 + 一个部署目录。 */
function makeEnv() {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'oab-freshness-'));
  const resources = path.join(root, 'resources');
  const addonDir = path.join(root, 'jsaddons');
  fs.mkdirSync(resources, { recursive: true });
  fs.mkdirSync(addonDir, { recursive: true });
  process.env.WPS_BRIDGE_RESOURCES = resources;
  process.env.WPS_BRIDGE_ADDON_DIR = addonDir;
  return {
    root,
    resources,
    addonDir,
    bundledFile: path.join(resources, 'wps-addon', 'addon-core.js'),
    deployedFile: path.join(addonDir, 'wps-bridge', 'addon-core.js')
  };
}

const { AddonInstaller } = await import('../src/main/addon-installer.js');
const { ensureAddonUpToDate } = await import('../src/main/addon-autoupgrade.js');

test('readAddonFingerprint 能从产物头部解析指纹，缺失时返回 null', () => {
  const env = makeEnv();
  writeAddon(env.bundledFile, FP_A);
  assert.equal(AddonInstaller.readAddonFingerprint(env.bundledFile), FP_A);

  const noFp = path.join(env.root, 'no-fp.js');
  writeAddon(noFp, null);
  assert.equal(AddonInstaller.readAddonFingerprint(noFp), null);

  assert.equal(AddonInstaller.readAddonFingerprint(path.join(env.root, '不存在.js')), null);
  assert.equal(AddonInstaller.readAddonFingerprint(null), null);
});

test('版本号相同但构建指纹不同 → 判定为需要升级', () => {
  const env = makeEnv();
  writeAddon(env.bundledFile, FP_A);
  writeAddon(env.deployedFile, FP_B); // 模拟"包更新了但部署的还是旧构建"

  const f = AddonInstaller.checkBuildFreshness();
  assert.equal(f.stale, true, '指纹不同必须判为过期——只比版本号会漏掉这种情况');
  assert.equal(f.bundledFingerprint, FP_A);
  assert.match(f.reason, /不一致/);
});

test('指纹一致时判定为最新', () => {
  const env = makeEnv();
  writeAddon(env.bundledFile, FP_A);
  writeAddon(env.deployedFile, FP_A);
  const f = AddonInstaller.checkBuildFreshness();
  assert.equal(f.stale, false);
  assert.match(f.reason, /一致/);
});

test('没有已部署副本 → 需要安装', () => {
  const env = makeEnv();
  writeAddon(env.bundledFile, FP_A);
  const f = AddonInstaller.checkBuildFreshness();
  assert.equal(f.stale, true);
  assert.match(f.reason, /未找到已部署/);
});

test('读不到包内指纹时返回 null（判不了），不猜测', () => {
  const env = makeEnv();
  writeAddon(env.bundledFile, null); // 产物没有注入指纹
  writeAddon(env.deployedFile, FP_A);
  const f = AddonInstaller.checkBuildFreshness();
  assert.equal(f.stale, null, '判不了就必须返回 null，不能默认"需要升级"或"已最新"');
});

test('自动升级：过期时重新部署并触发重载', async () => {
  const env = makeEnv();
  writeAddon(env.bundledFile, FP_A);
  writeAddon(env.deployedFile, FP_B);

  let installed = 0;
  let reloaded = 0;
  const r = await ensureAddonUpToDate({
    installFn: async () => { installed++; writeAddon(env.deployedFile, FP_A); return { success: true }; },
    reloadFn: async () => { reloaded++; }
  });

  assert.equal(installed, 1);
  assert.equal(reloaded, 1, '部署后必须触发重载，否则磁盘新、进程旧');
  assert.equal(r.upgraded, true);
  assert.equal(r.reloaded, true);
});

test('自动升级：已最新时不重复部署', async () => {
  const env = makeEnv();
  writeAddon(env.bundledFile, FP_A);
  writeAddon(env.deployedFile, FP_A);

  let installed = 0;
  const r = await ensureAddonUpToDate({
    installFn: async () => { installed++; return { success: true }; },
    reloadFn: async () => {}
  });
  assert.equal(installed, 0, '已一致时不应产生部署副作用');
  assert.equal(r.upgraded, false);
});

test('自动升级：判不了（无包内指纹）时不动部署', async () => {
  const env = makeEnv();
  writeAddon(env.bundledFile, null);
  writeAddon(env.deployedFile, FP_B);

  let installed = 0;
  const r = await ensureAddonUpToDate({
    installFn: async () => { installed++; return { success: true }; },
    reloadFn: async () => {}
  });
  assert.equal(installed, 0, '判不了时宁可不动，也不要盲目重新部署');
  assert.equal(r.upgraded, false);
});

test('自动升级：部署失败时不触发重载（避免制造已修复的假象）', async () => {
  const env = makeEnv();
  writeAddon(env.bundledFile, FP_A);
  writeAddon(env.deployedFile, FP_B);

  let reloaded = 0;
  const r = await ensureAddonUpToDate({
    installFn: async () => ({ success: false, message: '权限不足' }),
    reloadFn: async () => { reloaded++; }
  });
  assert.equal(reloaded, 0, '部署没成功就重载，只会再次加载旧文件');
  assert.equal(r.upgraded, false);
  assert.match(r.message, /未成功/);
});
