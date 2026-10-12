import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { spawnSync } from 'node:child_process';
// @ts-ignore 只读门禁与命令行使用同一入口。
import { checkVersion } from '../scripts/check-version.mjs';

const sourceRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const files = [
  'package.json', 'package-lock.json', 'CHANGELOG.md', 'src/bridge/runtime.ts',
  'wps-addon/src/shared.js', 'office-addon/src/state.js', 'wps-addon/manifest.xml',
  'office-addon/excel/manifest.xml', 'src/renderer/App.tsx',
  'src/renderer/components/LiveWorkspaceCard.tsx', 'src/renderer/components/SettingsDrawer.tsx',
];

test('当前源码版本一致，检查不依赖生成物或历史包', () => {
  assert.deepEqual(checkVersion(sourceRoot).errors, []);
});

test('命令行拒绝漂移，beforePack 在改动共享副本前拒绝漂移', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-version-hook-'));
  const write = (name: string, value: string) => {
    const target = path.join(root, name);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, value);
  };
  try {
    for (const name of [...files, 'scripts/check-version.mjs', 'scripts/before-pack.cjs']) {
      write(name, fs.readFileSync(path.join(sourceRoot, name), 'utf8'));
    }
    write('scripts/lib/release-protection.mjs', `
      export const ROOT = ${JSON.stringify(root)};
      export const STAGE = ROOT + '/stage';
      export function verifyStage() { throw new Error('不应进入发布副本验证'); }
    `);
    write('stage/dist/bridge/cli-full.cjs', '原候选文件');
    write('dist/bridge/cli-full.cjs', '新候选文件');
    const lock = JSON.parse(fs.readFileSync(path.join(root, 'package-lock.json'), 'utf8'));
    lock.version = '999.999.999';
    write('package-lock.json', JSON.stringify(lock));
    const cli = spawnSync(process.execPath, [path.join(root, 'scripts/check-version.mjs')], { encoding: 'utf8' });
    assert.equal(cli.status, 1, cli.stderr);
    assert.match(cli.stderr, /package-lock.json version/);
    const { default: beforePack } = await import(pathToFileURL(path.join(root, 'scripts/before-pack.cjs')).href);
    for (const electronPlatformName of ['win32', 'darwin']) {
      await assert.rejects(beforePack({ electronPlatformName }), /package-lock.json version/);
      assert.equal(fs.readFileSync(path.join(root, 'stage/dist/bridge/cli-full.cjs'), 'utf8'), '原候选文件');
    }
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});

test('版本门禁拒绝各处漏改、声明缺失、重复和不支持的版本格式', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-version-'));
  const version = JSON.parse(fs.readFileSync(path.join(sourceRoot, 'package.json'), 'utf8')).version;
  const write = (name: string, content: string) => {
    const target = path.join(root, name);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, content);
  };
  const reset = () => {
    for (const name of files) write(name, fs.readFileSync(path.join(sourceRoot, name), 'utf8'));
  };
  try {
    reset();
    assert.deepEqual(checkVersion(root).errors, []);
    for (const name of files.slice(1)) {
      const original = fs.readFileSync(path.join(root, name), 'utf8');
      write(name, original.replaceAll(version, '999.999.999'));
      assert.ok(checkVersion(root).errors.some((error: string) => error.startsWith(name)), name);
      write(name, original);
    }
    const lock = JSON.parse(fs.readFileSync(path.join(root, 'package-lock.json'), 'utf8'));
    lock.packages[''].version = '999.999.999';
    write('package-lock.json', JSON.stringify(lock));
    assert.ok(checkVersion(root).errors.some((error: string) => error.includes('packages[""]')));
    reset();
    fs.unlinkSync(path.join(root, 'wps-addon/manifest.xml'));
    assert.ok(checkVersion(root).errors.some((error: string) => error.startsWith('wps-addon/manifest.xml')));
    reset();
    write('src/bridge/runtime.ts', '// 版本声明被移走');
    assert.ok(checkVersion(root).errors.some((error: string) => error.includes('匹配到 0 个')));
    write('src/bridge/runtime.ts', `export const VERSION = '${version}';\nexport const VERSION = '${version}';`);
    assert.ok(checkVersion(root).errors.some((error: string) => error.includes('匹配到 2 个')));
    reset();
    write('CHANGELOG.md', `## [${version}]\n## [${version}]\n`);
    assert.ok(checkVersion(root).errors.some((error: string) => error.startsWith('CHANGELOG.md')));
    for (const invalid of ['02.2.1', '2.2', '2.2.1-beta.1']) {
      write('package.json', JSON.stringify({ version: invalid }));
      assert.ok(checkVersion(root).errors.some((error: string) => error.startsWith('package.json')), invalid);
    }
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});
