// Build full application ZIPs into a NEW local directory. Never signs manifests or uploads anything.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import { checkDependencies } from './check-dependencies.mjs';
import { checkRelease } from './lib/check-release.mjs';
import { assertPublicPaths } from './lib/update-key-safety.mjs';
const root = fileURLToPath(new URL('../', import.meta.url));
const [outputArg, platform = 'all'] = process.argv.slice(2);
if (!outputArg || !['all', 'mac', 'win'].includes(platform)) throw new Error('用法：node scripts/build-update-packages.mjs <全新输出目录> [all|mac|win]');
const output = path.resolve(outputArg); assertPublicPaths([output]);
if (fs.existsSync(output)) throw new Error('输出目录已存在，未覆盖。');
if (platform !== 'win' && process.platform !== 'darwin') throw new Error('Mac 更新 ZIP 必须在 macOS 构建。');
const dependencies = checkDependencies(); if (dependencies.errors.length) throw new Error(dependencies.errors.join('\n'));
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const run = (command, args, env = {}) => execFileSync(command, args, { cwd: root, stdio: 'inherit', shell: process.platform === 'win32' && command === 'npm.cmd', env: { ...process.env, ...env } });
const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm';
fs.mkdirSync(output, { recursive: true });
run(npm, ['run', 'typecheck']);
run(npm, ['run', 'check:agents']);
run(npm, ['run', 'build']);
run(npm, ['test']);
run(npm, ['run', 'build:release-assets']);
run(npm, ['run', 'build:win-icon']);
const results = [];
for (const target of platform === 'all' ? ['mac', 'win'] : [platform]) {
  // Serial: beforePack intentionally changes the platform-specific CLI fallback in the release copy.
  const destination = path.join(output, target);
  run(process.execPath, [path.join(root, 'node_modules/electron-builder/out/cli/cli.js'), `--${target}`, target === 'mac' ? '--arm64' : '--x64', '--publish', 'never', `-c.directories.output=${destination}`, ...(target === 'mac' ? ['-c.mac.identity=null'] : [])], { CSC_IDENTITY_AUTO_DISCOVERY: 'false' });
  const zipNames = fs.readdirSync(destination).filter(name => name.endsWith('.zip'));
  if (zipNames.length !== 1) throw new Error('每个平台必须只生成一个完整程序 ZIP。');
  const file = path.join(destination, zipNames[0]);
  const checked = await checkRelease(file, { allowPlainCli: target === 'win' });
  run(process.execPath, ['--import', 'tsx', 'scripts/check-update-package.ts', target === 'mac' ? 'darwin' : 'win32', pkg.version, file]);
  run(process.execPath, ['scripts/check-update-secrets.mjs', file]);
  results.push({ platform: target === 'mac' ? 'darwin' : 'win32', arch: target === 'mac' ? 'arm64' : 'x64', kind: target === 'mac' ? 'mac-zip' : 'windows-zip', file: path.relative(output, file), size: fs.statSync(file).size, sha256: checked.sha256 });
}
fs.writeFileSync(path.join(output, 'packages.json'), JSON.stringify({ version: pkg.version, dependenciesChecked: dependencies.checked, assets: results, hostAcceptance: 'not performed by build script', uploaded: false }, null, 2) + '\n', { flag: 'wx' });
console.log(`本地程序 ZIP 已生成：${output}；平台实机更新仍须分别验收。`);
