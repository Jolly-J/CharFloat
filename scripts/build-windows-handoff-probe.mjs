// Build a standalone offline diagnostic only; does not launch or update an app.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { build } from 'esbuild';
import { assertPublicPaths } from './lib/update-key-safety.mjs';
const root = fileURLToPath(new URL('../', import.meta.url));
const [outArg, configArg] = process.argv.slice(2);
if (!outArg || !configArg) throw new Error('需要全新输出目录与已有验收配置 JSON。');
const output = path.resolve(outArg); assertPublicPaths([output, path.resolve(configArg)]);
const stat = fs.lstatSync(configArg);
if (!stat.isFile() || stat.isSymbolicLink() || stat.nlink !== 1 || stat.size > 128 * 1024) throw new Error('验收配置文件类型无效。');
let input;
try { input = JSON.parse(fs.readFileSync(configArg, 'utf8')); } catch { throw new Error('验收配置不是有效 JSON，未输出其内容。'); }
const config = { version: input.version, installRootName: input.installRootName, runtimeDirectoryName: input.runtimeDirectoryName, runtimeSha256: input.runtimeSha256 };
if (!/^[A-Za-z0-9-]{1,80}$/.test(config.installRootName || '') || config.installRootName === 'CharFloat' || !/^[A-Za-z0-9-]{1,80}$/.test(config.runtimeDirectoryName || '') || !/^\d+\.\d+\.\d+$/.test(config.version || '') || !/^[a-f0-9]{64}$/.test(config.runtimeSha256 || '') || fs.existsSync(output)) throw new Error('诊断配置无效或输出已存在。');
fs.mkdirSync(output, { recursive: true });
await build({ entryPoints: [path.join(root, 'scripts/windows-update-handoff-probe.ts')], bundle: true, platform: 'node', format: 'cjs', target: 'node22', outfile: path.join(output, 'windows-update-handoff-probe.cjs') });
fs.writeFileSync(path.join(output, 'handoff-config.json'), JSON.stringify(config, null, 2) + '\n');
const cmd = `@echo off
setlocal
chcp 65001 >nul
set "CHARFLOAT_TEST_EXE="
for /d %%D in ("%LOCALAPPDATA%\\Programs\\${config.installRootName}\\${config.version}-*") do if exist "%%~fD\\字浮 CharFloat.exe" set "CHARFLOAT_TEST_EXE=%%~fD\\字浮 CharFloat.exe"
if not defined CHARFLOAT_TEST_EXE goto missing
echo 正在检查更新助手，仅取证，不安装。请等待约一分钟。
set "ELECTRON_RUN_AS_NODE=1"
"%CHARFLOAT_TEST_EXE%" "%~dp0windows-update-handoff-probe.cjs" > "%~dp0handoff-console.log" 2>&1
echo 检查结束，请把本目录 handoff-console.log 发回。日志内已包含完整脱敏报告。
pause
exit /b
:missing
echo 未找到本轮 A，请保留现有程序并反馈。
pause
exit /b 2
`;
fs.writeFileSync(path.join(output, '更新交接取证.cmd'), cmd.replaceAll('\n', '\r\n'));
console.log(JSON.stringify({ files: 3, output, bundleSha256: createHash('sha256').update(fs.readFileSync(path.join(output, 'windows-update-handoff-probe.cjs'))).digest('hex'), installs: false, grantsAuthorization: false }));
