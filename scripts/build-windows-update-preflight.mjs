// Build standalone diagnostic files only; never installs, launches or uploads.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { build } from 'esbuild';
import { assertPublicPaths } from './lib/update-key-safety.mjs';

const root = fileURLToPath(new URL('../', import.meta.url));
const [outArg, installRootName, version, exeArg] = process.argv.slice(2);
if (!outArg || !/^[A-Za-z0-9-]{1,80}$/.test(installRootName || '') || installRootName === 'CharFloat' || !/^\d+\.\d+\.\d+$/.test(version || '') || !exeArg) throw new Error('用法：node scripts/build-windows-update-preflight.mjs <全新目录> <独立Windows安装根名称> <A版本> <本次A主程序EXE>');
const output = path.resolve(outArg), executable = path.resolve(exeArg); assertPublicPaths([output, executable]);
if (fs.existsSync(output)) throw new Error('输出目录已存在，未覆盖。');
const header = fs.readFileSync(executable);
if (header.toString('ascii', 0, 2) !== 'MZ') throw new Error('需要本次 Windows 程序运行时。');
fs.mkdirSync(output, { recursive: true });
await build({ entryPoints: [path.join(root, 'scripts/windows-update-preflight.ts')], bundle: true, platform: 'node', format: 'cjs', target: 'node22', outfile: path.join(output, 'windows-update-preflight.cjs') });
const inputs = ['scripts/windows-update-preflight.ts', 'src/main/update-install.ts', 'src/main/update-protocol.ts', 'windows-bootstrap/update.mjs', 'windows-bootstrap/core.mjs'];
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
fs.writeFileSync(path.join(output, 'preflight-config.json'), JSON.stringify({ version, installRootName, runtimeSha256: sha(header), sourceInputs: Object.fromEntries(inputs.map(file => [file, sha(fs.readFileSync(path.join(root, file)))])), network: false, updatesInstalled: false }, null, 2) + '\n');
const command = `@echo off
setlocal
chcp 65001 >nul
set "CHARFLOAT_TEST_EXE="
for /d %%D in ("%LOCALAPPDATA%\\Programs\\${installRootName}\\${version}-*") do if exist "%%~fD\\字浮 CharFloat.exe" set "CHARFLOAT_TEST_EXE=%%~fD\\字浮 CharFloat.exe"
if not defined CHARFLOAT_TEST_EXE goto missing
set "ELECTRON_RUN_AS_NODE=1"
"%CHARFLOAT_TEST_EXE%" "%~dp0windows-update-preflight.cjs" > "%~dp0preflight-console.log" 2>&1
set "CHARFLOAT_TEST_RESULT=%errorlevel%"
if "%CHARFLOAT_TEST_RESULT%"=="0" (echo 检查通过，请把本目录 preflight-数字.json 发回。) else (echo 检查未通过，请保留报告，暂不点击更新。)
pause
exit /b %CHARFLOAT_TEST_RESULT%
:missing
echo 尚未找到本轮 A 版，请先安装 A，再运行此检查。
pause
exit /b 2
`;
fs.writeFileSync(path.join(output, '验收前检查.cmd'), command.replaceAll('\n', '\r\n'));
console.log(JSON.stringify({ output, files: 3, scope: 'offline security preflight only', runtimeSha256: sha(header) }));
