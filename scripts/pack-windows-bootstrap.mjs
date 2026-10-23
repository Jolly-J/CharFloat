// 对真实 NSIS 编译参数做断言，再生成便携自解压 EXE；不签名、不发布、不执行 Windows 应用。
import fs from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import { build, Platform, Arch } from 'electron-builder';
import { NsisTarget } from 'app-builder-lib/out/targets/nsis/NsisTarget.js';
import { assertPortableDefines, installPortableCompileHook, PORTABLE_TEMPLATE_SHA256 } from './lib/windows-portable.mjs';
const root=path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const config=JSON.parse(fs.readFileSync(path.join(root,'dist/windows-bootstrap/builder.json'),'utf8'));
const require = createRequire(import.meta.url);
const builderRoot = path.dirname(require.resolve('app-builder-lib/package.json'));
const version = JSON.parse(fs.readFileSync(path.join(builderRoot,'package.json'),'utf8')).version;
const template = fs.readFileSync(path.join(builderRoot,'templates/nsis/portable.nsi'),'utf8');
let checked=false, compiled=false;
const restore = installPortableCompileHook(NsisTarget, { template, version, onCompiled: script => {
  compiled=true;
  fs.writeFileSync(path.join(root,'dist/windows-bootstrap/nsis-check.json'),JSON.stringify({requestExecutionLevel:'user',uniquePerLaunchTemp:true,splash:false,explicitInitBeforeSilent:true,unpackDirNameDefined:false,templateSha256:PORTABLE_TEMPLATE_SHA256,compilerInputSha256:createHash('sha256').update(script).digest('hex')},null,2)+'\n');
  fs.writeFileSync(path.join(root,'dist/windows-bootstrap/nsis-input.nsi'),script);
} });
try { await build({ projectDir:root,targets:Platform.WINDOWS.createTarget('portable',Arch.x64),config:path.join(root,'dist/windows-bootstrap/builder.json'),publish:'never',
  effectiveOptionComputed: async ([defines]) => {
    assertPortableDefines(defines);
    checked=true;
    return false;
  } }); } finally { restore(); }
assert.ok(checked,'没有实际执行 NSIS 参数检查');
assert.ok(compiled,'没有实际编译已核查的静默外壳脚本');
