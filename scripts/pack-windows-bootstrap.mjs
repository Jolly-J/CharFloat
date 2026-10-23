// 对真实 NSIS 编译参数做断言，再生成便携自解压 EXE；不签名、不发布、不执行 Windows 应用。
import fs from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import { build, Platform, Arch } from 'electron-builder';
const root=path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const config=JSON.parse(fs.readFileSync(path.join(root,'dist/windows-bootstrap/builder.json'),'utf8'));
let checked=false;
await build({ projectDir:root,targets:Platform.WINDOWS.createTarget('portable',Arch.x64),config:path.join(root,'dist/windows-bootstrap/builder.json'),publish:'never',
  effectiveOptionComputed: async ([defines]) => {
    assert.equal(defines.REQUEST_EXECUTION_LEVEL,'user','外壳不得请求管理员权限');
    assert.ok(!('UNPACK_DIR_NAME' in defines),'必须使用每次启动独立的临时目录');
    assert.ok(defines.SPLASH_IMAGE,'须显式初始化插件临时目录');
    checked=true;
    fs.writeFileSync(path.join(root,'dist/windows-bootstrap/nsis-check.json'),JSON.stringify({requestExecutionLevel:defines.REQUEST_EXECUTION_LEVEL,uniquePerLaunchTemp:true,splash:true,unpackDirNameDefined:false},null,2)+'\n');
    return false;
  } });
assert.ok(checked,'没有实际执行 NSIS 参数检查');
