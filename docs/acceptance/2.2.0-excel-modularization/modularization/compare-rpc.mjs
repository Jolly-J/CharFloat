// 用同一套内存宿主夹具比较固定基线与当前压缩生成物；只在临时目录恢复旧产物。
import fs from 'node:fs';import path from 'node:path';import os from 'node:os';import assert from 'node:assert/strict';
import {execFileSync} from 'node:child_process';import {fileURLToPath} from 'node:url';
const root=fileURLToPath(new URL('../../../../',import.meta.url));const dir=path.dirname(fileURLToPath(import.meta.url));
const baseline=JSON.parse(fs.readFileSync(path.join(dir,'baseline.json')));const tmp=fs.mkdtempSync(path.join(os.tmpdir(),'charfloat-rpc-'));
try{
 const before=path.join(tmp,'baseline.js');fs.writeFileSync(before,execFileSync('git',['show',`${baseline.commit}:wps-addon/addon-core.js`],{cwd:root,maxBuffer:2**22}));
 const timings={};
 for(const [name,artifact] of [['before',before],['after',path.join(root,'wps-addon/addon-core.js')]]){
  const start=performance.now();const output=execFileSync(process.execPath,['--import','tsx','--test','tests/wps-excel-rpc.test.ts'],{cwd:root,encoding:'utf8',timeout:30000,env:{...process.env,OAB_TEST_WPS_ARTIFACT:artifact,CHARFLOAT_EXCEL_SNAPSHOT:path.join(dir,`rpc-${name}.json`)}});
  timings[name]=Math.round(performance.now()-start);fs.writeFileSync(path.join(dir,`rpc-${name}.log`),output);
 }
 const expected=JSON.parse(fs.readFileSync(path.join(dir,'rpc-before.json'))),actual=JSON.parse(fs.readFileSync(path.join(dir,'rpc-after.json')));
 assert.deepEqual(actual,expected,'RPC 响应、写入值和宿主调用记录不得变化');
 const result={cases:Object.keys(expected).length,responsesAndHostCallsUnchanged:true,testProcessElapsedMs:timings,timingScope:'整套模拟测试进程耗时，含启动与编译；不是宿主性能基准',excludedFields:['运行时构建指纹仅单独验证一致性；事件用例比对注册次数'],host:'内存模拟，无真实宿主操作'};
 fs.writeFileSync(path.join(dir,'rpc-equivalence.json'),JSON.stringify(result,null,2)+'\n');console.log(JSON.stringify(result,null,2));
}finally{fs.rmSync(tmp,{recursive:true,force:true})}
