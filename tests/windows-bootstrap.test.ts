import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { install, fileHash, safeName, validateManifest, verifyInstallation, PRODUCT, EXE } from '../windows-bootstrap/core.mjs';
import { createShortcut } from '../windows-bootstrap/shortcut.mjs';
import { runStages } from '../windows-bootstrap/stages.mjs';
import { launchApplication } from '../windows-bootstrap/launch.mjs';
import { EventEmitter } from 'node:events';
import vm from 'node:vm';
import { createRequire } from 'node:module';
import { build } from 'esbuild';
import { deflateRawSync } from 'node:zlib';
import { createProgressSender } from '../windows-bootstrap/progress.mjs';
import { createDiagnostics } from '../windows-bootstrap/diagnostics.mjs';
import physicalFs from '../windows-bootstrap/physical-fs.mjs';
import { preparePortableScript, assertPortableDefines, installPortableCompileHook } from '../scripts/lib/windows-portable.mjs';
import { spawn } from 'node:child_process';

const sha = (b: Buffer) => createHash('sha256').update(b).digest('hex');
const builderRequire=createRequire(import.meta.url);
const portableTemplate=fs.readFileSync(path.join(path.dirname(builderRequire.resolve('app-builder-lib/package.json')),'templates/nsis/portable.nsi'),'utf8');
test('静默portable在onInit先建独立插件目录；载荷提取与ExecWait保持原模板',()=>{
  const script=preparePortableScript('header\n'+portableTemplate,portableTemplate,'25.1.8');
  const init=script.slice(script.indexOf('Function .onInit'),script.indexOf('FunctionEnd',script.indexOf('Function .onInit')));
  assert.ok(init.indexOf('InitPluginsDir')<init.indexOf('SetSilent silent'));
  assert.equal((script.match(/^  InitPluginsDir$/gm)||[]).length,1);
  assert.equal(script.slice(script.indexOf('\nSection\n')),portableTemplate.slice(portableTemplate.indexOf('\nSection\n')));
  assertPortableDefines({REQUEST_EXECUTION_LEVEL:'user'});
  for(const d of [{REQUEST_EXECUTION_LEVEL:'admin'},{REQUEST_EXECUTION_LEVEL:'user',SPLASH_IMAGE:'old.bmp'},{REQUEST_EXECUTION_LEVEL:'user',UNPACK_DIR_NAME:'fixed'}])assert.throws(()=>assertPortableDefines(d));
  assert.throws(()=>preparePortableScript(portableTemplate,portableTemplate,'26.0.0'));
  assert.throws(()=>preparePortableScript(portableTemplate+'changed',portableTemplate,'25.1.8'));
  assert.throws(()=>preparePortableScript(portableTemplate,portableTemplate+'changed','25.1.8'));
});
test('实际编译适配器传递已初始化脚本，成功/失败均可恢复原方法',async()=>{
  let received='';class Target{isPortable=true;async executeMakensis(_d:any,_c:any,s:string){received=s;}}
  const original=Target.prototype.executeMakensis;let compiled=0;
  assert.throws(()=>installPortableCompileHook(Target,{template:portableTemplate,version:'26.0.0'}));assert.equal(Target.prototype.executeMakensis,original);
  const restore=installPortableCompileHook(Target,{template:portableTemplate,version:'25.1.8',onCompiled:()=>compiled++});
  try{await new Target().executeMakensis({REQUEST_EXECUTION_LEVEL:'user'},{},portableTemplate);}finally{restore();}
  assert.match(received,/Function \.onInit\n[^]*?InitPluginsDir[^]*?SetSilent silent/);assert.equal(compiled,1);assert.equal(Target.prototype.executeMakensis,original);
  const restoreFailure=installPortableCompileHook(Target,{template:portableTemplate,version:'25.1.8'});
  try{await assert.rejects(new Target().executeMakensis({REQUEST_EXECUTION_LEVEL:'admin'},{},portableTemplate));}finally{restoreFailure();}
  assert.equal(Target.prototype.executeMakensis,original);
});
test('引导界面移除边角文字和按钮符号，关闭按钮保留可访问SVG',()=>{
  const html=fs.readFileSync('windows-bootstrap/index.html','utf8'),code=fs.readFileSync('windows-bootstrap/renderer.js','utf8');
  assert.match(html,/<button id="close" aria-label="关闭"><svg[^>]*aria-hidden="true"/);assert.match(html,/<button id="launch" hidden disabled>开始使用<\/button>/);
  assert.doesNotMatch(html,/<footer|id="step"|↗|给你的 AI/);assert.doesNotMatch(code,/byId\('step'\)/);
});
test('Windows真实Node写句柄与PowerShell共享快照、UTF8半行、去重和暂态恢复', {skip:process.platform!=='win32'}, async t=>{
  const f=await fixture(t), log=path.join(f.dir,'shared.jsonl'), ready=path.join(f.dir,'reader-ready'),gate=path.join(f.dir,'continue'),result=path.join(f.dir,'snapshot.json'),runId='12345678-1234-1234-1234-123456789abc';
  const fd=fs.openSync(log,'wx');let child:any, timer:any;
  try {
  const second=Buffer.from(JSON.stringify({runId,sequence:2,message:'合成内容'})), split=second.indexOf(Buffer.from('合'))+1;
  fs.writeSync(fd,JSON.stringify({runId,sequence:1,event:'first'})+'\n');fs.writeSync(fd,second.subarray(0,split));
  child=spawn('powershell.exe',['-NoLogo','-NoProfile','-File',path.resolve('tests/windows-bootstrap-sharing.ps1'),log,runId,ready,gate,result],{stdio:['ignore','pipe','pipe']});
  let stderr='';child.stderr.on('data',b=>stderr=(stderr+b).slice(-16384));
  const exited=new Promise<number|null>((resolve,reject)=>{child.once('error',reject);child.once('close',resolve);});
  timer=setTimeout(()=>child.kill(),15000);
  const deadline=Date.now()+10000;while(!fs.existsSync(ready)){assert.ok(Date.now()<deadline&&child.exitCode===null,stderr||'PowerShell reader did not start');await new Promise(r=>setTimeout(r,20));}
  fs.writeSync(fd,second.subarray(split));fs.writeSync(fd,'\n'+second.toString()+'\n');fs.writeFileSync(gate,'continue');
  const code=await exited;assert.equal(code,0,stderr);const data=JSON.parse(fs.readFileSync(result,'utf8'));
  assert.equal(data.legacyBlocked,true);assert.equal(data.firstCount,1);assert.equal(data.firstIncomplete,true);assert.equal(data.secondCount,2);assert.equal(data.secondIncomplete,false);assert.equal(data.failurePreserved,true);assert.equal(data.recovered,1);assert.equal(data.warningCleared,true);assert.equal(data.message,'合成内容');
  } finally { clearTimeout(timer);if(child&&child.exitCode===null)child.kill();fs.closeSync(fd); }
});
test('普通 Node 使用原生 fs；诊断默认关闭，显式日志白名单/脱敏/大小上限与终态落盘',async t=>{
  assert.equal(physicalFs,fs);assert.equal(createDiagnostics({env:{}}).enabled,false);
  const f=await fixture(t), runId='12345678-1234-1234-1234-123456789abc', local=await fs.promises.realpath(f.local), dir=path.join(local,'CharFloatStartupDiagnostics',runId);
  fs.mkdirSync(dir,{recursive:true});const env={LOCALAPPDATA:local,USERPROFILE:await fs.promises.realpath(f.dir),CHARFLOAT_DIAGNOSTIC_RUN_ID:runId,CHARFLOAT_DIAGNOSTIC_DIR:dir};
  const log=createDiagnostics({env,maximumBytes:1600});assert.equal(log.enabled,true);
  log.record('error',{message:`${env.USERPROFILE.toUpperCase()}/temp token=secret-value Bearer test.secret`,stack:'secret-stack',password:'secret-password',bytes:NaN});
  for(let i=0;i<100;i++)log.record('progress',{file:'resources/app.asar',bytes:i});log.flush();log.close({exitCode:1});
  const text=fs.readFileSync(path.join(dir,'bootstrap.jsonl'),'utf8');assert.ok(Buffer.byteLength(text)<=1600);assert.ok(text.includes('<user-path>'));assert.equal(/secret-value|test.secret|secret-stack|secret-password/.test(text),false);
  const first=JSON.parse(text.split('\n')[0]);assert.equal('bytes' in first,false);assert.equal(first.runId,runId);
  const final=JSON.parse(fs.readFileSync(path.join(dir,'bootstrap-final.json'),'utf8'));assert.equal(final.exitCode,1);assert.equal(final.logCapped,true);
  assert.equal(createDiagnostics({env:{...env,CHARFLOAT_DIAGNOSTIC_DIR:f.dir}}).enabled,false);
});
test('诊断用物理目录归属匹配等价路径表示，未知目录不会静默假称已采集',async t=>{
  const f=await fixture(t), local=await fs.promises.realpath(f.local), alias=path.join(f.dir,'alias-local'), runId='87654321-1234-1234-1234-123456789abc';
  fs.symlinkSync(local,alias,'junction');const dir=path.join(local,'CharFloatStartupDiagnostics',runId);fs.mkdirSync(dir,{recursive:true});
  const log=createDiagnostics({env:{LOCALAPPDATA:alias,CHARFLOAT_DIAGNOSTIC_RUN_ID:runId,CHARFLOAT_DIAGNOSTIC_DIR:dir}});
  assert.equal(log.enabled,true);log.record('main-start');log.close();assert.ok(fs.existsSync(path.join(dir,'bootstrap-final.json')));
  const errors:string[]=[];t.mock.method(console,'error',(value:string)=>errors.push(value));
  const disabled=createDiagnostics({env:{LOCALAPPDATA:local,CHARFLOAT_DIAGNOSTIC_RUN_ID:runId,CHARFLOAT_DIAGNOSTIC_DIR:f.dir}});
  assert.equal(disabled.enabled,false);assert.match(disabled.reason,/directory/);assert.match(errors[0],/diagnostics unavailable/);
});
test('离线启动入口保留执行策略，先校验再启动，读取日志有上限，失败也保存报告',()=>{
  const cmd=fs.readFileSync('windows-bootstrap/Start-Local-Debug.cmd','utf8'), ps=fs.readFileSync('windows-bootstrap/Start-Local-Debug.ps1','utf8');
  assert.match(cmd,/-NoProfile -File/);assert.doesNotMatch(cmd+ps,/ExecutionPolicy\s+Bypass|Set-ExecutionPolicy|SetEnvironmentVariable|Set-ItemProperty|runAs\b/i);
  assert.ok(ps.indexOf('Get-FileHash')<ps.indexOf('$child.Start()'));assert.match(ps,/ReadAsync/);assert.match(ps,/65536/);assert.match(ps,/finally[\s\S]*Save-Report/);
  assert.match(ps,/EnvironmentVariables\['CHARFLOAT_DIAGNOSTIC_RUN_ID'\]/);assert.match(ps,/Assert-SafeDirectory \$runDir/);
});
test('跨小文件全局合并进度，保留最新值与阶段切换；终态后没有延迟工作状态',()=>{
  let time=0,id=0;const timers=new Map<number,()=>void>(),sent:any[]=[];
  const progress=createProgressSender((p:any)=>sent.push(p),{now:()=>time,setTimer:(fn:any)=>{timers.set(++id,fn);return id;},clearTimer:(token:number)=>timers.delete(token)});
  progress.push({phase:'verify',bytes:0});
  for(let i=0;i<1404;i++)progress.push({phase:'extract',file:`files/${i}`,bytes:i});
  assert.equal(sent.length,2);assert.equal(timers.size,1);
  time=200;const tick=[...timers.values()][0];timers.clear();tick();
  assert.equal(sent.length,3);assert.equal(sent.at(-1).bytes,1403);assert.equal(sent.at(-1).file,'files/1403');
  progress.push({phase:'extract',bytes:1404});const stale=[...timers.values()][0];
  progress.push({phase:'verify-files',bytes:0});assert.equal(sent.at(-1).phase,'verify-files');assert.equal(timers.size,0);
  stale();assert.equal(sent.length,4);
  progress.push({phase:'verify-files',bytes:10});const afterFailure=[...timers.values()][0];progress.dispose();
  afterFailure();progress.push({phase:'extract',bytes:999});assert.equal(sent.length,4);assert.equal(timers.size,0);
});

function crc32(b: Buffer) { let c = 0xffffffff; for (const value of b) { c ^= value; for (let i=0;i<8;i++) c=(c>>>1)^((c&1)?0xedb88320:0); } return (c^0xffffffff)>>>0; }
function zip(entries: [string,Buffer][]) {
  const locals: Buffer[] = [], centers: Buffer[] = []; let offset=0;
  for (const [name,data] of entries) {
    const n=Buffer.from(name), c=crc32(data), compressed=deflateRawSync(data), l=Buffer.alloc(30), h=Buffer.alloc(46);
    l.writeUInt32LE(0x04034b50); l.writeUInt16LE(20,4); l.writeUInt16LE(0x800,6); l.writeUInt16LE(8,8); l.writeUInt32LE(c,14); l.writeUInt32LE(compressed.length,18); l.writeUInt32LE(data.length,22); l.writeUInt16LE(n.length,26);
    h.writeUInt32LE(0x02014b50); h.writeUInt16LE(20,4); h.writeUInt16LE(20,6); h.writeUInt16LE(0x800,8); h.writeUInt16LE(8,10); h.writeUInt32LE(c,16); h.writeUInt32LE(compressed.length,20); h.writeUInt32LE(data.length,24); h.writeUInt16LE(n.length,28); h.writeUInt32LE(offset,42);
    locals.push(l,n,compressed); centers.push(h,n); offset+=l.length+n.length+compressed.length;
  }
  const central=Buffer.concat(centers), end=Buffer.alloc(22); end.writeUInt32LE(0x06054b50); end.writeUInt16LE(entries.length,8); end.writeUInt16LE(entries.length,10); end.writeUInt32LE(central.length,12); end.writeUInt32LE(offset,16);
  return Buffer.concat([...locals,central,end]);
}
async function fixture(t: any, suppliedEntries?: [string,Buffer][]) {
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'charfloat-setup-test-')); t.after(()=>fs.rmSync(dir,{recursive:true,force:true}));
  const local=path.join(dir,'用户 甲'), desktop=path.join(dir,'桌面'); fs.mkdirSync(local); fs.mkdirSync(desktop);
  const entries: [string,Buffer][]=suppliedEntries || [[EXE,Buffer.from('synthetic executable')],['resources/app.asar',Buffer.from('synthetic archive')],['resources/中文.txt',Buffer.from('合成内容')]];
  const payload=path.join(dir,'payload.zip'); fs.writeFileSync(payload,zip(entries));
  const manifest={product:PRODUCT,version:'2.2.3',sourceCommit:'a'.repeat(40),bootstrapCommit:'b'.repeat(40),payloadSha256:await fileHash(payload),totalBytes:entries.reduce((n,[,b])=>n+b.length,0),files:entries.map(([name,b])=>({name,size:b.length,sha256:sha(b)}))};
  const shell={writeShortcutLink(file:string,_op:string,options:object){fs.writeFileSync(file,JSON.stringify(options));return true;},readShortcutLink(file:string){return JSON.parse(fs.readFileSync(file,'utf8'));}};
  const shortcut=(args:any)=>createShortcut({...args,desktop,shell});
  return {dir,local,desktop,payload,manifest,shell,shortcut,options:{localAppData:local,payload,manifest,shortcut}};
}
test('用户目录解压、中文路径、逐文件哈希和真实字节进度；重复配置保留额外数据',async t=>{
  const f=await fixture(t), progress:any[]=[]; const result=await install({...f.options,onProgress:(p:any)=>progress.push(p)});
  assert.equal(path.dirname(result.directory),path.join(await fs.promises.realpath(f.local),'Programs','CharFloat')); await verifyInstallation(result.directory,f.manifest);
  assert.equal(fs.readFileSync(path.join(result.directory,'resources/中文.txt'),'utf8'),'合成内容');
  assert.deepEqual(progress.filter(p=>p.phase==='extract'&&p.activity==='file-verified').map(p=>p.verifiedBytes),[20,37,49]);
  assert.ok(progress.some(p=>p.phase==='verify-files'&&p.file==='resources/app.asar'));
  assert.equal(progress.at(-1).bytes,f.manifest.totalBytes); assert.equal(f.shell.readShortcutLink(path.join(f.desktop,'字浮 CharFloat.lnk')).target,result.executable);
  fs.writeFileSync(path.join(result.directory,'my-notes.txt'),'keep'); await install(f.options); assert.equal(fs.readFileSync(path.join(result.directory,'my-notes.txt'),'utf8'),'keep');
});
test('真实压缩大文件和慢磁盘写入持续报告已写字节，整文件哈希通过前不增加已核验值',async t=>{
  const data=Buffer.alloc(8*1024*1024,0x62), f=await fixture(t,[[EXE,Buffer.from('synthetic executable')],['resources/app.asar',data]]);
  const original=fs.createWriteStream;
  // 使用真实 ZIP、解压流和磁盘文件，仅延迟真实 fs.write 回调以稳定复现慢盘背压。
  t.mock.method(fs,'createWriteStream',(file:any,options:any)=>original(file,{...options,highWaterMark:16384,fs:{
    open:fs.open,close:fs.close,
    write(fd:any,buffer:any,offset:any,length:any,position:any,callback:any){fs.write(fd,buffer,offset,length,position,(error,...args)=>setTimeout(()=>callback(error,...args),3));},
    writev(fd:any,buffers:any,position:any,callback:any){fs.writev(fd,buffers,position,(error,...args)=>setTimeout(()=>callback(error,...args),3));}
  }}));
  const events:any[]=[];
  const result=await install({...f.options,onProgress:(p:any)=>{
    events.push({...p,at:Date.now()});
    if(p.phase==='extract'&&p.fileBytes>0&&p.activity==='write'){
      const root=path.join(f.local,'Programs','CharFloat'), pending=fs.readdirSync(root).find(n=>n.startsWith('.pending-'))!;
      assert.ok(fs.statSync(path.join(root,pending,p.file)).size>=p.fileBytes,'计数不能领先真实磁盘写入');
    }
  }});
  const partial=events.filter(p=>p.phase==='extract'&&p.file==='resources/app.asar'&&p.activity==='write'&&p.fileBytes>0&&p.fileBytes<data.length);
  assert.ok(partial.length>=3,`应在单个大文件内部连续更新，实际 ${partial.length} 次`);
  assert.ok(partial.every(p=>p.verifiedBytes===20&&p.bytes===20+p.fileBytes));
  assert.ok(partial.every((p,i)=>!i||p.fileBytes>partial[i-1].fileBytes));
  assert.ok(partial.every((p,i)=>!i||p.at-partial[i-1].at>=180),'IPC 更新有节流');
  const check=events.find(p=>p.phase==='extract'&&p.file==='resources/app.asar'&&p.activity==='check-file');
  assert.equal(check.verifiedBytes,20); assert.equal(check.fileBytes,data.length);
  assert.equal(events.filter(p=>p.phase==='extract'&&p.activity==='file-verified').at(-1).verifiedBytes,f.manifest.totalBytes);
  assert.equal(await fileHash(path.join(result.directory,'resources/app.asar')),sha(data));
  assert.ok(events.some(p=>p.phase==='verify'&&p.bytes>0));
  assert.ok(events.some(p=>p.phase==='verify-files'&&p.activity==='read'&&p.fileBytes>0&&p.verifiedBytes<p.bytes));
});
test('显示写入完成不等于校验通过；错误哈希没有文件完成或成功状态',async t=>{
  const f=await fixture(t), events:any[]=[];
  f.manifest.files[1].sha256='0'.repeat(64);
  await assert.rejects(install({...f.options,onProgress:(p:any)=>events.push(p)}),/解压文件校验失败/);
  const unverified=events.find(p=>p.phase==='extract'&&p.file==='resources/app.asar'&&p.activity==='check-file');
  assert.equal(unverified.bytes,37);assert.equal(unverified.verifiedBytes,20);
  assert.equal(events.some(p=>p.file==='resources/app.asar'&&p.activity==='file-verified'),false);
  assert.equal(events.some(p=>p.phase==='shortcut'),false);assert.equal(fs.readdirSync(f.desktop).length,0);
});
test('真实写流报错会退出，不吞异常、不进入完成态，也不留下进度定时器',async t=>{
  const f=await fixture(t), original=fs.createWriteStream, events:any[]=[];
  t.mock.method(fs,'createWriteStream',(file:any,options:any)=>original(file,{...options,fs:{open:fs.open,close:fs.close,
    write(_fd:any,_buffer:any,_offset:any,_length:any,_position:any,callback:any){queueMicrotask(()=>callback(Object.assign(new Error('simulated write denial'),{code:'EACCES'})));}
  }}));
  await assert.rejects(install({...f.options,onProgress:(p:any)=>events.push(p)}),{code:'EACCES'});
  const count=events.length;await new Promise(r=>setTimeout(r,250));assert.equal(events.length,count);
  assert.equal(events.some(p=>p.activity==='file-verified'||p.phase==='shortcut'),false);
  assert.equal(fs.readdirSync(f.desktop).length,0);
  assert.deepEqual(fs.readdirSync(path.join(f.local,'Programs','CharFloat')),['.charfloat-install.json']);
});
test('工作界面展示真实阶段、当前文件与字节，复核不能冒充配置完成',()=>{
  const nodes=new Map<string,any>(), context=vm.createContext({document:{getElementById(id:string){if(!nodes.has(id))nodes.set(id,{});return nodes.get(id);}},window:{},URLSearchParams,location:{search:'?stage=work'}});
  vm.runInContext(fs.readFileSync('windows-bootstrap/renderer.js','utf8'),context);
  vm.runInContext(`render({stage:'work',phase:'extract',activity:'write',file:'resources/app.asar',bytes:90*1048576,verifiedBytes:85*1048576,totalBytes:315*1048576})`,context);
  assert.equal(nodes.get('status').textContent,'正在解压应用文件');assert.equal(nodes.get('amount').textContent,'90.0 / 315.0 MB');assert.equal(nodes.get('current-file').textContent,'resources/app.asar');
  vm.runInContext(`render({stage:'work',phase:'verify-files',activity:'read',file:'字浮 CharFloat.exe',bytes:200,totalBytes:315})`,context);
  assert.equal(nodes.get('status').textContent,'正在复核应用文件');assert.equal(nodes.get('launch').hidden,true);assert.equal(nodes.get('launch').disabled,true);
});
test('陌生安装目录和同名快捷方式不覆盖',async t=>{
  const f=await fixture(t); const link=path.join(f.desktop,'字浮 CharFloat.lnk'); fs.writeFileSync(link,JSON.stringify({target:path.join(f.dir,'other.exe')}));
  const old=fs.readFileSync(link); await assert.rejects(install(f.options),/同名快捷方式/); assert.deepEqual(fs.readFileSync(link),old);
  const g=await fixture(t), root=path.join(g.local,'Programs','CharFloat'); fs.mkdirSync(root,{recursive:true}); fs.writeFileSync(path.join(root,'user-file'),'keep');
  await assert.rejects(install(g.options)); assert.equal(fs.readFileSync(path.join(root,'user-file'),'utf8'),'keep');
});
test('载荷损坏、文件哈希不符均不创建快捷方式；临时目录清理',async t=>{
  const f=await fixture(t); fs.appendFileSync(f.payload,'x'); await assert.rejects(install(f.options),/安装包校验失败/); assert.equal(fs.readdirSync(f.desktop).length,0);
  const g=await fixture(t); g.manifest.files[0].sha256='0'.repeat(64); await assert.rejects(install(g.options),/解压文件校验失败/);
  assert.deepEqual(fs.readdirSync(path.join(g.local,'Programs','CharFloat')),['.charfloat-install.json']);
});
test('已有安装缺文件或内容改变不覆盖；旧版本仍保留',async t=>{
  const f=await fixture(t); const result=await install(f.options); fs.writeFileSync(result.executable,'user change');
  await assert.rejects(install(f.options),/已改变/); assert.equal(fs.readFileSync(result.executable,'utf8'),'user change');
  fs.unlinkSync(result.executable); await assert.rejects(install(f.options),/不完整/);
});
test('快捷方式写失败不报成功，已解压应用保留，后续重试成功',async t=>{
  const f=await fixture(t), original=f.shell.writeShortcutLink; f.shell.writeShortcutLink=()=>false;
  await assert.rejects(install(f.options),/快捷方式创建失败/); const root=path.join(f.local,'Programs','CharFloat'); const app=fs.readdirSync(root).find(n=>n.startsWith('2.2.3-'))!;
  assert.ok(fs.existsSync(path.join(root,app,EXE))); f.shell.writeShortcutLink=original; await install(f.options);
});
test('预检后出现同名桌面文件不覆盖',async t=>{
  const f=await fixture(t), original=f.shell.writeShortcutLink; f.shell.writeShortcutLink=(file,op,opts)=>{ const r=original(file,op,opts); fs.writeFileSync(path.join(f.desktop,'字浮 CharFloat.lnk'),'new user file'); return r; };
  await assert.rejects(install(f.options)); assert.equal(fs.readFileSync(path.join(f.desktop,'字浮 CharFloat.lnk'),'utf8'),'new user file');
});
test('用户目录内的重定向链接与快捷方式备份链接不跟随',async t=>{
  const f=await fixture(t), outside=path.join(f.dir,'outside'); fs.mkdirSync(outside); fs.symlinkSync(outside,path.join(f.local,'Programs'),'junction');
  await assert.rejects(install(f.options),/链接/); assert.deepEqual(fs.readdirSync(outside),[]);
  const g=await fixture(t); const result=await install(g.options); const other=path.join(g.dir,'other'); fs.mkdirSync(other);
  fs.symlinkSync(other,path.join(path.dirname(result.directory),'shortcut-backups'),'junction');
  await assert.rejects(install(g.options),/备份目录不安全/); assert.deepEqual(fs.readdirSync(other),[]);
});
test('同时配置被锁拒绝；不终止进程，首次配置仍完成',async t=>{
  const f=await fixture(t); let unlock:()=>void=()=>{}, entered:()=>void=()=>{}; const enteredPromise=new Promise<void>(r=>entered=r); const gate=new Promise<void>(r=>unlock=r);
  const first=install({...f.options,shortcut:async (args:any)=>{entered();await gate;await f.shortcut(args);}});
  await enteredPromise; await assert.rejects(install(f.options),/另一个配置正在进行/); unlock(); await first;
});
test('Windows路径特殊名、大小和重复大小写清单拒绝',async t=>{
  for (const name of ['../outside','C:/bad','/root','a\\b','a:stream','a/CON.txt','a/..','a.','a /b','x\0y']) assert.throws(()=>safeName(name));
  const f=await fixture(t); assert.throws(()=>validateManifest({...f.manifest,totalBytes:4*1024**3+1})); assert.throws(()=>validateManifest({...f.manifest,files:[...f.manifest.files,{...f.manifest.files[0],name:EXE.toUpperCase()}]}));
});
test('压缩包越界及重复条目不写入外部目录',async t=>{
  const f=await fixture(t); fs.writeFileSync(f.payload,zip([[EXE,Buffer.from('synthetic executable')],['../outside',Buffer.from('bad')]])); f.manifest.payloadSha256=await fileHash(f.payload);
  await assert.rejects(install(f.options)); assert.equal(fs.existsSync(path.join(f.dir,'outside')),false);
});
test('三个阶段各至少5秒，工作真实耗时不缩短；失败从不进入完成阶段',async()=>{
  let time=0; const states:any[]=[]; const clock={now:()=>time,sleep:async(ms:number)=>{time+=ms;}};
  const result=await runStages(async()=>{time+=2000;return 'ok';},(s:any)=>states.push({...s,time}),clock);
  assert.equal(result,'ok'); assert.deepEqual(states.map(s=>[s.stage,s.enabled,s.time]),[['hello',undefined,0],['work',undefined,5000],['done',false,10000],['done',true,15000]]);
  time=0; states.length=0; await assert.rejects(runStages(async()=>{throw new Error('write failed');},(s:any)=>states.push({...s,time}),clock),/write failed/);
  assert.equal(time,10000); assert.equal(states.some(s=>s.stage==='done'),false);
  time=0; states.length=0; await runStages(async()=>{time+=8000;},(s:any)=>states.push({...s,time}),clock); assert.equal(states.find(s=>s.stage==='done').time,13000);
});
test('显式启动再次校验文件、清除外壳环境，启动失败不当成功',async t=>{
  const f=await fixture(t), result=await install(f.options); let calls=0, options:any;
  const spawnProcess=(exe:string,args:any[],opts:any)=>{calls++;assert.equal(exe,result.executable);assert.deepEqual(args,[]);options=opts;const child=new EventEmitter() as any;child.unref=()=>{};queueMicrotask(()=>child.emit('spawn'));return child;};
  await launchApplication({...result,manifest:f.manifest,spawnProcess,environment:{ELECTRON_RUN_AS_NODE:'1',PORTABLE_EXECUTABLE_DIR:'outer',CHARFLOAT_DIAGNOSTIC_RUN_ID:'temporary',CHARFLOAT_DIAGNOSTIC_DIR:'temporary',charfloat_diagnostic_lower:'temporary',KEEP:'yes'}});
  assert.equal(calls,1);assert.deepEqual(options.env,{KEEP:'yes'});assert.equal(options.cwd,result.directory);assert.equal(options.shell,undefined);
  fs.writeFileSync(result.executable,'changed');await assert.rejects(launchApplication({...result,manifest:f.manifest,spawnProcess}));assert.equal(calls,1);
  const g=await fixture(t), good=await install(g.options);
  await assert.rejects(launchApplication({...good,manifest:g.manifest,spawnProcess:()=>{const child=new EventEmitter() as any;child.unref=()=>{};queueMicrotask(()=>child.emit('error',new Error('spawn denied')));return child;}}),/spawn denied/);
});
test('实际打包主进程：单实例/IPC门禁、三态计时、重复请求、失败关闭和点击启动',async t=>{
  const compiled=await build({entryPoints:[path.resolve('windows-bootstrap/main.mjs')],bundle:true,platform:'node',format:'cjs',external:['electron'],write:false});
  async function harness(f:any,{lock=true,fail=false,profileFailure=false}={}) {
    let time=0,quits=0,spawned=0; const states:any[]=[],handlers=new Map(),windows:any[]=[];
    fs.writeFileSync(path.join(f.dir,'payload-manifest.json'),JSON.stringify(f.manifest));
    if(profileFailure)fs.writeFileSync(path.join(f.dir,'CharFloat Setup'),'existing file');
    class Window {
      visible=false; events=new Map();webContents={mainFrame:{},setWindowOpenHandler(){},on(){},send(_channel:string,state:any){states.push({...state,time});if(state.paintToken)queueMicrotask(()=>handlers.get('setup:painted')({sender:this,senderFrame:this.mainFrame},state.paintToken));}};
      constructor(){windows.push(this);}once(name:string,fn:()=>void){if(name==='ready-to-show')queueMicrotask(fn);return this;}
      show(){this.visible=true;}isVisible(){return this.visible;}isDestroyed(){return false;}on(name:string,fn:any){this.events.set(name,fn);}async loadFile(){}isMinimized(){return false;}restore(){}focus(){}
    }
    const nativeRequire=createRequire(import.meta.url);
    const electron={app:{getPath:(name:string)=>name==='desktop'?f.desktop:f.dir,setName(){},setPath(){},requestSingleInstanceLock:()=>lock,on(){},whenReady:()=>Promise.resolve(),quit(){quits++;},exit(){quits++;}},BrowserWindow:Window,ipcMain:{handle:(name:string,fn:any)=>handlers.set(name,fn)},shell:fail?{...f.shell,writeShortcutLink:()=>false}:f.shell,dialog:{showMessageBox:async()=>{}}};
    const proc={...process,platform:'win32',resourcesPath:f.dir,env:{...process.env,LOCALAPPDATA:f.local}};
    vm.runInNewContext(compiled.outputFiles[0].text,{require:(name:string)=>name==='electron'?electron:name==='node:child_process'?{spawn:()=>{spawned++;const child=new EventEmitter() as any;child.unref=()=>{};queueMicrotask(()=>child.emit('spawn'));return child;}}:nativeRequire(name),__dirname:f.dir,process:proc,Buffer,console,setInterval,clearInterval,setImmediate,clearImmediate,queueMicrotask,Date:class extends Date{static now(){return time;}},setTimeout:(fn:()=>void,ms:number)=>{if(ms>=20000)return 2;time+=ms;queueMicrotask(fn);return 1;},clearTimeout(){},exports:{},module:{exports:{}}});
    await new Promise(r=>setImmediate(r));
    const event=windows[0]?{sender:windows[0].webContents,senderFrame:windows[0].webContents.mainFrame}:{};
    return {states,handlers,windows,event,quits:()=>quits,spawned:()=>spawned};
  }
  const f=await fixture(t), h=await harness(f);
  await assert.rejects(h.handlers.get('setup:start')({sender:{}}),/无效请求/);
  const pending=h.handlers.get('setup:start')(h.event);
  assert.equal((await h.handlers.get('setup:launch')(h.event)).ok,false);
  let prevented=false;h.windows[0].events.get('close')({preventDefault(){prevented=true;}});assert.equal(prevented,true);
  await pending; assert.equal(h.states.find(s=>s.stage==='error')?.message,undefined); assert.deepEqual(h.states.filter(s=>s.stage!=='work'||s.phase==='verify'&&s.bytes===0).map(s=>[s.stage,s.enabled,s.time]),[['hello',undefined,0],['work',undefined,5000],['work',undefined,5000],['done',false,10000],['done',true,15000]]);
  const count=h.states.length; await h.handlers.get('setup:start')(h.event);assert.equal(h.states.length,count);
  assert.equal((await h.handlers.get('setup:launch')(h.event)).ok,true);assert.equal(h.spawned(),1);assert.equal((await h.handlers.get('setup:launch')(h.event)).ok,false);assert.equal(h.spawned(),1);
  const bad=await harness(await fixture(t),{fail:true});await bad.handlers.get('setup:start')(bad.event);assert.equal(bad.states.at(-1).stage,'error');assert.equal(bad.states.some(s=>s.stage==='done'),false);
  await bad.handlers.get('setup:close')(bad.event);assert.equal(bad.quits(),1);assert.equal(bad.spawned(),0);
  const second=await harness(await fixture(t),{lock:false});assert.equal(second.windows.length,0);assert.equal(second.quits(),1);
  const noProfile=await harness(await fixture(t),{profileFailure:true});assert.equal(noProfile.windows.length,0);assert.equal(noProfile.quits(),1);
});
