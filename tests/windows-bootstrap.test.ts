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

const sha = (b: Buffer) => createHash('sha256').update(b).digest('hex');
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
async function fixture(t: any) {
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'charfloat-setup-test-')); t.after(()=>fs.rmSync(dir,{recursive:true,force:true}));
  const local=path.join(dir,'用户 甲'), desktop=path.join(dir,'桌面'); fs.mkdirSync(local); fs.mkdirSync(desktop);
  const entries: [string,Buffer][]=[[EXE,Buffer.from('synthetic executable')],['resources/app.asar',Buffer.from('synthetic archive')],['resources/中文.txt',Buffer.from('合成内容')]];
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
  assert.deepEqual(progress.filter(p=>p.phase==='extract').map(p=>p.bytes),[20,37,49]);
  assert.equal(progress.at(-1).bytes,f.manifest.totalBytes); assert.equal(f.shell.readShortcutLink(path.join(f.desktop,'字浮 CharFloat.lnk')).target,result.executable);
  fs.writeFileSync(path.join(result.directory,'my-notes.txt'),'keep'); await install(f.options); assert.equal(fs.readFileSync(path.join(result.directory,'my-notes.txt'),'utf8'),'keep');
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
  await launchApplication({...result,manifest:f.manifest,spawnProcess,environment:{ELECTRON_RUN_AS_NODE:'1',PORTABLE_EXECUTABLE_DIR:'outer',KEEP:'yes'}});
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
    vm.runInNewContext(compiled.outputFiles[0].text,{require:(name:string)=>name==='electron'?electron:name==='node:child_process'?{spawn:()=>{spawned++;const child=new EventEmitter() as any;child.unref=()=>{};queueMicrotask(()=>child.emit('spawn'));return child;}}:nativeRequire(name),__dirname:f.dir,process:proc,Buffer,console,setImmediate,clearImmediate,queueMicrotask,Date:class extends Date{static now(){return time;}},setTimeout:(fn:()=>void,ms:number)=>{if(ms>=20000)return 2;time+=ms;queueMicrotask(fn);return 1;},clearTimeout(){},exports:{},module:{exports:{}}});
    await new Promise(r=>setImmediate(r));
    const event=windows[0]?{sender:windows[0].webContents,senderFrame:windows[0].webContents.mainFrame}:{};
    return {states,handlers,windows,event,quits:()=>quits,spawned:()=>spawned};
  }
  const f=await fixture(t), h=await harness(f);
  await assert.rejects(h.handlers.get('setup:start')({sender:{}}),/无效请求/);
  const pending=h.handlers.get('setup:start')(h.event);
  assert.equal((await h.handlers.get('setup:launch')(h.event)).ok,false);
  let prevented=false;h.windows[0].events.get('close')({preventDefault(){prevented=true;}});assert.equal(prevented,true);
  await pending; assert.equal(h.states.find(s=>s.stage==='error')?.message,undefined); assert.deepEqual(h.states.filter(s=>s.stage!=='work'||s.phase==='verify').map(s=>[s.stage,s.enabled,s.time]),[['hello',undefined,0],['work',undefined,5000],['work',undefined,5000],['done',false,10000],['done',true,15000]]);
  const count=h.states.length; await h.handlers.get('setup:start')(h.event);assert.equal(h.states.length,count);
  assert.equal((await h.handlers.get('setup:launch')(h.event)).ok,true);assert.equal(h.spawned(),1);assert.equal((await h.handlers.get('setup:launch')(h.event)).ok,false);assert.equal(h.spawned(),1);
  const bad=await harness(await fixture(t),{fail:true});await bad.handlers.get('setup:start')(bad.event);assert.equal(bad.states.at(-1).stage,'error');assert.equal(bad.states.some(s=>s.stage==='done'),false);
  await bad.handlers.get('setup:close')(bad.event);assert.equal(bad.quits(),1);assert.equal(bad.spawned(),0);
  const second=await harness(await fixture(t),{lock:false});assert.equal(second.windows.length,0);assert.equal(second.quits(),1);
  const noProfile=await harness(await fixture(t),{profileFailure:true});assert.equal(noProfile.windows.length,0);assert.equal(noProfile.quits(),1);
});
