// 在临时副本删除部署入口后重建；不删除当前工作区生成物，不部署。
import fs from 'node:fs';import path from 'node:path';import os from 'node:os';import crypto from 'node:crypto';import assert from 'node:assert/strict';
import {execFileSync} from 'node:child_process';import {fileURLToPath} from 'node:url';
const root=fileURLToPath(new URL('../../../../',import.meta.url)),dir=path.dirname(fileURLToPath(import.meta.url));
const tmp=fs.mkdtempSync(path.join(os.tmpdir(),'charfloat-clean-addons-'));
const hash=file=>crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
const artifacts=['wps-addon/addon-core.js','office-addon/public/taskpane.js'];
const expected=Object.fromEntries(artifacts.map(file=>[file,hash(path.join(root,file))]));
try{
 const copy=(relative)=>{const target=path.join(tmp,relative);fs.mkdirSync(path.dirname(target),{recursive:true});fs.cpSync(path.join(root,relative),target,{recursive:true});};
 for(const file of ['package.json','wps-addon','office-addon','resources/icon.png','scripts/build-wps-addon.mjs','scripts/build-office-addon.mjs','scripts/build-wps-ribbon-icons.mjs','scripts/build-win-icon.mjs','scripts/sync-addon-panel-style.mjs','scripts/lib/wps-addon-modules.mjs'])copy(file);
 fs.symlinkSync(path.join(root,'node_modules'),path.join(tmp,'node_modules'),'dir');
 for(const file of artifacts)fs.unlinkSync(path.join(tmp,file));
 const output=execFileSync('npm',['run','build:addons'],{cwd:tmp,encoding:'utf8',timeout:30000});
 fs.writeFileSync(path.join(dir,'clean-build.log'),output);
 for(const file of artifacts)assert.equal(hash(path.join(tmp,file)),expected[file],`干净重建不一致: ${file}`);
 for(const script of ['build-wps-addon.mjs','build-office-addon.mjs'])execFileSync(process.execPath,['scripts/'+script,'--check'],{cwd:tmp});
 execFileSync('npm',['run','build:addons'],{cwd:tmp,stdio:'pipe',timeout:30000});
 for(const file of artifacts){assert.equal(hash(path.join(tmp,file)),expected[file],`重复构建不一致: ${file}`);assert.equal(hash(path.join(root,file)),expected[file],`当前工作区产物被改动: ${file}`);}
 const rejection=(pattern)=>{
  let failure;try{execFileSync(process.execPath,['scripts/build-wps-addon.mjs','--check'],{cwd:tmp,encoding:'utf8',stdio:'pipe'});}catch(error){failure=error;}
  assert.equal(failure?.status,1,'真实构建入口必须以退出码 1 拒绝漂移');
  assert.match(String(failure.stderr),pattern);
  assert.equal(hash(path.join(tmp,artifacts[0])),expected[artifacts[0]],'检查失败不得覆写生成物');
 };
 const forgotten=path.join(tmp,'wps-addon/src/excel/unlisted.js');fs.writeFileSync(forgotten,'function unlisted() {}');rejection(/未登记模块.*excel\/unlisted.js/);fs.unlinkSync(forgotten);
 const existing=path.join(tmp,'wps-addon/src/excel/pure.js'),content=fs.readFileSync(existing);fs.unlinkSync(existing);rejection(/文件不存在.*excel\/pure.js/);fs.writeFileSync(existing,content);
 const script=path.join(tmp,'scripts/build-wps-addon.mjs'),scriptContent=fs.readFileSync(script,'utf8');fs.writeFileSync(script,scriptContent.replace('  "excel/pure.js",','  "excel/pure.js",\n  "excel/pure.js",'));rejection(/重复登记.*excel\/pure.js/);fs.writeFileSync(script,scriptContent);
 const result={cleanBuildAndChecksPassed:true,repeatBuildByteIdentical:true,cliDriftRejections:{unlisted:1,missing:1,duplicate:1},currentWorkspaceUntouched:true,sha256:expected};
 fs.writeFileSync(path.join(dir,'clean-build.json'),JSON.stringify(result,null,2)+'\n');console.log(JSON.stringify(result,null,2));
}finally{fs.rmSync(tmp,{recursive:true,force:true})}
