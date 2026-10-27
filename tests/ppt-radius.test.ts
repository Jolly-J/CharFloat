import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import { execFileSync } from 'node:child_process';
import { buildSync } from 'esbuild';
import { XMLSerializer } from '@xmldom/xmldom';
import { readPptx, writePptx, inspectRadius, patchRadius, verifyRadius, xml, type Parts } from '../src/bridge/gateway/pptx-radius.js';
import { setCornerRadius } from '../src/bridge/gateway/ppt-radius.js';
import { pptToolDefinitions } from '../src/bridge/tools/definitions/ppt.js';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main';
function fixture():Parts {
  const shape=(id:number,name:string,w:number,h:number)=>`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="127000" y="254000"/><a:ext cx="${w*12700}" cy="${h*12700}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></p:spPr></p:sp>`;
  const slide=`<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>${shape(2,'bgA',240,100)}${shape(3,'bgB',300,80)}</p:spTree></p:cSld></p:sld>`;
  return new Map(Object.entries({
    'ppt/presentation.xml':`<p:presentation xmlns:p="${P}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId1"/></p:sldIdLst></p:presentation>`,
    'ppt/_rels/presentation.xml.rels':'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="slides/slide1.xml"/><Relationship Id="rId2" Target="slides/slide7.xml"/></Relationships>',
    'ppt/slides/slide7.xml':slide,'ppt/slides/slide1.xml':slide,'ppt/media/test.png':'synthetic-image-bytes',
  }).map(([n,b])=>[n,Buffer.from(b)]));
}
const putXml=(parts:Parts,n:string,doc:any)=>parts.set(n,Buffer.from(new XMLSerializer().serializeToString(doc)));
function simulatedImport(before:Parts,radius=8) {
  const plan=inspectRadius(before,1,['bgA','bgB'],radius),patched=patchRadius(before,inspectRadius(before,1,['bgA','bgB'],radius));
  const doc=xml(patched.get(plan.part)!), tree=doc.getElementsByTagNameNS(P,'spTree')[0], originals=Array.from(tree.getElementsByTagNameNS(P,'sp'));
  const receipts=originals.map((s,i)=>{
    const fresh=s.cloneNode(true) as any,nv=s.getElementsByTagNameNS(P,'cNvPr')[0],freshNv=fresh.getElementsByTagNameNS(P,'cNvPr')[0];
    const name=nv.getAttribute('name'),oldId=Number(nv.getAttribute('id')),newId=100+i;
    // Restore the archived original's geometry, as Copy/Paste only modifies the new shape.
    const original=xml(before.get(plan.part)!).getElementsByTagNameNS(P,'sp')[i];
    tree.replaceChild(original,s);const oldNv=original.getElementsByTagNameNS(P,'cNvPr')[0];
    oldNv.setAttribute('name','__cfRadius_test_'+i);oldNv.setAttribute('hidden','1');
    freshNv.setAttribute('id',String(newId));tree.insertBefore(fresh,original);
    return {name,oldId,newId,oldName:'__cfRadius_test_'+i};
  });
  putXml(patched,plan.part,doc);return {parts:patched,plan,receipts};
}
test('absolute radius uses actual short sides, follows presentation relationships and preserves unrelated parts',async()=>{
  const before=fixture(),plan=inspectRadius(before,1,['bgA','bgB'],8);
  assert.equal(plan.part,'ppt/slides/slide7.xml');assert.deepEqual(plan.targets.map(t=>t.adjustment),[8000,10000]);
  const patched=patchRadius(before,plan),r=inspectRadius(patched,1,['bgA','bgB'],8);
  assert.deepEqual(r.targets.map(t=>t.currentRadiusPt),[8,8]);
  for(const [n,b] of before) if(n!==plan.part) assert.deepEqual(patched.get(n),b);
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'corner-zip-')),file=path.join(dir,'test.pptx');
  try {writePptx(patched,file);assert.deepEqual(await readPptx(file),patched);assert.throws(()=>writePptx(patched,file),/EEXIST/);}finally{fs.rmSync(dir,{recursive:true,force:true});}
});
test('zero/max radii work; oversize, text, group, hidden, non-roundRect, duplicate and identity references reject',()=>{
  for(const radius of [0,40])assert.ok(inspectRadius(fixture(),1,['bgB'],radius));
  assert.throws(()=>inspectRadius(fixture(),1,['bgB'],41),/短边/);
  for(const r of [-1,NaN,Infinity])assert.throws(()=>inspectRadius(fixture(),1,['bgB'],r),/无效/);
  const edits=[['name="bgB"','name="bgA"'],['prst="roundRect"','prst="rect"'],['id="2" name="bgA"','id="2" name="bgA" hidden="1"'],['<p:spPr>','<p:txBody><a:p><a:r><a:t>正文</a:t></a:r></a:p></p:txBody><p:spPr>'],['</p:cSld>','</p:cSld><p:timing/>'],['</p:spTree>','<p:cxnSp><p:cNvCxnSpPr><a:stCxn id="2" idx="0"/></p:cNvCxnSpPr></p:cxnSp></p:spTree>']];
  for(const [a,b] of edits){const p=fixture();p.set('ppt/slides/slide7.xml',Buffer.from(p.get('ppt/slides/slide7.xml')!.toString().replace(a,b)));assert.throws(()=>inspectRadius(p,1,['bgA'],8));}
  assert.throws(()=>xml('<!DOCTYPE x [<!ENTITY x SYSTEM "file:///etc/passwd">]><x>&x;</x>'),/DTD/);
});
test('saved verification catches wrong radius, style drift, unrelated changes and layer reordering',()=>{
  const before=fixture(),saved=simulatedImport(before);
  assert.deepEqual(verifyRadius(before,saved.parts,saved.plan,saved.receipts).map(t=>t.actualRadiusPt),[8,8]);
  for(const [a,b] of [['val 8000','val 2000'],['val="0070C0"','val="FFFFFF"']]) {
    const bad=new Map(saved.parts);bad.set(saved.plan.part,Buffer.from(bad.get(saved.plan.part)!.toString().replace(a,b)));
    assert.throws(()=>verifyRadius(before,bad,saved.plan,saved.receipts));
  }
  const bad=new Map(saved.parts);bad.set('ppt/media/test.png',Buffer.from('changed'));assert.throws(()=>verifyRadius(before,bad,saved.plan,saved.receipts),/媒体/);
  const reordered=new Map(saved.parts),doc=xml(reordered.get(saved.plan.part)!),tree=doc.getElementsByTagNameNS(P,'spTree')[0];
  tree.appendChild(tree.firstChild!);putXml(reordered,saved.plan.part,doc);assert.throws(()=>verifyRadius(before,reordered,saved.plan,saved.receipts),/顺序/);
});

function hostFixture(failAt?:string) {
  let copied:any,sourceClosed=0;const items:any[]=[];
  const make=(id:number,name:string)=>({Id:id,Name:name,Type:1,AutoShapeType:5,Visible:-1,Left:10,Top:20,Width:240,Height:100,Rotation:0,
    TextFrame:{TextRange:{Text:''}},Fill:{ForeColor:{RGB:123}},Line:{Visible:0,Weight:1},
    get ZOrderPosition(){return items.indexOf(this)+1;},
    Copy(){copied=this;},ZOrder(command:number){const i=items.indexOf(this);items.splice(i,1);if(command===1)items.unshift(this);else items.splice(Math.min(i+1,items.length),0,this);},
    Delete(){items.splice(items.indexOf(this),1);}});
  items.push(make(2,'bgA'),make(3,'bgB'));items[1].Width=300;items[1].Height=80;
  const shapes={get Count(){return items.length;},Item:(i:any)=>typeof i==='number'?items[i-1]:items.find(s=>s.Name===i),Paste:()=>{const s=make(100+items.length,copied.Name);s.Width=copied.Width;s.Height=copied.Height;items.push(s);if(failAt==='rename')Object.defineProperty(s,'Name',{get:()=>copied.Name,set:()=>{throw Error('rename injection');}});return {Item:()=>s};}};
  const slide={SlideID:256,Shapes:shapes},pres={Name:'test.pptx',FullName:'/tmp/test.pptx',Path:'/tmp',Saved:true,Slides:{Count:1,Item:()=>slide},Save:()=>{if(failAt==='save')throw Error('save injection');},SaveCopyAs:()=>{}};
  const source={FullName:'/tmp/source.pptx',Slides:{Item:()=>({Shapes:{Item:(n:string)=>items.find(s=>s.Name===n)}})},Close:()=>{sourceClosed++;}};
  const docs:any[]=[pres],app={ActiveWindow:{Activate:()=>{}},Presentations:{get Count(){return docs.length;},Item:(i:number)=>docs[i-1],Open:()=>docs.push(source)}};
  const ctx:any=vm.createContext({pptReadShapeStyle:(s:any)=>({color:s.Fill.ForeColor.RGB}),pptRequireSlideIndex:(_:any,n:number)=>n});
  vm.runInContext(fs.readFileSync('wps-addon/src/ppt-radius.js','utf8'),ctx);
  const call=(action:string,p:any={})=>JSON.parse(JSON.stringify(ctx.pptCornerRadius(app,{action,presentationName:'test.pptx',slideIndex:1,operationId:'test',...p})));
  const snap=call('snapshot',{shapeNames:['bgA','bgB'],snapshotPath:'/tmp/snapshot.pptx'});
  return {app,items,call,snap,get sourceClosed(){return sourceClosed;},apply:()=>call('apply',{...snap,sourcePath:'/tmp/source.pptx'})};
}
test('host import preserves visible order, is idempotent, status detects later changes, rollback protects them',()=>{
  const h=hostFixture(),r=h.apply();assert.equal(r.state,'applied');assert.equal(r.receipts.length,2);assert.equal(h.sourceClosed,1);
  assert.deepEqual(h.items.filter(s=>s.Visible!==0).map(s=>s.Name),['bgA','bgB']);
  assert.equal(h.apply().receipts.length,2);assert.equal(h.items.length,4);
  h.items.find(s=>s.Name==='bgA').Fill.ForeColor.RGB=456;
  assert.equal(h.call('status').targetChanged,true);assert.throws(()=>h.call('rollback'),/后续修改/);assert.equal(h.items.length,4);
});
test('host rejects concurrent target changes before paste and records partial writes for readback/rollback',()=>{
  const h=hostFixture();h.items[0].Left=11;assert.throws(()=>h.apply(),/修改/);assert.equal(h.items.length,2);
  const partial=hostFixture('save'),r=partial.apply();assert.equal(r.state,'partial');assert.equal(r.saved,false);assert.equal(r.receipts.length,2);assert.equal(partial.apply().state,'partial');
  const good=hostFixture();good.apply();assert.equal(good.call('rollback').state,'rolledBack');assert.equal(good.items.length,2);assert.ok(good.items.every(s=>s.Visible===-1));
  const layer=hostFixture();layer.apply();layer.items.find(s=>s.Name==='bgB').ZOrder(1);assert.equal(layer.call('status').targetChanged,true);assert.throws(()=>layer.call('rollback'),/后续修改/);
  const rename=hostFixture('rename');assert.equal(rename.apply().receipts.length,1);assert.equal(rename.items.length,3);
});
test('new tool is appended, declares pt and explicit targets, and gateway validates without any host call',async()=>{
  const f=pptToolDefinitions().at(-1)!.function;assert.equal(f.name,'wps_ppt_set_corner_radius');assert.deepEqual(f.parameters.properties.action.enum,['plan','apply','status','rollback']);
  let calls=0;const ctx:any={args:{presentationName:'test.pptx',slideIndex:1,action:'apply',operationId:'test',shapeNames:['a'],radiusPt:NaN},currentHost:()=> 'wps',previewPath:()=>path.join(os.tmpdir(),'test.json'),callOffice:()=>{calls++;}};
  await assert.rejects(()=>setCornerRadius(ctx));assert.equal(calls,0);
});
test('rebuilt or protected addon dispatches corner snapshot and native import through the real RPC branch',async()=>{
  const h=hostFixture();let socket:any;const sent:any[]=[];
  class Socket{static OPEN=1;readyState=1;onopen:any;onmessage:any;constructor(){socket=this;}send(s:string){sent.push(JSON.parse(s));}}
  const ctx=vm.createContext({window:{WPS_BRIDGE_CONFIG:{port:12345,token:'fake'},addEventListener:()=>{}},document:{getElementById:()=>null,readyState:'complete'},wps:{WppApplication:()=>h.app},WebSocket:Socket,console:{log:()=>{}},setTimeout:()=>0,clearTimeout:()=>{},setInterval:()=>0,alert:()=>{}});
  vm.runInContext(fs.readFileSync(process.env.OAB_TEST_WPS_ARTIFACT||'wps-addon/addon-core.js','utf8'),ctx);socket.onopen();
  async function rpc(id:string,params:any){socket.onmessage({data:JSON.stringify({id,method:'ppt_corner_radius',params})});await new Promise(r=>setImmediate(r));const response=sent.find(s=>s.type==='rpc_response'&&s.id===id);assert.equal(response.error,null);return response.result;}
  const base={presentationName:'test.pptx',slideIndex:1,operationId:'rpc_test'};
  const snap=await rpc('snapshot',{...base,action:'snapshot',shapeNames:['bgA','bgB'],snapshotPath:'/tmp/snapshot.pptx'});
  const result=await rpc('apply',{...base,...snap,action:'apply',sourcePath:'/tmp/source.pptx'});
  assert.equal(result.state,'applied');assert.equal(result.receipts.length,2);assert.equal(h.items.length,4);
});
test('bounded ZIP reader rejects path traversal and CRC corruption',async()=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'corner-invalid-'));
  try {
    const bad=path.join(dir,'path.pptx');writePptx(new Map([['../outside',Buffer.from('x')]]),bad);await assert.rejects(()=>readPptx(bad));
    const corrupt=path.join(dir,'crc.pptx');writePptx(fixture(),corrupt);const b=fs.readFileSync(corrupt),off=b.indexOf(Buffer.from([0x50,0x4b,0x01,0x02]));b.writeUInt32LE(0,off+16);fs.writeFileSync(corrupt,b);await assert.rejects(()=>readPptx(corrupt),/校验/);
  }finally{fs.rmSync(dir,{recursive:true,force:true});}
});
test('gateway plan leaves original bytes unchanged; apply verifies saved radius; unknown calls never replay', {skip:process.platform!=='darwin'},async()=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'corner-flow-')),target=path.join(dir,'test.pptx'),before=fixture();writePptx(before,target);
  const original=fs.readFileSync(target),calls:string[]=[];let timeout=false,partial=false;
  const args={presentationName:'test.pptx',slideIndex:1,shapeNames:['bgA','bgB'],radiusPt:8,operationId:'flow_test',action:'plan'};
  const ctx:any={args,currentHost:()=> 'wps',previewPath:()=>path.join(dir,'random.json'),callOffice:async(method:string,p:any)=>{
    calls.push(p.action||method);
    if(p.action==='snapshot'){fs.copyFileSync(target,p.snapshotPath);return {targetPath:target,slideId:256,shapeCount:2,targets:[{id:2,name:'bgA'},{id:3,name:'bgB'}]};}
    if(p.action==='apply'){
      if(timeout)throw Error('transport timeout');
      if(partial)return {state:'partial',liveApplied:false,saved:false,failure:'paste failed',receipts:[]};
      const imported=simulatedImport(before);fs.unlinkSync(target);writePptx(imported.parts,target);
      return {state:'applied',liveApplied:true,saved:true,receipts:imported.receipts};
    }
    return {success:false};
  }};
  try{
    const plan:any=await setCornerRadius(ctx);assert.equal(plan.state,'planned');assert.deepEqual(fs.readFileSync(target),original);
    args.action='apply';const result:any=await setCornerRadius(ctx);assert.equal(result.ooxmlVerified,true);assert.equal(result.previewVerified,false);
    const n=calls.length;await setCornerRadius(ctx);assert.equal(calls.length,n,'completed operation is not replayed');
    fs.writeFileSync(target,original);args.operationId='unknown_test';timeout=true;
    const failed:any=await setCornerRadius(ctx);assert.equal(failed.state,'unknown');assert.equal(failed.success,false);const count=calls.length;
    await setCornerRadius(ctx);assert.equal(calls.length,count,'unknown mutation cannot be replayed');
    timeout=false;partial=true;args.operationId='partial_test';const failedPartial:any=await setCornerRadius(ctx);assert.equal(failedPartial.success,false);
    const repeated:any=await setCornerRadius(ctx);assert.equal(repeated.success,false);assert.equal(repeated.error,'paste failed');
  }finally{fs.rmSync(dir,{recursive:true,force:true});}
});
test('PPTX backend reads and writes native geometry inside the actual Electron bytecode runtime',()=>{
  const root=process.cwd(),runtime=path.join(root,'node_modules/electron/dist',process.platform==='darwin'?'Electron.app/Contents/MacOS/Electron':process.platform==='win32'?'electron.exe':'electron');
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'corner-bytecode-')),input=path.join(dir,'input.pptx'),output=path.join(dir,'output.pptx'),source=path.join(dir,'probe.cjs'),compiled=path.join(dir,'probe.jsc');
  try{
    writePptx(fixture(),input);
    buildSync({stdin:{contents:`import assert from 'node:assert/strict';import {readPptx,inspectRadius,patchRadius,writePptx} from ${JSON.stringify(path.join(root,'src/bridge/gateway/pptx-radius.ts'))};
      (async()=>{const before=await readPptx(${JSON.stringify(input)}),plan=inspectRadius(before,1,['bgA','bgB'],8);writePptx(patchRadius(before,plan),${JSON.stringify(output)});const saved=inspectRadius(await readPptx(${JSON.stringify(output)}),1,['bgA','bgB'],8);assert.deepEqual(saved.targets.map(t=>t.currentRadiusPt),[8,8]);console.log('radius-bytecode-pass');})().catch(e=>{console.error(e);process.exitCode=1;});`,resolveDir:root,loader:'ts'},outfile:source,bundle:true,platform:'node',format:'cjs',target:'node20',minify:true});
    const bytenode=path.join(root,'node_modules/bytenode'),env={...process.env,ELECTRON_RUN_AS_NODE:'1'};
    execFileSync(runtime,['-e',`require(${JSON.stringify(bytenode)}).compileFile({filename:${JSON.stringify(source)},output:${JSON.stringify(compiled)}});`],{env,timeout:30000});
    const result=execFileSync(runtime,['-e',`require(${JSON.stringify(bytenode)});require(${JSON.stringify(compiled)});`],{env,encoding:'utf8',timeout:30000});assert.equal(result.trim(),'radius-bytecode-pass');
  }finally{fs.rmSync(dir,{recursive:true,force:true});}
});
