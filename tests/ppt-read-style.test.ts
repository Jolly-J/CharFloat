import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { pptToolDefinitions } from '../src/bridge/tools/definitions/ppt.js';
import { getSlideShapes, manageSlides, captureSlidePreview } from '../src/bridge/gateway/ppt.js';

const fragment = fs.readFileSync('wps-addon/src/ppt.js', 'utf8');
const fixtureRoot = 'tests/fixtures/ppt-style/';
const fixtureMetadata = JSON.parse(fs.readFileSync(fixtureRoot+'metadata.json','utf8'));
const legacy = fs.readFileSync(fixtureRoot+'legacy-get-slide-shapes.js','utf8');
function collection(items: any[]) { return { Count: items.length, Item: (i: number) => items[i-1] }; }
function shape(id: number, children?: any[]) {
  return { Id:id, Name:`shape${id}`, Type: children ? 6 : 17, Left:10, Top:20, Width:100, Height:40, Rotation:0, ZOrderPosition:1,
    HasTextFrame:true, HasTable:false, HasChart:false,
    TextFrame:{HasText:true,MarginLeft:2,MarginRight:2,MarginTop:1,MarginBottom:1,AutoSize:0,WordWrap:-1,VerticalAnchor:1,
      TextRange:{Text:'合成示例', Font:{Name:'Sample',NameFarEast:'示例字体',Size:18,Bold:-1,Color:{RGB:123}},
        ParagraphFormat:{Alignment:1,SpaceWithin:1.1,LineRuleWithin:-1,SpaceBefore:0,SpaceAfter:0},BoundLeft:12,BoundTop:21,BoundWidth:60,BoundHeight:20}},
    Fill:{Visible:-1,Type:1,ForeColor:{RGB:456},Transparency:0},Line:{Visible:0,ForeColor:{RGB:789},Weight:1},
    ...(children ? {GroupItems:collection(children)} : {}) };
}
function read(items:any[], params:any={}, source=fragment) {
  const shapes:any = collection(items); shapes.Placeholders=collection([]);
  const slide={SlideID:101,Shapes:shapes,Layout:12,CustomLayout:{Name:'样本版式'}};
  const pres={Name:'合成.pptx',Slides:collection([slide]),PageSetup:{SlideWidth:1000,SlideHeight:600}};
  const ctx:any=vm.createContext({getPptPresentation:()=>pres,getPptApp:()=>null,getShapeTypeName:(t:number)=>t===6?'group':'textbox',
    pptPageSize:()=>({pageWidth:1000,pageHeight:600,unit:'pt'}),log:()=>{}});
  vm.runInContext(source,ctx);
  return JSON.parse(JSON.stringify(ctx.pptGetSlideShapes({}, {slideIndex:1,...params})));
}

test('default successful shape response matches anchored legacy implementation; no style/group reads',()=>{
  const s=shape(10); Object.defineProperty(s,'GroupItems',{get(){throw Error('must not traverse');}});
  assert.deepEqual(read([s]),read([s],{},legacy));
  assert.equal(read([s]).slideId,undefined);
});
test('optional style returns observed zero/false values, fields and stable ID path; bounded text',()=>{
  const s=shape(10);s.TextFrame.TextRange.Text='文'.repeat(201);
  const r=read([s],{includeStyle:true});const x=r.shapes[0];
  assert.deepEqual(x.locator,{slideId:101,shapeIdPath:[10]});
  assert.equal(x.style.values.font.nameFarEast,'示例字体');assert.equal(x.style.values.font.bold,-1);
  assert.equal(x.style.values.fill.transparency,0);assert.equal(x.style.values.paragraph.spaceBefore,0);
  assert.equal(x.style.values.bounds.height,20);assert.deepEqual(x.style.unavailable,[]);
  assert.equal(x.textTruncated,true);assert.equal(x.text.length,203);
});
test('unreadable Font does not drop shape or invent zero; default failure also diagnosed',()=>{
  const s=shape(10);Object.defineProperty(s.TextFrame.TextRange,'Font',{get(){throw Error('host getter');}});
  for(const params of [{},{includeStyle:true}]){
    const r=read([s],params);assert.equal(r.shapes.length,1);assert.equal(r.shapes[0].shapeId,10);
    assert.equal(r.shapes[0].fontSize,null);assert.match(r.message,/缺失/);
    assert.ok(r.shapes[0].readErrors.some((x:any)=>x.field==='fontSize'));
    if((params as any).includeStyle){assert.equal(r.shapes[0].style.values.font.size,null);assert.ok(r.shapes[0].style.unavailable.includes('font.size'));}
  }
});
test('group depth and global object budgets bound traversal including self cycles',()=>{
  const child=shape(30), group=shape(20,[child]),root=shape(10,[group]);
  let r=read([root],{recursive:true,maxDepth:1});assert.equal(r.readSummary.visited,2);assert.equal(r.shapes[0].children[0].childrenTruncated,true);
  r=read([root,shape(40)],{recursive:true,maxObjects:2});assert.equal(r.readSummary.visited,2);assert.equal(r.readSummary.truncated,true);
  const cycle=shape(50,[]);cycle.GroupItems=collection([cycle]);r=read([cycle],{recursive:true,maxDepth:2});assert.equal(r.readSummary.visited,3);
  r=read([root],{recursive:true,maxDepth:0});assert.equal(r.readSummary.visited,1);
  assert.throws(()=>read([root],{recursive:true,maxObjects:0}),/预算/);
});
test('unavailable group collection and object getter preserve diagnostics',()=>{
  const s=shape(10,[]);Object.defineProperty(s,'GroupItems',{get(){throw Error('denied');}});
  const r=read([s],{recursive:true});assert.equal(r.shapes[0].childrenUnavailable,true);assert.equal(r.readSummary.issues[0].reason,'groupUnavailable');
});
test('schema extension and actual gateway forwarding; save/export descriptions match parameters',async()=>{
  const defs=pptToolDefinitions();const tool=defs.find(x=>x.function.name==='wps_ppt_get_slide_shapes')!.function;
  assert.equal(tool.parameters.properties.maxDepth.maximum,8);assert.equal(tool.parameters.properties.maxObjects.maximum,2000);
  const args={presentationName:'合成.pptx',slideIndex:1,includeStyle:true,recursive:true,maxDepth:2,maxObjects:30};
  let received:any;
  await getSlideShapes({args,callOffice:async(method:any,params:any)=>{received={method,params};return {};}} as any);
  assert.equal(received.method,'ppt_get_slide_shapes');assert.deepEqual(received.params,args);
  await manageSlides({args:{presentationName:'合成.pptx',action:'save_as',filePath:'/tmp/sample.pdf',format:'pdf'},callOffice:async(_:any,p:any)=>{received=p;return {};}} as any);
  assert.equal(received.filePath,'/tmp/sample.pdf');assert.equal(received.format,'pdf');
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'ppt-preview-'));const file=path.join(dir,'sample.png');
  try {
    const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=','base64');
    const r:any=await captureSlidePreview({args:{presentationName:'合成.pptx',slideIndex:1,outputPath:file},callOffice:async(_:any,p:any)=>{received=p;fs.writeFileSync(p.outputPath,png);return {success:true};}} as any);
    assert.equal(received.outputPath,file);assert.equal(r.imageBase64,png.toString('base64'));
  } finally {fs.rmSync(dir,{recursive:true,force:true});}
  assert.doesNotMatch(defs.find(x=>x.function.name==='wps_ppt_manage_slides')!.function.parameters.properties.filePath.description,/尚未透传/);
  assert.doesNotMatch(defs.find(x=>x.function.name==='wps_ppt_capture_slide_preview')!.function.parameters.properties.outputPath.description,/另一|额外导出/);
});

test('anchored PPT contract changes only optional read fields and corrected save/preview descriptions',async()=>{
  const before=JSON.parse(fs.readFileSync(fixtureRoot+'ppt-tools.before.json','utf8'));
  const after=JSON.parse(JSON.stringify(pptToolDefinitions()));
  assert.deepEqual(after.map((t:any)=>t.function.name),before.map((t:any)=>t.function.name));
  for(const tool of after){
    const f=tool.function;
    if(f.name==='wps_ppt_get_slide_shapes')for(const key of ['includeStyle','recursive','maxDepth','maxObjects'])delete f.parameters.properties[key];
    const old=before.find((t:any)=>t.function.name===f.name).function;
    if(f.name==='wps_ppt_manage_slides')for(const key of ['filePath','format'])f.parameters.properties[key].description=old.parameters.properties[key].description;
    if(f.name==='wps_ppt_capture_slide_preview')f.parameters.properties.outputPath.description=old.parameters.properties.outputPath.description;
  }
  assert.deepEqual(after,before,'no unrelated PPT contract change');
});

test('missing geometry remains unknown and failed child fetches consume budget',()=>{
  const s:any=shape(10);s.Left=null;s.Width=undefined;
  const r=read([s],{includeStyle:true});assert.equal(r.shapes[0].left,null);assert.equal(r.shapes[0].width,null);
  const group:any=shape(20,[]);let attempts=0;group.GroupItems={Count:100000,Item:()=>{attempts++;throw Error('unreadable');}};
  const result=read([group],{recursive:true,maxObjects:3});assert.equal(attempts,2);assert.equal(result.readSummary.visited,3);assert.equal(result.readSummary.truncated,true);
  assert.equal(result.shapes[0].children.length,2);assert.equal(result.shapes[0].children[0].shapeIndex,1);
});

test('rebuilt minified addon RPC returns extended PPT fields without host writes',async()=>{
  const shapes:any=collection([shape(10,[shape(20)])]);shapes.Placeholders=collection([]);
  const pres={Name:'合成.pptx',Slides:collection([{SlideID:101,Shapes:shapes,Layout:12,CustomLayout:{Name:'样本'}}]),PageSetup:{SlideWidth:1000,SlideHeight:600}};
  const app={ActivePresentation:pres,Presentations:{Count:1,Item:()=>pres}};
  let socket:any;const sent:any[]=[];
  class Socket{static OPEN=1;static CONNECTING=0;readyState=1;onmessage:any;onopen:any;constructor(_:any){socket=this;}send(text:string){sent.push(JSON.parse(text));}}
  const ctx=vm.createContext({window:{WPS_BRIDGE_CONFIG:{port:12345,token:'fake-only'},addEventListener:()=>{}},document:{getElementById:()=>null,readyState:'complete'},wps:{WppApplication:()=>app},WebSocket:Socket,console:{log:()=>{}},setTimeout:()=>0,clearTimeout:()=>{},setInterval:()=>0,alert:()=>{},confirm:()=>false});
  vm.runInContext(fs.readFileSync('wps-addon/addon-core.js','utf8'),ctx);socket.onopen();
  socket.onmessage({data:JSON.stringify({id:'ppt-test',method:'ppt_get_slide_shapes',params:{presentationName:'合成.pptx',slideIndex:1,includeStyle:true,recursive:true,maxObjects:3}})});
  await new Promise(resolve=>setImmediate(resolve));
  const response=sent.find(x=>x.type==='rpc_response'&&x.id==='ppt-test');assert.equal(response.error,null);
  assert.equal(response.result.slideId,101);assert.equal(response.result.shapes[0].children[0].style.values.font.nameFarEast,'示例字体');
});

test('extended missing group count/table/chart/rotation are unknown while legacy defaults remain',()=>{
  for(const count of [null,undefined,'']){
    const s:any=shape(10,[]);s.GroupItems={Count:count,Item:()=>{throw Error('must not fetch');}};
    const r=read([s],{recursive:true});assert.equal(r.shapes[0].childrenUnavailable,true);
    assert.ok(r.readSummary.issues.some((x:any)=>x.reason==='groupUnavailable'));
  }
  const s:any=shape(20);s.HasTable=undefined;s.HasChart=null;s.Rotation=undefined;
  const r=read([s],{includeStyle:true});
  for(const key of ['hasTable','hasChart','rotation'])assert.equal(r.shapes[0][key],null);
  const legacy=read([s]);assert.equal(legacy.shapes[0].hasTable,false);assert.equal(legacy.shapes[0].rotation,0);
});


test('baseline fixtures are pinned and require no git history',()=>{
  assert.equal(fixtureMetadata.baselineCommit,'968281789b5594494003292c1b68dcb6819f3e23');
  for(const [name,hash] of Object.entries(fixtureMetadata.sha256)){
    const text=fs.readFileSync(fixtureRoot+name,'utf8');
    const normalized=(value:string)=>createHash('sha256').update(value.replace(/\r\n/g,'\n'),'utf8').digest('hex');
    assert.equal(normalized(text),hash);
    assert.equal(normalized(text.replace(/\r?\n/g,'\r\n')),hash,'Windows CRLF checkout preserves the same text fixture');
  }
});

test('extended unknown text and z-order remain unknown rather than empty text/index',()=>{
  for(const text of [null,undefined]){
    const s:any=shape(10);s.TextFrame.TextRange.Text=text;s.ZOrderPosition=undefined;
    const r=read([s],{recursive:true});const x=r.shapes[0];
    assert.equal(x.hasText,true);assert.equal(x.text,null);assert.equal(x.textTruncated,null);assert.equal(x.zOrderPosition,null);
    assert.ok(x.readErrors.some((e:any)=>e.field==='text'));assert.ok(x.readErrors.some((e:any)=>e.field==='zOrderPosition'));assert.match(r.message,/缺失/);
  }
});
