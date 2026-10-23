import test from 'node:test';
import { layoutFailure } from './fixtures/ppt-style/layout-failure.js';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { once } from 'node:events';
import { WebSocket } from 'ws';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
const expectedVersion = JSON.parse(fs.readFileSync('package.json', 'utf8')).version;
const root=fs.mkdtempSync(path.join(os.tmpdir(),'bridge-service-'));
process.env.WPS_BRIDGE_HOME=root;process.env.WPS_BRIDGE_PORT=String(22000+Math.floor(Math.random()*10000));
const {bridgeServer}=await import('../src/bridge/ws-server.js');
const {composeBridgeServer}=await import('../src/bridge/compose.js');
const {getToken}=await import('../src/bridge/runtime.js');
const {auditStore}=await import('../src/bridge/audit-store.js');
const port=Number(process.env.WPS_BRIDGE_PORT),base=`http://127.0.0.1:${port}`;
let ws:WebSocket,token:string;
const headers=()=>({Authorization:`Bearer ${token}`,'Content-Type':'application/json'});
let values:any[][]=[[7]],formulas:any[][]=[[7]],writes=0;
async function call(name:string,args:any={}) {const r=await fetch(base+'/api/v1/tool/call',{method:'POST',headers:headers(),body:JSON.stringify({name,arguments:args})});return {status:r.status,result:await r.json() as any};}
test.before(async()=>{
  composeBridgeServer();   // 组装入口：注入工具服务（P2.2）
  await bridgeServer.start();token=getToken();
  ws=new WebSocket(`ws://127.0.0.1:${port}/addon?token=${token}`);await once(ws,'open');
  ws.on('message',raw=>{const p=JSON.parse(raw.toString());let result:any;
    if(p.method==='get_workspace_summary')result={workbookName:'Test.xlsx',hasOpenWorkbook:true};
    else if(p.method==='read_range')result={sheetName:'Sheet1',address:'$A$1',rowCount:1,columnCount:1,values,formulas};
    else if(p.method==='rollback_cells'){writes++;values=p.params.snapshot.values;formulas=p.params.snapshot.formulas;result={success:true};}
    else result={success:true};
    ws.send(JSON.stringify({type:'rpc_response',id:p.id,result}));
  });
  ws.send(JSON.stringify({type:'register',client:'wps-et-addon',version:'2.0.0',summary:{workbookName:'Test.xlsx',activeSheetName:'Sheet1'}}));
  await new Promise(r=>setTimeout(r,40));
});
test.after(()=>{ws?.terminate();bridgeServer.stop();});
test('loopback health is public, documents require auth, foreign origins rejected',async()=>{
  assert.equal((await fetch(base+'/health')).status,200);
  assert.equal((await fetch(base+'/api/v1/status')).status,401);
  assert.equal((await fetch(base+'/api/v1/tool/call',{method:'OPTIONS',headers:{Origin:'null','Access-Control-Request-Headers':'Authorization, Content-Type'}})).status,204);
  assert.equal((await fetch(base+'/api/v1/tool/call',{method:'POST',headers:{Origin:'null','Content-Type':'application/json'},body:'{}'})).status,401);
  assert.equal((await fetch(base+'/api/v1/status',{headers:headers()})).status,200);
  assert.equal((await fetch(base+'/api/v1/status',{headers:{...headers(),Origin:'https://evil.example'}})).status,403);
});
test('schemas reject malformed writes before reaching the host',async()=>{
  const result=await call('excel_patch_cells',{host:'wps',address:'A1',values:[[1],[2,3]]});assert.equal(result.status,422);assert.equal(writes,0);
});
test('Office 图标资源允许缓存，HTML/JS 仍禁止缓存，图标响应是实际 PNG',async()=>{
  for (const size of [16,20,24,32,48,64,80]) {
    const r=await fetch(base+`/assets/icon-${size}.png?v=charfloat-20261003`);
    assert.equal(r.status,200);
    assert.match(r.headers.get('content-type')||'',/image\/png/);
    assert.equal(r.headers.get('cache-control'),null);
    assert.equal(r.headers.get('pragma'),null);
    assert.equal(r.headers.get('expires'),null);
    assert.ok(r.headers.get('etag'));
    const data=Buffer.from(await r.arrayBuffer());
    assert.equal(data.subarray(1,4).toString(),'PNG');
    assert.ok(data.equals(fs.readFileSync(`office-addon/public/assets/icon-${size}.png`)));
  }
  for (const name of ['taskpane.html','taskpane.js','taskpane.css']) {
    const r=await fetch(base+'/office-addon/'+name);
    assert.equal(r.status,200);
    assert.match(r.headers.get('cache-control')||'',/no-store/);
  }
});
test('rollback refuses later edits and restores only matching snapshot',async()=>{
  const snapshot=(v:number)=>({sheetName:'Sheet1',address:'$A$1',rowCount:1,columnCount:1,values:[[v]],formulas:[[v]]});
  const rec=auditStore.addRecord({host:'wps',workbookName:'Test.xlsx',sheetName:'Sheet1',address:'$A$1',actionType:'update_values',description:'test',patchResult:{sheetName:'Sheet1',address:'$A$1',modifiedCount:1,diff:[],beforeSnapshot:snapshot(2),afterSnapshot:snapshot(6)}});
  assert.equal((await call('wps_rollback',{auditId:rec.id})).status,422);assert.equal(writes,0);
  values=[[6]];formulas=[[6]];
  assert.equal((await call('wps_rollback',{auditId:rec.id})).status,200);assert.equal(writes,1);assert.deepEqual(values,[[2]]);
});
test('HTTP MCP handshake, resources and real RPC routing',async()=>{
  const client=new Client({name:'integration',version:'1.0'});const transport=new StreamableHTTPClientTransport(new URL(base+'/mcp'),{requestInit:{headers:headers()}});
  await client.connect(transport);const tools=await client.listTools();assert.ok(tools.tools.some(t=>t.name==='excel_read_range'));
  const result:any=await client.callTool({name:'excel_get_workspace_summary',arguments:{host:'wps'}});assert.equal(JSON.parse(result.content[0].text).workbookName,'Test.xlsx');
  const resources=await client.readResource({uri:'bridge://capabilities'});assert.match((resources.contents[0] as any).text,/Windows/);
  await transport.terminateSession();await client.close();
});
test('two stdio clients reuse service; exiting one preserves the other and daemon',async()=>{
  const env={...process.env} as Record<string,string>;
  const a=new Client({name:'stdio-a',version:'1'}),b=new Client({name:'stdio-b',version:'1'});
  const ta=new StdioClientTransport({command:process.execPath,args:[path.resolve('dist/bridge/cli.cjs')],env,stderr:'pipe'});
  const tb=new StdioClientTransport({command:process.execPath,args:[path.resolve('dist/bridge/cli.cjs')],env,stderr:'pipe'});
  try{
    await Promise.all([a.connect(ta),b.connect(tb)]);
    const read=await a.callTool({name:'excel_get_workspace_summary',arguments:{host:'wps'}});assert.equal((read as any).isError,undefined);
    await a.close();
    const next:any=await b.callTool({name:'bridge_get_capabilities',arguments:{}});assert.equal(JSON.parse(next.content[0].text).version,expectedVersion);
    assert.equal((await fetch(base+'/health')).status,200);
  }finally{await a.close();await b.close();}
});

test('async tool schema and failed/unknown evidence survive actual HTTP and MCP SDK',async()=>{
  const original=(bridgeServer as any).tools();
  const tools=[{name:'test_partial_failure',description:'synthetic failure',inputSchema:{type:'object',properties:{}}}];
  let calls=0;
  bridgeServer.setToolService({
    list:async()=>tools,
    capabilities:async()=>({}),
    execute:async()=>{calls++;return {success:false,error:'synthetic partial failure',errorDetails:{kind:'failed',channel:'wps-addon',executed:'unknown'},partial:{completedShapeIds:[101],document:'must-not-leak'}};}
  });
  const client=new Client({name:'safe-failure-test',version:'1'});
  const transport=new StreamableHTTPClientTransport(new URL(base+'/mcp'),{requestInit:{headers:headers()}});
  try{
    const schemaResponse=await fetch(base+'/api/v1/mcp-tools',{headers:headers()});
    assert.deepEqual(await schemaResponse.json(),tools);
    const http=await call('test_partial_failure');
    assert.equal(http.status,422);assert.equal(http.result.errorDetails.executed,'unknown');
    assert.deepEqual(http.result.partial,{completedShapeIds:[101]});
    await client.connect(transport);
    const mcp:any=await client.callTool({name:'test_partial_failure',arguments:{}});
    assert.equal(mcp.isError,true);
    const failure=JSON.parse(mcp.content[1].text);
    assert.equal(failure.errorDetails.kind,'failed');assert.equal(failure.errorDetails.executed,'unknown');
    assert.deepEqual(failure.partial,{completedShapeIds:[101]});assert.doesNotMatch(JSON.stringify(failure),/must-not-leak/);
    assert.equal(calls,2,'one explicit HTTP call and one explicit MCP call; no retry');
  }finally{
    await transport.terminateSession().catch(()=>{});await client.close();bridgeServer.setToolService(original);
  }
});


test('actual host layout partial failure retains readback through HTTP and real MCP SDK',async()=>{
  const original=(bridgeServer as any).tools();
  let calls=0;
  bridgeServer.setToolService({list:()=>[{name:'wps_ppt_configure_layout',description:'synthetic',inputSchema:{type:'object',properties:{}}}],capabilities:()=>({}),execute:async()=>{calls++;return layoutFailure();}});
  const client=new Client({name:'layout-failure-test',version:'1'});
  const stdio=new Client({name:'layout-failure-stdio',version:'1'});
  const stdioTransport=new StdioClientTransport({command:process.execPath,args:[path.resolve('dist/bridge/cli.cjs')],env:{...process.env} as Record<string,string>,stderr:'pipe'});
  const transport=new StreamableHTTPClientTransport(new URL(base+'/mcp'),{requestInit:{headers:headers()}});
  function verify(result:any){
    assert.equal(result.errorDetails.executed,'unknown');assert.equal(result.applied.slideWidth,1000);
    assert.equal(result.before.slideWidth,900);assert.equal(result.after.slideHeight,600);
    assert.equal(result.warningDetails[0].field,'slideHeight');assert.equal(result.warningDetails[0].reason,'host_rejected');
    assert.doesNotMatch(JSON.stringify(result),/token=secret|private-document/);
  }
  try{
    const http=await call('wps_ppt_configure_layout');assert.equal(http.status,422);verify(http.result);
    await client.connect(transport);
    const result:any=await client.callTool({name:'wps_ppt_configure_layout',arguments:{}});assert.equal(result.isError,true);verify(JSON.parse(result.content[1].text));
    await stdio.connect(stdioTransport);
    const viaStdio:any=await stdio.callTool({name:'wps_ppt_configure_layout',arguments:{}});
    assert.equal(viaStdio.isError,true);verify(JSON.parse(viaStdio.content[1].text));
    assert.equal(calls,3,'HTTP, MCP HTTP and stdio each explicitly called once');
  }finally{await stdio.close();await transport.terminateSession().catch(()=>{});await client.close();bridgeServer.setToolService(original);}
});
