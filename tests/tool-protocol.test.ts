/** Executes real protocol modules with injected transports/runtime; no sockets, credentials or SDK installation. */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { createRequire } from 'node:module';
import { Readable } from 'node:stream';
import { execFileSync } from 'node:child_process';
import ts from 'typescript';
import * as errors from '../src/bridge/errors.js';
const require = createRequire(import.meta.url);
function load(file:string, mocks:Record<string,unknown>, globals:Record<string,unknown>={}) {
  const code=ts.transpileModule(fs.readFileSync(file,'utf8'),{compilerOptions:{target:ts.ScriptTarget.ES2022,module:ts.ModuleKind.CommonJS,esModuleInterop:true}}).outputText;
  const exports:any={};
  vm.runInNewContext(code,{exports,require:(id:string)=>{
    if(id in mocks)return mocks[id];
    if(id.startsWith('node:')||id==='ws')return require(id);
    throw Error('unexpected dependency '+id);
  },Buffer,URL,console,process,setTimeout,clearTimeout,setInterval,clearInterval,...globals},{filename:file});
  return exports;
}
const runtime={runtimePort:()=>19890,runtimeHttpsPort:()=>19891,validToken:(t:string)=>t==='test-only',VERSION:'test',PROTOCOL:2};
function httpModule(){return load('src/bridge/ws-server.ts',{
  './errors.js':errors,'./runtime.js':runtime,'./context.js':{},'./build-fingerprint.js':{},'./mcp-server.js':{},
  '@modelcontextprotocol/sdk/server/sse.js':{},'@modelcontextprotocol/sdk/server/streamableHttp.js':{}
});}
async function callHttp(service:any,url='/api/v1/tool/call',method='POST'){
  const {WpsBridgeServer}=httpModule();const server=new WpsBridgeServer();server.setToolService(service);
  const req:any=Readable.from([Buffer.from(JSON.stringify({name:'wps_ppt_manage_slides',arguments:{presentationName:'合成.pptx'},sessionId:'task-1'}))]);
  req.url=url;req.method=method;req.headers={host:'127.0.0.1:19890',authorization:'Bearer test-only'};
  let status=0,body='';const res={writeHead:(s:number)=>status=s,end:(s:string)=>body=s,setHeader:()=>{}};
  await server.handleHttp(req,res);return {status,body:JSON.parse(body)};
}
test('HTTP real handler preserves failed/unknown and partial IDs; Python consumes same wire response once',async()=>{
  let count=0;
  const r=await callHttp({execute:async(_:any,__:any,ctx:any)=>{count++;assert.equal(ctx.sessionId,'task-1');return {success:false,error:'宿主失败',errorDetails:{kind:'failed',channel:'wps-addon',executed:'unknown'},partial:{completedShapeIds:[10],document:'private'}};}});
  assert.equal(count,1);assert.equal(r.status,422);assert.equal(r.body.errorDetails.executed,'unknown');assert.deepEqual(r.body.partial,{completedShapeIds:[10]});
  const output=execFileSync('python3',['-B','-c',`import importlib.util,json,sys,io,urllib.error\nfrom unittest.mock import patch\ns=importlib.util.spec_from_file_location('bridge','skills/charfloat/scripts/bridge_client.py');b=importlib.util.module_from_spec(s);s.loader.exec_module(b)\nwire=json.loads(sys.stdin.read())\nerror=urllib.error.HTTPError('http://127.0.0.1',422,'failed',{},io.BytesIO(json.dumps(wire).encode()))\nwith patch.object(b.Path,'read_text',return_value='fake-only'),patch.object(b.urllib.request,'urlopen',side_effect=error) as m:\n try:b.request('/api/v1/tool/call',{'name':'tool'})\n except b.BridgeFailure as e:\n  assert m.call_count==1\n  print(json.dumps(e.detail))`],{input:JSON.stringify(r.body),encoding:'utf8',env:{...process.env,PYTHONDONTWRITEBYTECODE:'1'}});
  const cli=JSON.parse(output);assert.equal(cli.executed,'unknown');assert.deepEqual(cli.partial,{completedShapeIds:[10]});assert.match(cli.recovery,/读回/);
});
test('HTTP thrown unclassified error does not claim host not executed',async()=>{
  const r=await callHttp({execute:async()=>{throw new Error('timeout');}});
  assert.equal(r.status,422);assert.equal(r.body.errorDetails.kind,'unknown');assert.equal(r.body.errorDetails.executed,'unknown');
});
test('MCP real callback keeps isError and readable text plus safe structured failure',async()=>{
  class Server{handlers=new Map();setRequestHandler(k:any,v:any){this.handlers.set(k,v);}}
  const types={CallToolRequestSchema:'call',ListToolsRequestSchema:'list',ListPromptsRequestSchema:'prompts',GetPromptRequestSchema:'prompt',ListResourcesRequestSchema:'resources',ReadResourceRequestSchema:'resource'};
  const {createMcpServer}=load('src/bridge/mcp-server.ts',{'./errors.js':errors,'./runtime.js':runtime,'@modelcontextprotocol/sdk/server/index.js':{Server},'@modelcontextprotocol/sdk/types.js':types});
  for(const failure of [new errors.BridgeError('failed','宿主失败',{channel:'wps-addon'}),{success:false,error:'部分完成',partial:{completedCount:1,completedShapeIds:[3]}}]){
    let calls=0;
    const server=createMcpServer({execute:async()=>{calls++;if(failure instanceof Error)throw failure;return failure;}});
    const r=await server.handlers.get('call')({params:{name:'tool',arguments:{}}});
    assert.equal(calls,1);assert.equal(r.isError,true);assert.equal(r.content[0].type,'text');
    const detail=JSON.parse(r.content[1].text);assert.equal(detail.errorDetails.executed,'unknown');
    if(!(failure instanceof Error))assert.deepEqual(detail.partial,{completedCount:1,completedShapeIds:[3]});
  }
});

test('background service client keeps safe failure fields for stdio MCP without retry',async()=>{
  let calls=0;
  const neverStart=()=>{throw new Error('protocol failure forwarding must not start or pause a service');};
  const client=load('src/bridge/service-client.ts',{'./errors.js':errors,'./runtime.js':{...runtime,getToken:()=> 'test-only'},'./service-start-control.js':{captureServiceStartIntent:neverStart,commitServiceStartIntent:neverStart,pauseServiceStarts:neverStart}},{AbortSignal,fetch:async()=>{calls++;return {ok:false,status:422,json:async()=>({success:false,error:'宿主失败',errorDetails:{kind:'failed',channel:'wps-addon',executed:'unknown'},partial:{completedShapeIds:[3]}})};}});
  try{await client.serviceRequest('/api/v1/tool/call',{name:'tool'});assert.fail('must fail');}
  catch(e:any){assert.equal(e.errorDetails.executed,'unknown');assert.deepEqual(JSON.parse(JSON.stringify(e.partial)),{completedShapeIds:[3]});}
  assert.equal(calls,1);
});

test('existing HTTP schema endpoint awaits async ToolService.list and CLI receives the array',async()=>{
  const tools=[{name:'wps_ppt_get_slide_shapes',description:'读取',inputSchema:{type:'object'}}];
  const response=await callHttp({list:async()=>tools},'/api/v1/mcp-tools','GET');
  assert.equal(response.status,200);assert.deepEqual(response.body,tools);
  const out=execFileSync('python3',['-B','-c',`import importlib.util,json,sys
from unittest.mock import patch
s=importlib.util.spec_from_file_location('bridge','skills/charfloat/scripts/bridge_client.py');b=importlib.util.module_from_spec(s);s.loader.exec_module(b)
wire=json.loads(sys.stdin.read())
with patch.object(b,'request',return_value=wire) as m:
 b.main(['--schema','wps_ppt_get_slide_shapes'])
 assert m.call_count==1
 assert m.call_args.args[0]=='/api/v1/mcp-tools'`],{input:JSON.stringify(response.body),encoding:'utf8'});
  assert.deepEqual(JSON.parse(out),tools[0]);
});
