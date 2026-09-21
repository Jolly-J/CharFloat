import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs';
function load() {
  let socket:any;const sent:any[]=[];
  class Socket {
    static OPEN=1;static CONNECTING=0;readyState=1;onmessage:any;onopen:any;
    constructor(public url:string){socket=this}send(text:string){sent.push(JSON.parse(text))}
  }
  let value:any=0,formula:any=0;
  const range:any={Rows:{Count:1},Columns:{Count:1},Address:()=>'$A$1',Cells:{Item:()=>({Address:()=>'$A$1',NumberFormat:'0'})}};
  Object.defineProperty(range,'Value2',{get:()=>value,set:v=>{value=v;formula=v}});
  Object.defineProperty(range,'Formula',{get:()=>formula,set:v=>{value=v;formula=v}});
  const sheet:any={Name:'Test',Visible:-1,Range:()=>range};
  const book:any={Name:'Exact.xlsx',FullName:'/tmp/Exact.xlsx',ActiveSheet:sheet,Worksheets:{Count:1,Item:()=>sheet}};sheet.Parent=book;
  const app:any={ActiveWorkbook:book,Workbooks:{Count:1,Item:(id:any)=>{if(id===1||id==='Exact.xlsx')return book;throw new Error('not found')}}};
  const context=vm.createContext({window:{WPS_BRIDGE_CONFIG:{port:12345,token:'test-token'},addEventListener:()=>{}},document:{getElementById:()=>null,readyState:'complete'},wps:{EtApplication:()=>app},WebSocket:Socket,console:{log:()=>{}},setTimeout:()=>0,clearTimeout:()=>{},setInterval:()=>0,alert:()=>{},confirm:()=>false});
  vm.runInContext(fs.readFileSync('wps-addon/addon-core.js','utf8'),context);socket.onopen();
  return {socket,sent,read:()=>value,call:async(method:string,params:any)=>{socket.onmessage({data:JSON.stringify({id:'test',method,params})});await new Promise(r=>setImmediate(r));return sent.findLast(x=>x.type==='rpc_response')}};
}
test('addon uses installed credentials and zero-valued formulas survive reads',async()=>{
  const host=load();assert.match(host.socket.url,/\/addon\?token=test-token/);
  const result=await host.call('read_range',{address:'A1',workbookName:'Exact.xlsx',includeFormulas:true});
  assert.deepEqual(result.result.values,[[0]]);assert.deepEqual(result.result.formulas,[[0]]);
});
test('addon rejects wrong dimensions before writing and rejects partial document names',async()=>{
  const host=load();const result=await host.call('patch_cells',{address:'A1',values:[[1,2]],workbookName:'Exact.xlsx'});assert.match(result.error,/尺寸一致/);assert.equal(host.read(),0);
  const missing=await host.call('read_range',{address:'A1',workbookName:'Exact'});assert.match(missing.error,/未在 WPS/);
});
