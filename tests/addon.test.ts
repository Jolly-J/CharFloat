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
/**
 * 1 行 2 列的宿主模拟，**复刻 WPS 的真实行为**：给 `Formula` 赋空字符串会清空该单元格。
 * 这正是 `patch_cells` 同时传 values 与 formulas 时静默丢数据的原因（问题台账 ISS-01 / DP7）。
 */
function loadGrid(){
  let socket:any;const sent:any[]=[];
  class Socket {
    static OPEN=1;static CONNECTING=0;readyState=1;onmessage:any;onopen:any;
    constructor(public url:string){socket=this}send(text:string){sent.push(JSON.parse(text))}
  }
  const cells=[{value:'旧A',formula:'旧A'},{value:'旧B',formula:'旧B'}];
  const writeCell=(index:number,raw:any)=>{if(raw===''||raw===null||raw===undefined){cells[index].value=null;cells[index].formula=''}else{cells[index].value=raw;cells[index].formula=raw}};
  const cellObjects=cells.map((_cell,index)=>({Address:()=>['$A$1','$B$1'][index],NumberFormat:'0',get Formula(){return cells[index].formula},set Formula(raw:any){writeCell(index,raw)}}));
  const range:any={Rows:{Count:1},Columns:{Count:2},Address:()=>'$A$1:$B$1',Cells:{Item:(_row:number,column:number)=>cellObjects[column-1]}};
  Object.defineProperty(range,'Value2',{get:()=>cells.map(cell=>cell.value),set:matrix=>{matrix[0].forEach((raw:any,index:number)=>writeCell(index,raw))}});
  Object.defineProperty(range,'Formula',{get:()=>cells.map(cell=>cell.formula),set:matrix=>{matrix[0].forEach((raw:any,index:number)=>writeCell(index,raw))}});
  const sheet:any={Name:'Test',Visible:-1,Range:()=>range};
  const book:any={Name:'Exact.xlsx',FullName:'/tmp/Exact.xlsx',ActiveSheet:sheet,Worksheets:{Count:1,Item:()=>sheet}};sheet.Parent=book;
  const app:any={ActiveWorkbook:book,Workbooks:{Count:1,Item:(id:any)=>{if(id===1||id==='Exact.xlsx')return book;throw new Error('not found')}}};
  const context=vm.createContext({window:{WPS_BRIDGE_CONFIG:{port:12345,token:'test-token'},addEventListener:()=>{}},document:{getElementById:()=>null,readyState:'complete'},wps:{EtApplication:()=>app},WebSocket:Socket,console:{log:()=>{}},setTimeout:()=>0,clearTimeout:()=>{},setInterval:()=>0,alert:()=>{},confirm:()=>false});
  vm.runInContext(fs.readFileSync('wps-addon/addon-core.js','utf8'),context);socket.onopen();
  return {socket,cells:()=>cells.map(cell=>({...cell})),call:async(method:string,params:any)=>{socket.onmessage({data:JSON.stringify({id:'test',method,params})});await new Promise(r=>setImmediate(r));return sent.findLast(x=>x.type==='rpc_response')}};
}
test('patch_cells 的 formulas 空项表示"不改公式"，不得清空同批写入的值（ISS-01 / DP7）',async()=>{
  const host=loadGrid();
  const result=await host.call('patch_cells',{address:'A1:B1',values:[['新A','新B']],formulas:[['','']],workbookName:'Exact.xlsx'});
  assert.ok(result.result,result.error);
  // 修复前：整批赋 '' 会把刚写入的值清空，这里实际是 [null,null]
  assert.deepEqual(host.cells().map(cell=>cell.value),['新A','新B']);
});
test('patch_cells 混合矩阵：空项保留同批写入的值，非空项正常落公式',async()=>{
  const host=loadGrid();
  const result=await host.call('patch_cells',{address:'A1:B1',values:[['新A','新B']],formulas:[['=1+1','']],workbookName:'Exact.xlsx'});
  assert.ok(result.result,result.error);
  const cells=host.cells();
  assert.equal(cells[0].formula,'=1+1');
  assert.equal(cells[1].value,'新B');
});
test('patch_cells 无空项时仍走整批赋值，值写入行为不变',async()=>{
  const host=loadGrid();
  const result=await host.call('patch_cells',{address:'A1:B1',values:[['新A','新B']],workbookName:'Exact.xlsx'});
  assert.ok(result.result,result.error);
  assert.deepEqual(host.cells().map(cell=>cell.value),['新A','新B']);
});
