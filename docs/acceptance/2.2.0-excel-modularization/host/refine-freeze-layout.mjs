// 验收夹具收尾：严格核对冻结行列数，并将图表移到复制区下方，不改业务源码。
import fs from 'node:fs';import path from 'node:path';import assert from 'node:assert/strict';import {fileURLToPath} from 'node:url';
import {Client} from '@modelcontextprotocol/sdk/client/index.js';import {StdioClientTransport} from '@modelcontextprotocol/sdk/client/stdio.js';
const root=fileURLToPath(new URL('../../../../',import.meta.url)),dir=path.dirname(fileURLToPath(import.meta.url)),book='CharFloat_M1H_20261012.xlsx';
const runtime=fs.readFileSync('/tmp/charfloat-m1h-runtime-path','utf8').trim();const client=new Client({name:'CharFloat M1-H Freeze Readback',version:'1.0.0'});
await client.connect(new StdioClientTransport({command:path.join(root,'node_modules/electron/dist/Electron.app/Contents/MacOS/Electron'),args:[path.join(root,'dist/bridge/cli.cjs')],cwd:root,env:{...process.env,ELECTRON_RUN_AS_NODE:'1',WPS_BRIDGE_HOME:runtime,WPS_BRIDGE_PORT:'19890',WPS_BRIDGE_RESOURCES:root},stderr:'inherit'}));
async function call(name,args){const result=await client.callTool({name,arguments:{workbookName:book,...args}},undefined,{timeout:60000});const text=result.content.find(c=>c.type==='text')?.text;const data=JSON.parse(text);assert.ok(!result.isError,text);assert.notEqual(data.success,false);return data;}
try{
 await call('wps_freeze_panes',{sheetName:'验收数据',unfreeze:true});
 await call('wps_freeze_panes',{sheetName:'验收数据',freezeRowIndex:2,freezeColumnIndex:1});
 const freeze=await call('wps_freeze_panes',{sheetName:'验收数据',action:'read'});assert.equal(freeze.frozen,true);assert.equal(freeze.splitRow,1);assert.equal(freeze.splitColumn,0);
 const layout=await call('wps_execute_script',{component:'excel',readOnly:false,code:"const s=wb.Worksheets.Item('验收数据');s.Columns.Item(7).ColumnWidth=18;return {width:s.Columns.Item(7).ColumnWidth};"});
 await call('wps_update_chart',{sheetName:'验收数据',chartName:'Chart 1',position:{leftCell:'G8',width:350,height:190}});
 fs.writeFileSync(path.join(dir,'freeze-layout-final.json'),JSON.stringify({freeze,layout,reason:'freezeRowIndex 是首个未冻结行，2 代表冻结第 1 行；先解除已有分割。测试图表改用单系列双数据点，并移至复制区下方。'},null,2)+'\n');console.log(JSON.stringify({freeze,layout}));
}finally{await client.close();}
