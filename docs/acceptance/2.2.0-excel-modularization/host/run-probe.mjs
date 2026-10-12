// 限定为本批已有测试请求文件，避免终端长行截断；不自动重放。
import fs from 'node:fs';import path from 'node:path';import {fileURLToPath} from 'node:url';
import {Client} from '@modelcontextprotocol/sdk/client/index.js';import {StdioClientTransport} from '@modelcontextprotocol/sdk/client/stdio.js';
const root=fileURLToPath(new URL('../../../../',import.meta.url)),dir=path.dirname(fileURLToPath(import.meta.url));
const allowed={view:'/tmp/charfloat-baseline-view-request.json',chart:'/tmp/charfloat-baseline-chart-request.json',inactive:'/tmp/charfloat-baseline-inactive-chart-request.json',legend:'/tmp/charfloat-baseline-no-legend-request.json',exact:'/tmp/charfloat-baseline-exact-request.json'};
const file=allowed[process.argv[2]];if(!file)throw Error('仅允许本批基线探针');
const request=JSON.parse(fs.readFileSync(file));
if(!['CharFloat_M1H_20261012.xlsx','CharFloat_M1H_ViewControl_20261012.xlsx'].includes(request.args.workbookName))throw Error('目标保护');
const runtime=fs.readFileSync('/tmp/charfloat-m1h-runtime-path','utf8').trim();
const client=new Client({name:'CharFloat M1-H Baseline Probe',version:'1.0.0'});
await client.connect(new StdioClientTransport({command:path.join(root,'node_modules/electron/dist/Electron.app/Contents/MacOS/Electron'),args:[path.join(root,'dist/bridge/cli.cjs')],cwd:root,env:{...process.env,ELECTRON_RUN_AS_NODE:'1',WPS_BRIDGE_HOME:runtime,WPS_BRIDGE_PORT:'19890',WPS_BRIDGE_RESOURCES:root},stderr:'inherit'}));
try{const result=await client.callTool({name:request.tool,arguments:request.args},undefined,{timeout:60000});let data;const text=result.content.find(c=>c.type==='text')?.text;try{data=JSON.parse(text);}catch{data={error:text};}
const record={source:'固定基线函数，独立CLI实际宿主调用',tool:request.tool,arguments:request.args,isError:result.isError===true,data,atUTC:new Date().toISOString()};fs.writeFileSync(path.join(dir,request.tag+'.json'),JSON.stringify(record,null,2)+'\n');console.log(JSON.stringify({tag:request.tag,isError:record.isError,data}));}finally{await client.close();}
