// 已授权的真实宿主验收控制台；通过独立 stdio CLI 调用，仅按输入逐条执行，不自动重放。
import fs from 'node:fs';import path from 'node:path';import readline from 'node:readline';import {fileURLToPath} from 'node:url';
import {Client} from '@modelcontextprotocol/sdk/client/index.js';import {StdioClientTransport} from '@modelcontextprotocol/sdk/client/stdio.js';
const root=fileURLToPath(new URL('../../../../',import.meta.url)),dir=path.dirname(fileURLToPath(import.meta.url));
const runtime=fs.readFileSync('/tmp/charfloat-m1h-runtime-path','utf8').trim();
const command=path.join(root,'node_modules/electron/dist/Electron.app/Contents/MacOS/Electron');
const client=new Client({name:'CharFloat M1-H Acceptance',version:'1.0.0'});
const transport=new StdioClientTransport({command,args:[path.join(root,'dist/bridge/cli.cjs')],cwd:root,env:{...process.env,ELECTRON_RUN_AS_NODE:'1',WPS_BRIDGE_HOME:runtime,WPS_BRIDGE_PORT:'19890',WPS_BRIDGE_RESOURCES:root},stderr:'inherit'});
await client.connect(transport);
const tools=await client.listTools();fs.writeFileSync(path.join(dir,'cli-tools.json'),JSON.stringify(tools,null,2)+'\n');
console.log(JSON.stringify({ready:true,pid:process.pid,toolCount:tools.tools.length,transport:'独立 stdio CLI → 独立后台 → WPS'}));
const concise=value=>typeof value==='string'&&value.length>4000?`<长字段 ${value.length} 字符，见证据文件>`:Array.isArray(value)?value.map(concise):value&&typeof value==='object'?Object.fromEntries(Object.entries(value).map(([k,v])=>[k,concise(v)])):value;
let count=0;
for await(const line of readline.createInterface({input:process.stdin,terminal:false})){
 if(!line.trim())continue;
 const input=JSON.parse(line);if(input.quit){await client.close();break;}
 const start=new Date().toISOString(),tag=String(input.tag||++count).replace(/[^a-zA-Z0-9_-]/g,'_');
 try{
  const response=await client.callTool({name:input.tool,arguments:input.args||{}},undefined,{timeout:60000});
  let data;try{data=JSON.parse(response.content.find(x=>x.type==='text')?.text||'null');}catch{data=response.content;}
  const record={tool:input.tool,arguments:input.args||{},startUTC:start,endUTC:new Date().toISOString(),isError:response.isError===true,data};
  fs.writeFileSync(path.join(dir,tag+'.json'),JSON.stringify(record,null,2)+'\n');
  console.log(JSON.stringify({tag,isError:record.isError,data:concise(data)}));
 }catch(e){const record={tool:input.tool,startUTC:start,endUTC:new Date().toISOString(),transportError:e.message,executed:'unknown; must read back before any replay'};fs.writeFileSync(path.join(dir,tag+'.json'),JSON.stringify(record,null,2)+'\n');console.log(JSON.stringify({tag,...record}));}
}

await client.close();
