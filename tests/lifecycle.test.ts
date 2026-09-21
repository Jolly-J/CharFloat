import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
const dir=fs.mkdtempSync(path.join(os.tmpdir(),'bridge-lifecycle-'));
const port=String(33000+Math.floor(Math.random()*12000));
const env={...process.env,WPS_BRIDGE_HOME:dir,WPS_BRIDGE_PORT:port,WPS_BRIDGE_RESOURCES:process.cwd()} as Record<string,string>;
test('cold stdio start launches a daemon; stdio exit leaves it running',async()=>{
  const client=new Client({name:'cold-start',version:'1'});
  const transport=new StdioClientTransport({command:process.execPath,args:[path.resolve('dist/bridge/cli.cjs')],env,stderr:'pipe'});
  const base=`http://127.0.0.1:${port}`;
  try{
    await client.connect(transport);
    const info:any=await (await fetch(base+'/health')).json();
    assert.equal(info.protocol,2);
    assert.ok(fs.existsSync(path.join(dir,'token')));
    await client.close();
    const later:any=await (await fetch(base+'/health')).json();assert.equal(info.pid,later.pid);
    const token=fs.readFileSync(path.join(dir,'token'),'utf8');
    const prompt=await fetch(base+'/api/v1/prompts/aesthetic',{headers:{Authorization:`Bearer ${token}`}});assert.equal(prompt.status,200);assert.match(await prompt.text(),/表格/);
  }finally{
    await client.close();
    if(fs.existsSync(path.join(dir,'token')))await fetch(base+'/api/v1/service/stop',{method:'POST',headers:{Authorization:`Bearer ${fs.readFileSync(path.join(dir,'token'),'utf8')}`}}).catch(()=>{});
  }
});
