import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
const root=fs.mkdtempSync(path.join(os.tmpdir(),'bridge-platform-'));
process.env.WPS_BRIDGE_HOME=path.join(root,'home');
process.env.WPS_BRIDGE_ADDON_DIR=path.join(root,'addons');
const {mergePluginIndex,AddonInstaller}=await import('../src/main/addon-installer.js');
const {mergeMcpConfig}=await import('../src/main/installer-engine.js');
const {getTools,validateArgs}=await import('../src/bridge/catalog.js');
const {requestContext}=await import('../src/bridge/context.js');
const {TargetLockStore}=await import('../src/bridge/gateway.js');
const {runProcess}=await import('../src/bridge/process-runner.js');

test('XML merge preserves unrelated plugins and rejects damaged input',()=>{
  const raw='<jsplugins><jsplugin name="Other" type="et" url="./other"/><jsplugin name="WPS Bridge" type="et"/><jsplugin name="Office Agent Bridge (表格)" type="et"/></jsplugins>';
  const result=mergePluginIndex(raw);
  assert.match(result,/name="Other"/);
  assert.equal((result.match(/name="字浮 CharFloat/g)||[]).length,3);
  assert.equal((result.match(/name="Office Agent Bridge/g)||[]).length,0);
  assert.equal((result.match(/name="WPS Bridge/g)||[]).length,0);
  assert.equal(mergePluginIndex(result),result);
  assert.throws(()=>mergePluginIndex('<jsplugins><broken>'));
  // 验证带有 BOM 和 Unicode 替换字符警告时不抛异常，正常合并
  const bomRaw='\uFEFF<jsplugins><jsplugin name="Test\uFFFDPlugin" type="et" url="./test"/></jsplugins>';
  const resBom=mergePluginIndex(bomRaw);
  assert.match(resBom,/name="字浮 CharFloat/);
  // 验证缺少根节点（如仅声明、仅注释、DOCTYPE、空关闭标签或异常符号）时，自动安全初始化为 <jsplugins/>
  for (const emptyRoot of [
    '<?xml version="1.0" encoding="utf-8"?>\r\n',
    '<!-- only comment -->',
    '<!DOCTYPE note SYSTEM "Note.dtd">',
    '</jsplugins>',
    '   <   '
  ]) {
    const res = mergePluginIndex(emptyRoot);
    assert.match(res, /name="字浮 CharFloat/);
  }
});
test('MCP merge preserves other clients and corrupt files byte for byte',()=>{
  const file=path.join(root,'config.json');fs.writeFileSync(file,'{"other":true,"mcpServers":{"existing":{"command":"x"}}}');
  mergeMcpConfig(file,{command:'new'});
  assert.deepEqual(JSON.parse(fs.readFileSync(file,'utf8')).mcpServers.existing,{command:'x'});
  assert.equal(JSON.parse(fs.readFileSync(file,'utf8')).other,true);
  fs.writeFileSync(file,'{broken');assert.throws(()=>mergeMcpConfig(file,{}));assert.equal(fs.readFileSync(file,'utf8'),'{broken');
});
test('MCP 配置改名：只写新名字 charfloat，并把遗留 office-agent-bridge / wps-bridge 条目迁走',()=>{
  const file=path.join(root,'rename.json');
  const entry={command:'/A.app/binary',args:['/A.app/cli.cjs']};
  // 场景一：已有我们的遗留条目（命令与参数一致）→ 应被迁走，只留新名字 charfloat
  fs.writeFileSync(file,JSON.stringify({mcpServers:{'wps-bridge':entry,'office-agent-bridge':entry,other:{command:'x'}}}));
  mergeMcpConfig(file,entry);
  const a=JSON.parse(fs.readFileSync(file,'utf8'));
  assert.deepEqual(Object.keys(a.mcpServers).sort(),['charfloat','other']);
  assert.deepEqual(a.mcpServers['charfloat'],entry);
  // 场景二：别人的同名条目（命令不同）→ **必须原样保留**，不能被我们删掉
  fs.writeFileSync(file,JSON.stringify({mcpServers:{'office-agent-bridge':{command:'someone-else'}}}));
  mergeMcpConfig(file,entry);
  const b=JSON.parse(fs.readFileSync(file,'utf8'));
  assert.deepEqual(b.mcpServers['office-agent-bridge'],{command:'someone-else'});
  assert.deepEqual(b.mcpServers['charfloat'],entry);
  // 幂等：再跑一次结果不变
  mergeMcpConfig(file,entry);
  assert.equal(fs.readFileSync(file,'utf8'),JSON.stringify(b,null,2)+'\n');
});
test('status checks do not create plugin directories; install validates before writes',()=>{
  assert.equal(fs.existsSync(process.env.WPS_BRIDGE_ADDON_DIR!),false);AddonInstaller.checkStatus();assert.equal(fs.existsSync(process.env.WPS_BRIDGE_ADDON_DIR!),false);
  fs.mkdirSync(process.env.WPS_BRIDGE_ADDON_DIR!,{recursive:true});fs.writeFileSync(path.join(process.env.WPS_BRIDGE_ADDON_DIR!,'publish.xml'),'<broken>');
  assert.equal(AddonInstaller.install().success,false);assert.equal(fs.existsSync(path.join(process.env.WPS_BRIDGE_ADDON_DIR!,'wps-bridge')),false);
});
test('tool names are unique and unified Excel host is mandatory',()=>{
  const tools=getTools();assert.equal(new Set(tools.map(t=>t.name)).size,tools.length);
  const schema=tools.find(t=>t.name==='excel_read_range')!.inputSchema;
  assert.throws(()=>validateArgs(schema,{address:'A1'}));
  validateArgs(schema,{host:'microsoft',address:'A1'});
  assert.throws(()=>validateArgs(schema,{host:'other',address:'A1'}));
  assert.throws(()=>validateArgs(schema,{host:'wps',address:'A1',unknown:true}));
});
test('document locks are isolated between MCP sessions and hosts',()=>{
  requestContext.run({sessionId:'one',host:'wps'},()=>TargetLockStore.lock('excel','One.xlsx'));
  requestContext.run({sessionId:'two',host:'wps'},()=>assert.equal(TargetLockStore.resolve('excel'),undefined));
  requestContext.run({sessionId:'one',host:'microsoft'},()=>assert.equal(TargetLockStore.resolve('excel'),undefined));
  requestContext.run({sessionId:'one',host:'wps'},()=>assert.equal(TargetLockStore.resolve('excel'),'One.xlsx'));
});
test('native runner handles Unicode and shell metacharacters without interpolation',async()=>{
  const input='中文 $HOME `whoami` "quotes"\nsecond';
  assert.equal(await runProcess(process.execPath,['-e','process.stdin.pipe(process.stdout)'],input),input);
  await assert.rejects(runProcess(process.execPath,['-e','setTimeout(()=>{},1000)'],'',30),/超时/);
});
