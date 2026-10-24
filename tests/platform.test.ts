import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
const root=fs.mkdtempSync(path.join(os.tmpdir(),'bridge-platform-'));
process.env.WPS_BRIDGE_HOME=path.join(root,'home');
process.env.WPS_BRIDGE_ADDON_DIR=path.join(root,'addons');
const {mergePluginIndex,AddonInstaller,OfficeAddonInstaller}=await import('../src/main/addon-installer.js');
const {mergeMcpConfig,InstallerEngine}=await import('../src/main/installer-engine.js');
const {getTools,validateArgs}=await import('../src/bridge/catalog.js');
const {requestContext}=await import('../src/bridge/context.js');
const {TargetLockStore}=await import('../src/bridge/gateway.js');
const {runProcess}=await import('../src/bridge/process-runner.js');

test('XML merge preserves unrelated plugins and rejects damaged input',()=>{
  const oldDisplay = ['Office', 'Agent', 'Bridge'].join(' ');
  const raw=`<jsplugins><jsplugin name="Other" type="et" url="./other"/><jsplugin name="WPS Bridge" type="et"/><jsplugin name="${oldDisplay} (表格)" type="et"/></jsplugins>`;
  const result=mergePluginIndex(raw);
  assert.match(result,/name="Other"/);
  assert.equal((result.match(/name="字浮 CharFloat/g)||[]).length,3);
  assert.equal(result.includes(oldDisplay),false);
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
test('MCP 配置改名：只写新名字 charfloat，并安全迁走旧版服务条目',()=>{
  const oldKey = ['office', 'agent', 'bridge'].join('-');
  const file=path.join(root,'rename.json');
  const entry={command:'/A.app/binary',args:['/A.app/cli.cjs']};
  // 场景一：已有我们的遗留条目（命令与参数一致）→ 应被迁走，只留新名字 charfloat
  fs.writeFileSync(file,JSON.stringify({mcpServers:{'wps-bridge':entry,[oldKey]:entry,other:{command:'x'}}}));
  mergeMcpConfig(file,entry);
  const a=JSON.parse(fs.readFileSync(file,'utf8'));
  assert.deepEqual(Object.keys(a.mcpServers).sort(),['charfloat','other']);
  assert.deepEqual(a.mcpServers['charfloat'],entry);
  // 场景二：别人的同名条目（命令不同）→ **必须原样保留**，不能被我们删掉
  fs.writeFileSync(file,JSON.stringify({mcpServers:{[oldKey]:{command:'someone-else'}}}));
  mergeMcpConfig(file,entry);
  const b=JSON.parse(fs.readFileSync(file,'utf8'));
  assert.deepEqual(b.mcpServers[oldKey],{command:'someone-else'});
  assert.deepEqual(b.mcpServers['charfloat'],entry);
  // 幂等：保留原有排版与注释，第二次不产生改动。
  const beforeRetry=fs.readFileSync(file,'utf8');
  mergeMcpConfig(file,entry);
  assert.equal(fs.readFileSync(file,'utf8'),beforeRetry);
});
test('status checks do not create plugin directories; install validates before writes',()=>{
  assert.equal(fs.existsSync(process.env.WPS_BRIDGE_ADDON_DIR!),false);AddonInstaller.checkStatus();assert.equal(fs.existsSync(process.env.WPS_BRIDGE_ADDON_DIR!),false);
  fs.mkdirSync(process.env.WPS_BRIDGE_ADDON_DIR!,{recursive:true});fs.writeFileSync(path.join(process.env.WPS_BRIDGE_ADDON_DIR!,'publish.xml'),'<broken>');
  assert.equal(AddonInstaller.install().success,false);assert.equal(fs.existsSync(path.join(process.env.WPS_BRIDGE_ADDON_DIR!,'wps-bridge')),false);
});
test('WPS 新安装不创建旧品牌目录，升级仍刷新现存缓存入口并迁移插件索引',()=>{
  const previousDir=process.env.WPS_BRIDGE_ADDON_DIR;
  const oldKey=['office','agent','bridge'].join('-');
  const oldDisplay=['Office','Agent','Bridge'].join(' ');
  try {
    for (const upgrading of [false,true]) {
      const dir=path.join(root,upgrading?'upgrade':'fresh');
      fs.mkdirSync(dir,{recursive:true});
      if (upgrading) {
        fs.mkdirSync(path.join(dir,oldKey));
        fs.writeFileSync(path.join(dir,oldKey,'manifest.xml'),`<manifest><name>${oldDisplay}</name></manifest>`);
        fs.writeFileSync(path.join(dir,'publish.xml'),`<jsplugins><jsplugin name="${oldDisplay} (表格)" type="et"/><jsplugin name="Other" type="et" url="./other"/></jsplugins>`);
      }
      process.env.WPS_BRIDGE_ADDON_DIR=dir;
      const result=AddonInstaller.install();
      assert.equal(result.success,true,result.message);
      assert.equal(fs.existsSync(path.join(dir,oldKey)),upgrading);
      const manifest=fs.readFileSync(path.join(dir,'wps-bridge/manifest.xml'),'utf8');
      assert.match(manifest,/<name>字浮 CharFloat<\/name>/);
      const ribbon=fs.readFileSync(path.join(dir,'wps-bridge/ribbon.xml'),'utf8');
      assert.match(ribbon,/<tab idMso="TabHome">/);
      assert.match(ribbon,/onAction="OnActionShowTaskPane"/);
      const panel=fs.readFileSync(path.join(dir,'wps-bridge/panel.html'),'utf8');
      assert.match(panel,/addon-core\.js\?v=[0-9a-f]{16}/);
      assert.match(panel,/panel\.css\?v=[0-9a-f]{16}/);
      const index=fs.readFileSync(path.join(dir,'publish.xml'),'utf8');
      assert.equal(index.includes(oldDisplay),false);
      assert.equal((index.match(/name="字浮 CharFloat/g)||[]).length,3);
      if (upgrading) {
        assert.match(index,/name="Other"/);
        assert.equal(fs.readFileSync(path.join(dir,oldKey,'manifest.xml'),'utf8'),manifest);
        assert.match(fs.readFileSync(path.join(dir,oldKey,'index.html'),'utf8'),/addon-core\.js\?v=[0-9a-f]{16}/);
      }
    }
  } finally { process.env.WPS_BRIDGE_ADDON_DIR=previousDir; }
});
test('Office 同版本清单的品牌和菜单更新仍会部署，清单 UUID 保持兼容', {skip:process.platform==='win32'}, ()=>{
  // 只部署至临时目录；Windows 分支包含真实注册表/证书写入，不在这里执行。
  const source=OfficeAddonInstaller.getSourceManifestPath();
  const current=fs.readFileSync(source,'utf8');
  const target=path.join(root,'wef','manifest.xml');
  fs.mkdirSync(path.dirname(target),{recursive:true});
  const oldDisplay=['Office','Agent','Bridge'].join(' ');
  for (const previous of [current.replaceAll('字浮 CharFloat',oldDisplay),current.replaceAll('https://github.com/Jolly-J/CharFloat','https://example.test/previous')]) {
    fs.writeFileSync(target,previous);
    (OfficeAddonInstaller as any).safeDeploy(source,target);
    const deployed=fs.readFileSync(target,'utf8');
    assert.equal(deployed,current);
    assert.match(deployed,/<Id>55555555-aaaa-bbbb-cccc-777777777777<\/Id>/);
    assert.match(deployed,/id="CharFloat\.Group"/);
    assert.match(deployed,/DisplayName DefaultValue="字浮 CharFloat \(Excel AI\)"/);
  }
});
test('新客户端配置只输出当前品牌及既有协议环境变量',()=>{
  const entry=InstallerEngine.configEntry();
  const oldEnv=['OFFICE','AGENT','BRIDGE','HOME'].join('_');
  assert.equal(Object.hasOwn(entry.env,oldEnv),false);
  assert.equal(entry.env.CHARFLOAT_HOME,process.env.WPS_BRIDGE_HOME);
  assert.equal(entry.env.WPS_BRIDGE_HOME,process.env.WPS_BRIDGE_HOME);
});
test('Office 状态检测要求清单与包内资源一致，同版本旧品牌仍提示升级',async t=>{
  const dir=path.join(root,'office-status');
  fs.mkdirSync(dir,{recursive:true});
  t.mock.method(OfficeAddonInstaller,'getWefDirectory',()=>dir);
  t.mock.method(globalThis,'fetch',async()=>{throw new Error('隔离测试：不访问真实后台');});
  const source=fs.readFileSync(OfficeAddonInstaller.getSourceManifestPath(),'utf8');
  const oldDisplay=['Office','Agent','Bridge'].join(' ');
  const {officeManifestMatchesResource}=await import('../src/bridge/build-fingerprint.js');
  for (const [content,current] of [
    [source,true],
    [source.replaceAll('字浮 CharFloat',oldDisplay),false],
    [source.replaceAll('https://github.com/Jolly-J/CharFloat','https://example.test/previous'),false]
  ] as const) {
    fs.writeFileSync(path.join(dir,'wps-bridge-manifest.xml'),content);
    const status=await OfficeAddonInstaller.checkStatus();
    assert.equal(status.installed,true);
    assert.equal(status.current,current);
    assert.equal(status.needsUpgrade,!current);
    // HTTP 后台也调用同一判据，避免客户端回退与后台报告冲突。
    assert.equal(officeManifestMatchesResource(content),current);
  }
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

// ===== 多客户端接入：格式适配器与注册表 =====
const { mergeCodexToml, isCodexTomlConfigured, mergeOpencodeConfig, isOpencodeConfigured } = await import('../src/main/mcp-adapters.js');
const { CLIENT_DEFINITIONS, clientPathContext } = await import('../src/main/client-registry.js');

test('Codex config.toml 合并：保留其他段落、幂等，并可被状态检测识别', () => {
  const file = path.join(root, 'codex-config.toml');
  fs.writeFileSync(file, ['# 用户自己的配置', 'model = "gpt-5"', '', '[mcp_servers]', '', '[mcp_servers.playwright]', 'command = "npx"', 'args = ["x"]', ''].join('\n'));
  const entry = { command: '/A.app/bin', args: ['/A.app/cli.cjs'], env: { ELECTRON_RUN_AS_NODE: '1', CHARFLOAT_HOME: '/tmp/home' } };
  mergeCodexToml(file, entry);
  const first = fs.readFileSync(file, 'utf8');
  assert.match(first, /\[mcp_servers\.charfloat\]/);
  assert.match(first, /\[mcp_servers\.charfloat\.env\]/);
  assert.match(first, /model = "gpt-5"/);
  assert.match(first, /\[mcp_servers\.playwright\]/);
  assert.equal(isCodexTomlConfigured(file), true);
  mergeCodexToml(file, { ...entry, env: { ELECTRON_RUN_AS_NODE: '1' } });
  const second = fs.readFileSync(file, 'utf8');
  // 幂等：重复写入不会堆出第二份 charfloat 段
  assert.equal((second.match(/\[mcp_servers\.charfloat\]/g) || []).length, 1);
  assert.equal((second.match(/\[mcp_servers\.charfloat\.env\]/g) || []).length, 1);
  assert.match(second, /\[mcp_servers\.playwright\]/);
  assert.match(second, /# 用户自己的配置/);
});
test('opencode 配置合并：保留 schema 与其他服务，按 local 结构写入', () => {
  const file = path.join(root, 'opencode.json');
  fs.writeFileSync(file, JSON.stringify({ $schema: 'https://opencode.ai/config.json', mcp: { other: { type: 'remote', url: 'https://example.test' } } }));
  const entry = { command: '/A/bin', args: ['/A/cli.cjs'], env: { ELECTRON_RUN_AS_NODE: '1' } };
  mergeOpencodeConfig(file, entry);
  const config = JSON.parse(fs.readFileSync(file, 'utf8'));
  assert.equal(config.$schema, 'https://opencode.ai/config.json');
  assert.equal(config.mcp.other.type, 'remote');
  assert.deepEqual(config.mcp.charfloat, { type: 'local', command: ['/A/bin', '/A/cli.cjs'], environment: { ELECTRON_RUN_AS_NODE: '1' }, enabled: true });
  assert.equal(isOpencodeConfigured(file), true);
  fs.writeFileSync(file, '{broken');
  assert.throws(() => mergeOpencodeConfig(file, entry));
  assert.equal(fs.readFileSync(file, 'utf8'), '{broken');
});
test('客户端注册表：分类、格式与配置路径完整', () => {
  const ctx = clientPathContext({ home: '/tmp/regression-home', appData: '/tmp/regression-appdata', platform: 'darwin' });
  const ids = CLIENT_DEFINITIONS.map(def => def.id);
  assert.equal(new Set(ids).size, ids.length);
  for (const def of CLIENT_DEFINITIONS) {
    assert.ok(def.region === 'cn' || def.region === 'global', `${def.id} 缺少分类`);
    assert.ok(['json-mcp', 'json-vscode', 'toml-codex', 'json-opencode', 'yaml-dsh', 'manual'].includes(def.configKind), `${def.id} 配置格式非法`);
    const paths = def.paths(ctx);
    assert.ok(paths.detectPaths.length > 0, `${def.id} 缺少检测路径`);
    if (def.configKind === 'manual') continue;
    assert.ok(paths.configPath && path.isAbsolute(paths.configPath), `${def.id} 缺少配置路径`);
  }
  assert.deepEqual([...new Set(CLIENT_DEFINITIONS.map(def => def.region))].sort(), ['cn', 'global']);
  assert.ok(CLIENT_DEFINITIONS.filter(def => def.region === 'cn').length >= 4);
  assert.ok(CLIENT_DEFINITIONS.filter(def => def.region === 'global').length >= 6);
});
test('空配置文件与 BOM 视为未配置，可正常写入（Antigravity 生成的 0 字节文件）', () => {
  const file = path.join(root, 'antigravity-empty.json');
  fs.writeFileSync(file, '');
  const entry = { command: '/A/bin', args: ['/A/cli.cjs'], env: {} };
  mergeMcpConfig(file, entry);
  assert.deepEqual(JSON.parse(fs.readFileSync(file, 'utf8')).mcpServers.charfloat, entry);
  // 纯空白同样按空配置处理，不再是"JSON 损坏"
  fs.writeFileSync(file, '  \n\t');
  mergeMcpConfig(file, entry);
  assert.ok(JSON.parse(fs.readFileSync(file, 'utf8')).mcpServers.charfloat);
  // 带 BOM 的合法 JSON 也能解析
  fs.writeFileSync(file, '\uFEFF{"mcpServers":{"other":{"command":"x"}}}');
  mergeMcpConfig(file, entry);
  const config = JSON.parse(fs.readFileSync(file, 'utf8'));
  assert.ok(config.mcpServers.other);
  assert.ok(config.mcpServers.charfloat);
});
