// 只验证独立测试服务：临时 HOME、随机空闲端口、前台子进程、无真实宿主。
// 用法：node scripts/probe-release-mcp.mjs <electron> <cli.cjs> <resources-root> <tools-snapshot.json>
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import net from 'node:net';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

const [exeArg, cliArg, resourcesArg, snapshotArg] = process.argv.slice(2);
if (!snapshotArg) throw new Error('必须提供 electron、CLI、资源根目录与契约快照四个参数');
const [executable, cli, resources, snapshot] = [exeArg, cliArg, resourcesArg, snapshotArg].map(p => path.resolve(p));
assert.ok(!fs.existsSync(path.join(path.dirname(cli), 'cli-full.cjs')), '探针只接受没有明文回退的发布入口');
const home = fs.mkdtempSync(path.join(os.tmpdir(), 'oab-mcp-probe-'));
const ports = [];
const reservations = [];
for (let i = 0; i < 2; i++) {
  const s = net.createServer();
  s.listen(0, '127.0.0.1');
  await once(s, 'listening');
  ports.push(s.address().port); reservations.push(s);
}
await Promise.all(reservations.map(s => new Promise(resolve => s.close(resolve))));
const env = { ...process.env, ELECTRON_RUN_AS_NODE: '1', WPS_BRIDGE_HOME: home,
  WPS_BRIDGE_PORT: String(ports[0]), WPS_BRIDGE_HTTPS_PORT: String(ports[1]), WPS_BRIDGE_RESOURCES: resources };
const service = spawn(executable, [cli, '--serve'], { env, cwd: home, stdio: ['ignore', 'ignore', 'pipe'] });
let log = '', spawnError;
service.stderr.on('data', d => { log = (log + d).slice(-6000); });
service.on('error', e => { spawnError = e; });
const client = new Client({ name: 'release-protection-probe', version: '1' });
const timer = setTimeout(() => { service.kill(); process.exitCode = 1; }, 30000);
try {
  const start = Date.now();
  while (!log.includes(`listening on 127.0.0.1:${ports[0]}`)) {
    if (spawnError) throw spawnError;
    if (service.exitCode !== null || Date.now() - start > 15000) throw new Error(`测试服务未就绪: ${log}`);
    await new Promise(r => setTimeout(r, 50));
  }
  const transport = new StdioClientTransport({ command: executable, args: [cli], env, cwd: home, stderr: 'pipe' });
  await client.connect(transport);
  const { tools } = await client.listTools();
  const expected = JSON.parse(fs.readFileSync(snapshot)).tools;
  const normalize = ts => ts.map(t => ({ name: t.name, description: t.description, inputSchema: t.inputSchema,
    readOnlyHint: t.readOnlyHint ?? (t.annotations?.readOnlyHint === true) }));
  assert.deepEqual(normalize(tools), normalize(expected), '发布包 tools/list 名称、顺序、描述、参数和只读标记必须与快照一致');
  const resource = await client.readResource({ uri: 'bridge://capabilities' });
  assert.ok(JSON.parse(resource.contents[0].text));
  const prompt = await client.getPrompt({ name: 'excel_aesthetic_system' });
  assert.equal(prompt.messages[0].content.text, fs.readFileSync(path.join(resources, 'prompts/excel_aesthetic_system.md'), 'utf8'));
  const result = await client.callTool({ name: 'bridge_get_capabilities', arguments: {} });
  assert.ok(!result.isError, '发布包工具调用失败');
  const bad = await client.callTool({ name: '__missing_tool__', arguments: {} });
  assert.equal(bad.isError, true, '未知工具须返回可理解错误');
  console.log(JSON.stringify({ toolCount: tools.length, contractMatches: true, capabilitiesResource: true,
    promptMatches: true, toolCall: true, unknownToolError: true, bytecodeOnly: true,
    realOfficeHostVerified: false }, null, 2));
} finally {
  clearTimeout(timer);
  await client.close().catch(() => {});
  if (service.exitCode === null && service.signalCode === null) {
    service.kill();
    await Promise.race([once(service, 'exit'), new Promise(resolve => setTimeout(resolve, 3000))]);
    if (service.exitCode === null && service.signalCode === null) service.kill('SIGKILL');
  }
  fs.rmSync(home, { recursive: true, force: true });
}
