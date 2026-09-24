#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createMcpServer } from './mcp-server.js';
import type { ToolService } from './contracts/tool-service.js';
import { bridgeServer } from './ws-server.js';
import { composeBridgeServer } from './compose.js';
import { ensureService, serviceRequest, probeService } from './service-client.js';
import { atomicWrite, getToken, runtimeHome, runtimePort, VERSION, resourcePath } from './runtime.js';

// stdout is exclusively the MCP transport. Never log arguments or credentials.
console.log = (...args) => console.error(...args);
async function main() {
  const mode = process.argv[2];
  if (mode === '--serve') {
    composeBridgeServer();   // 组装入口：为 HTTP/MCP 路由注入工具服务（P2.2）
    await bridgeServer.start();
    atomicWrite(path.join(runtimeHome(), 'service.json'), JSON.stringify({ pid: process.pid, port: runtimePort(), version: VERSION, startedAt: Date.now() }));
    const cleanup = () => { bridgeServer.stop(); process.exit(0); };
    process.on('SIGTERM', cleanup); process.on('SIGINT', cleanup); return;
  }
  if (mode === '--repair-addon') {
    const { AddonInstaller } = await import('../main/addon-installer.js');
    const result = AddonInstaller.install();
    process.stdout.write(JSON.stringify(result) + '\n');
    if (!result.success) process.exitCode = 1;
    return;
  }
  if (mode === '--status') { process.stdout.write(JSON.stringify(await probeService()) + '\n'); return; }
  if (mode === '--doctor') {
    // ISS-50：原来 --doctor 与 --status 输出同一个对象，正常字段（凭据/安装记录）
    // 和原始网络错误混在一起，调用方无法判断"到底哪一步坏了"。这里分开：
    // ready/occupied 是结论，problem/recovery 是可执行动作，rawError 只作原始证据。
    const probe: any = await probeService();
    const home = runtimeHome();
    const report = {
      service: 'wps-bridge',
      version: VERSION,
      address: `http://127.0.0.1:${runtimePort()}`,
      home,
      ready: probe.ready === true,
      occupied: probe.occupied === true,
      credentialsPresent: fs.existsSync(path.join(home, 'token')),
      installationRecord: fs.existsSync(path.join(home, 'installation.json')),
      ...(probe.ready
        ? { connected: true, info: probe.info }
        : { connected: false, problem: probe.message, recovery: probe.recovery, rawError: probe.error })
    };
    process.stdout.write(JSON.stringify(report, null, 2) + '\n');
    if (!report.ready) process.exitCode = 1;
    return;
  }
  if (mode === '--stop') { await serviceRequest('/api/v1/service/stop', {}); return; }
  if (mode === '--help') { process.stdout.write('office-agent-bridge [--start | --status | --doctor | --stop]\n无参数：stdio MCP，按需启动独立后台。\n'); return; }
  const entry = path.resolve(process.argv[1]);
  await ensureService(entry);
  if (mode === '--start') atomicWrite(path.join(runtimeHome(), 'installation.json'), JSON.stringify({ executable: process.execPath, cli: entry, resources: path.dirname(resourcePath('package.json')), version: VERSION }));
  if (mode === '--start') { process.stdout.write('Bridge 后台已就绪\n'); return; }
  // stdio 客户端不直接持有工具实现：经后台服务转发，同样以注入的 ToolService 形状交给协议层。
  // sessionId 必须显式转发：后台按它隔离目标锁等会话状态，AsyncLocalStorage 不会跨 HTTP 传递。
  const remote: ToolService = {
    list: () => serviceRequest('/api/v1/mcp-tools'),
    capabilities: () => serviceRequest('/api/v1/capabilities'),
    execute: async (name, args, context) => (await serviceRequest('/api/v1/tool/call', {
      name, arguments: args, sessionId: context.sessionId, clientName: context.clientName
    })).data
  };
  const readPrompt = async () => {
    const r = await fetch(`http://127.0.0.1:${runtimePort()}/api/v1/prompts/aesthetic`, { headers: { Authorization: `Bearer ${getToken()}` }, signal: AbortSignal.timeout(5000) });
    if (!r.ok) throw new Error('提示词资源缺失'); return r.text();
  };
  const server = createMcpServer(remote, { prompt: readPrompt, clientName: 'stdio MCP' });
  await server.connect(new StdioServerTransport());
  const cleanup = async () => { await server.close(); process.exit(0); };
  process.on('SIGINT', cleanup); process.on('SIGTERM', cleanup); process.stdin.on('end', cleanup);
}
main().catch(error => { console.error(error.message); process.exit(1); });
