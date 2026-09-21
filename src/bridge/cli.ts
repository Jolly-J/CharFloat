#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createMcpServer } from './mcp-server.js';
import { bridgeServer } from './ws-server.js';
import { ensureService, serviceRequest, probeService } from './service-client.js';
import { atomicWrite, getToken, runtimeHome, runtimePort, VERSION, resourcePath } from './runtime.js';

// stdout is exclusively the MCP transport. Never log arguments or credentials.
console.log = (...args) => console.error(...args);
async function main() {
  const mode = process.argv[2];
  if (mode === '--serve') {
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
  if (mode === '--status' || mode === '--doctor') { process.stdout.write(JSON.stringify(await probeService()) + '\n'); return; }
  if (mode === '--stop') { await serviceRequest('/api/v1/service/stop', {}); return; }
  if (mode === '--help') { process.stdout.write('wps-bridge-mcp [--start | --status | --doctor | --stop]\n无参数：stdio MCP，按需启动独立后台。\n'); return; }
  const entry = path.resolve(process.argv[1]);
  await ensureService(entry);
  if (mode === '--start') atomicWrite(path.join(runtimeHome(), 'installation.json'), JSON.stringify({ executable: process.execPath, cli: entry, resources: path.dirname(resourcePath('package.json')), version: VERSION }));
  if (mode === '--start') { process.stdout.write('Bridge 后台已就绪\n'); return; }
  const server = createMcpServer({
    tools: () => serviceRequest('/api/v1/mcp-tools'),
    call: async (name, args, sessionId) => (await serviceRequest('/api/v1/tool/call', { name, arguments: args, sessionId, clientName: 'stdio MCP' })).data,
    capabilities: () => serviceRequest('/api/v1/capabilities'),
    prompt: async () => {
      const r = await fetch(`http://127.0.0.1:${runtimePort()}/api/v1/prompts/aesthetic`, { headers: { Authorization: `Bearer ${getToken()}` }, signal: AbortSignal.timeout(5000) });
      if (!r.ok) throw new Error('提示词资源缺失'); return r.text();
    }
  });
  await server.connect(new StdioServerTransport());
  const cleanup = async () => { await server.close(); process.exit(0); };
  process.on('SIGINT', cleanup); process.on('SIGTERM', cleanup); process.stdin.on('end', cleanup);
}
main().catch(error => { console.error(error.message); process.exit(1); });
