// 探针：以 stdio 启动 dist/bridge/cli.cjs，做一次 MCP 握手并**只输出工具数量**。
// 供 check-cli-protection.mjs 复用（它需要"能跑起来"的客观证据）。
import { spawn } from 'node:child_process';
import path from 'node:path';

const cli = path.join(process.cwd(), 'dist/bridge/cli.cjs');
const p = spawn(process.execPath, [cli, '--stdio'], {
  env: { ...process.env, WPS_BRIDGE_HOME: process.env.WPS_BRIDGE_HOME || path.join(process.env.HOME, '.wps-bridge') },
  stdio: ['pipe', 'pipe', 'pipe'],
});
let buf = ''; let count = 0;
p.stdout.on('data', (d) => {
  buf += d.toString();
  let i;
  while ((i = buf.indexOf('\n')) >= 0) {
    const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1);
    if (!line) continue;
    try {
      const m = JSON.parse(line);
      if (m.id === 2) count = m.result?.tools?.length ?? 0;
    } catch {}
  }
});
const send = (o) => p.stdin.write(JSON.stringify(o) + '\n');
send({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'probe', version: '1' } } });
setTimeout(() => send({ jsonrpc: '2.0', method: 'notifications/initialized' }), 300);
setTimeout(() => send({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} }), 700);
setTimeout(() => { console.log(String(count)); p.kill(); process.exit(0); }, 4000);
