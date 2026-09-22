/**
 * stdio 会话隔离回归测试（P2.2 补正）。
 *
 * 覆盖的真实回归：`cli.ts` 的远程 ToolService 代理转发 `/api/v1/tool/call` 时漏传 `sessionId`，
 * 后台把**所有** stdio 客户端都当成默认会话 `http-local`。后果是不同客户端的文档锁互相覆盖：
 * 客户端 A 锁定 A.xlsx 后，若 B 锁定 B.xlsx，A 再查询锁状态会拿到 B.xlsx——
 * A 后续依赖锁定目标的操作会误指向 B 的文件。
 *
 * 本测试走**真实链路**：两个独立 stdio 客户端进程 → 同一个后台服务 → 会话隔离。
 * 不使用进程内直接调用，因为 `AsyncLocalStorage` 不会自动跨 HTTP 传递，只有端到端才能发现该问题。
 *
 * 前置条件：需要已构建的 `dist/bridge/cli.cjs`（同 lifecycle.test.ts），运行前执行 `npm run build`。
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { WebSocket } from 'ws';

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'bridge-session-'));
const port = String(45000 + Math.floor(Math.random() * 8000));
const env = { ...process.env, WPS_BRIDGE_HOME: dir, WPS_BRIDGE_PORT: port, WPS_BRIDGE_RESOURCES: process.cwd() } as Record<string, string>;
const base = `http://127.0.0.1:${port}`;

let a: Client, b: Client;

async function connect(name: string): Promise<Client> {
  const client = new Client({ name, version: '1' });
  await client.connect(new StdioClientTransport({ command: process.execPath, args: [path.resolve('dist/bridge/cli.cjs')], env, stderr: 'pipe' }));
  return client;
}

async function call(client: Client, name: string, args: Record<string, unknown> = {}): Promise<any> {
  const res: any = await client.callTool({ name, arguments: args });
  const text = (res.content || []).find((c: any) => c.type === 'text')?.text;
  if (res.isError) throw new Error(text);
  return text ? JSON.parse(text) : null;
}

test.before(async () => {
  a = await connect('session-client-a');
  b = await connect('session-client-b');
});

test.after(async () => {
  await a?.close().catch(() => {});
  await b?.close().catch(() => {});
  try {
    if (fs.existsSync(path.join(dir, 'token'))) {
      await fetch(base + '/api/v1/service/stop', {
        method: 'POST',
        headers: { Authorization: `Bearer ${fs.readFileSync(path.join(dir, 'token'), 'utf8')}` }
      });
    }
  } catch {}
});

test('两个 stdio 客户端的 WPS 文档锁互不串扰', async () => {
  await call(a, 'wps_lock_target_document', { component: 'excel', targetName: 'A.xlsx' });
  await call(b, 'wps_lock_target_document', { component: 'excel', targetName: 'B.xlsx' });

  const statusA = await call(a, 'wps_get_locked_status');
  const statusB = await call(b, 'wps_get_locked_status');
  assert.equal(statusA.gatewayLocks.excel, 'A.xlsx', 'A 应看到自己锁定的 A.xlsx');
  assert.equal(statusB.gatewayLocks.excel, 'B.xlsx', 'B 应看到自己锁定的 B.xlsx，而不是被 A 覆盖或覆盖 A');

  // 单方解锁不得影响另一方
  await call(a, 'wps_unlock_target_document', { component: 'excel' });
  const afterA = await call(a, 'wps_get_locked_status');
  const afterB = await call(b, 'wps_get_locked_status');
  assert.equal(afterA.gatewayLocks.excel, undefined, 'A 解锁后自己不应再有 excel 锁');
  assert.equal(afterB.gatewayLocks.excel, 'B.xlsx', 'A 解锁不得影响 B 的锁');
});

test('两个 stdio 客户端的 Microsoft 目标锁互不串扰', async () => {
  const lockA = await call(a, 'office_lock_target', { component: 'excel', targetName: 'A2.xlsx' });
  const lockB = await call(b, 'office_lock_target', { component: 'ppt', targetName: 'B2.pptx' });

  assert.equal(lockA.lockedTarget, 'A2.xlsx');
  assert.deepEqual(Object.keys(lockA.allLocks).sort(), ['excel'], 'A 只应看到自己的 excel 锁');
  assert.deepEqual(Object.keys(lockB.allLocks).sort(), ['ppt'], 'B 只应看到自己的 ppt 锁');

  // 单方解锁不得影响另一方
  const unlockedA = await call(a, 'office_unlock_target', {});
  assert.deepEqual(Object.keys(unlockedA.allLocks).sort(), [], 'A 解锁后自己不应再有锁');
  const afterB = await call(b, 'office_lock_target', { component: 'ppt', targetName: 'B2.pptx' });
  assert.deepEqual(Object.keys(afterB.allLocks).sort(), ['ppt'], 'A 解锁不得影响 B 的 ppt 锁');
});

test('真实 stdio 链路的审计归属：clientName 经 CLI 与 HTTP 转发后仍为 stdio MCP', async () => {
  // 本项刻意**不**直接调用 executeCatalogTool：必须经过 stdio 客户端 → CLI 代理 → HTTP → 后台 → 模拟宿主，
  // 才能覆盖 P2.2 真实出错过的那一段（当时 clientName 被改成 MCP Agent，现有测试发现不了）。
  const token = fs.readFileSync(path.join(dir, 'token'), 'utf8');
  const socket = new WebSocket(`ws://127.0.0.1:${port}/addon?token=${token}`);
  let hostCalls = 0;
  try {
    await new Promise<void>((resolve, reject) => {
      socket.once('open', () => resolve());
      socket.once('error', reject);
    });
    socket.on('message', (raw: any) => {
      const packet = JSON.parse(raw.toString());
      if (packet.type !== 'rpc_request' && !packet.method) return;
      hostCalls += 1;
      socket.send(JSON.stringify({
        type: 'rpc_response',
        id: packet.id,
        result: {
          success: true, sheetName: 'Sheet1', address: 'A1:B1',
          rowCount: 1, columnCount: 2, modifiedCount: 2, diff: [], workbookName: 'Audit.xlsx'
        }
      }));
    });
    socket.send(JSON.stringify({
      type: 'register', client: 'wps-et-addon', version: '2.0.0',
      summary: { workbookName: 'Audit.xlsx', activeSheetName: 'Sheet1' }
    }));
    await new Promise(r => setTimeout(r, 60));

    const written = await call(a, 'excel_patch_cells', {
      host: 'wps', address: 'A1:B1', values: [[1, 2]], workbookName: 'Audit.xlsx'
    });
    assert.equal(hostCalls, 1, '模拟宿主应被调用一次');
    assert.ok(written.auditId, '写入应产生 auditId');

    const record = await call(a, 'wps_get_audit_record', { auditId: written.auditId });
    assert.equal(record.clientName, 'stdio MCP', 'stdio 路径的 clientName 必须是 stdio MCP，不能退回默认值');
    assert.notEqual(record.clientName, 'MCP Agent');
    assert.equal(record.host, 'wps');
    assert.equal(record.workbookName, 'Audit.xlsx');
  } finally {
    socket.terminate();
  }
});
