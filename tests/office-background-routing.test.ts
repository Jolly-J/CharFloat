import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { WpsBridgeServer } from '../src/bridge/ws-server.js';
import { BridgeError, routeOfficeFailure } from '../src/bridge/errors.js';

class WorkbookSocket extends EventEmitter {
  readyState = 1;
  closeCount = 0;
  requests: any[] = [];
  value: unknown = null;
  constructor(public workbookName: string, public documentUrl: string) { super(); }
  packet(data: any) { this.emit('message', Buffer.from(JSON.stringify(data))); }
  register() { this.packet({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.2.0',
    summary: { workbookName: this.workbookName, documentUrl: this.documentUrl, backgroundStartup: 'enabled' } }); }
  send(raw: string, callback?: (error?: Error) => void) {
    const data = JSON.parse(raw); this.requests.push(data);
    if (data.method === 'write_range') this.value = data.params.values[0][0];
    queueMicrotask(() => this.packet({ type: 'rpc_response', id: data.id,
      result: { workbookName: this.workbookName, documentUrl: this.documentUrl, value: this.value } }));
    callback?.();
  }
  close(code = 1000, reason = '') { this.closeCount++; this.readyState = 3; this.emit('close', code, Buffer.from(reason)); }
}
function add(server: WpsBridgeServer, name: string, url: string) {
  const socket = new WorkbookSocket(name, url);
  (server as any).acceptSocket(socket); socket.register(); return socket;
}

test('两个 Office 后台工作簿互不替换，写入按 workbookName 进入对应文档', async () => {
  const server = new WpsBridgeServer(23101, 23102);
  const a = add(server, 'A.xlsx', 'file:///tmp/A.xlsx'), b = add(server, 'B.xlsx', 'file:///tmp/B.xlsx');
  assert.equal(a.closeCount, 0);
  assert.equal(b.closeCount, 0);
  assert.equal(server.getState().officeWorkbooks?.length, 2);
  await server.callOfficeAddon('write_range', { workbookName: 'A.xlsx', values: [[42]] });
  await server.callOfficeAddon('write_range', { workbookName: 'B.xlsx', values: [[99]] });
  assert.equal(a.value, 42); assert.equal(b.value, 99);
  b.close();
  assert.equal(server.getState().components.msExcel?.connected, true);
  assert.equal(server.getState().officeWorkbooks?.length, 1);
  const result: any = await server.callOfficeAddon('read_range');
  assert.equal(result.workbookName, 'A.xlsx');
  a.close();
  assert.equal(server.getState().isMsOfficeConnected, false);
});

test('后台工作簿有歧义时先拒绝，不发写请求也不允许 COM 猜测目标', async () => {
  const server = new WpsBridgeServer(23103, 23104);
  const a = add(server, 'Same.xlsx', 'file:///C:/one/Same.xlsx'), b = add(server, 'Same.xlsx', 'file:///C:/two/Same.xlsx');
  for (const params of [{ values: [[1]] }, { workbookName: 'Same.xlsx', values: [[1]] }]) {
    await assert.rejects(server.callOfficeAddon('write_range', params), error => {
      assert.ok(error instanceof BridgeError);
      assert.equal(error.executed, 'no'); assert.equal(error.unsafe, true);
      assert.equal(routeOfficeFailure('write_range', error, { platform: 'win32', supportedMethods: new Set(['write_range']) }).action, 'reject');
      return true;
    });
  }
  assert.equal(a.requests.length, 0); assert.equal(b.requests.length, 0);
  const result: any = await server.callOfficeAddon('write_range', { workbookName: 'C:\\one\\Same.xlsx', values: [[7]] });
  assert.equal(result.documentUrl, 'file:///C:/one/Same.xlsx');
  assert.equal(a.value, 7); assert.equal(b.value, null);
  a.close(); b.close();
});

test('Office 文档改名事件更新路由，WPS 请求绝不回退到 Office 连接', async () => {
  const server = new WpsBridgeServer(23105, 23106);
  const a = add(server, 'Old.xlsx', 'file:///tmp/Old.xlsx');
  a.workbookName = 'New.xlsx'; a.documentUrl = 'file:///tmp/New.xlsx';
  a.packet({ type: 'event', event: 'selection_change', host: 'microsoft', data: { workbookName: a.workbookName, documentUrl: a.documentUrl, sheetName: 'Sheet1', address: 'A1' } });
  const result: any = await server.callOfficeAddon('read_range', { workbookName: 'New.xlsx' });
  assert.equal(result.workbookName, 'New.xlsx');
  const before = a.requests.length;
  await assert.rejects(server.callWps('get_workspace_summary'), /WPS .*加载项未连接/);
  assert.equal(a.requests.length, before);
  a.close();
});
