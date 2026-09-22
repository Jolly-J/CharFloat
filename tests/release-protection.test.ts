import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import { execFileSync } from 'node:child_process';
import { ADDONS, protectAddon, checkFiles, POLICY } from '../scripts/lib/release-protection.mjs';

test('发布 WPS 副本通过相同的宿主 RPC 行为回归，开发文件不变', () => {
  const source = fs.readFileSync(ADDONS[0], 'utf8');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'oab-protected-test-'));
  try {
    const code = protectAddon(source, ADDONS[0]);
    assert.equal(code, protectAddon(source, ADDONS[0]), '同输入须可重复构建');
    const file = path.join(tmp, 'addon.js');
    fs.writeFileSync(file, code);
    execFileSync(process.execPath, ['--import', 'tsx', '--test', 'tests/addon.test.ts'], {
      env: { ...process.env, OAB_TEST_WPS_ARTIFACT: file }, timeout: 30000, stdio: 'pipe'
    });
    assert.equal(fs.readFileSync(ADDONS[0], 'utf8'), source);
    const fingerprint = code.match(/ADDON_BUILD_FINGERPRINT:\s*(\w+)/)?.[1];
    let socket: any;
    const sent: any[] = [];
    class Socket { static OPEN = 1; readyState = 1; constructor() { socket = this; } send(s: string) { sent.push(JSON.parse(s)); } }
    vm.runInNewContext(code, { window: { WPS_BRIDGE_CONFIG: { port: 12345, token: 'test' }, addEventListener() {} },
      document: { getElementById: () => null, readyState: 'complete' }, WebSocket: Socket,
      console: { log() {} }, setTimeout() {}, setInterval() {}, clearTimeout() {}, alert() {} });
    socket.onopen();
    assert.ok(sent.some(p => JSON.stringify(p).includes(fingerprint)), '运行时上报必须与保护后指纹相同');
  } finally { fs.rmSync(tmp, { recursive: true, force: true }); }
});

async function officeRpc(code: string) {
  let socket: any, ready: any;
  const sent: any[] = [];
  const range: any = { address: 'Sheet1!A1', rowCount: 1, columnCount: 1, values: [[0]], formulas: [['=0']], text: [['0']], load() {} };
  range.getCell = () => range;
  const sheet = { name: 'Sheet1', load() {}, getRange: () => range, onSelectionChanged: { add() {} } };
  const context = { workbook: { worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }, getSelectedRange: () => range }, sync: async () => {} };
  class Socket { static OPEN = 1; static CONNECTING = 0; readyState = 1; constructor() { socket = this; } send(s: string) { sent.push(JSON.parse(s)); } }
  const window = { location: { hash: '', href: 'https://localhost:19891/office-addon/taskpane.html' } };
  vm.runInNewContext(code, { window, document: { readyState: 'complete', getElementById: () => null },
    Office: { onReady: (fn: any) => { ready = fn; }, HostType: { Excel: 'Excel' }, context: { document: { url: 'file:///Exact.xlsx' } } },
    Excel: { run: async (fn: any) => fn(context) }, WebSocket: Socket, console: { log() {}, error() {} },
    setTimeout() {}, clearTimeout() {}, URL });
  ready({ host: 'Excel' });
  await socket.onopen();
  for (const [id, method, params] of [[1, 'read_range', { address: 'A1' }], [2, 'write_range', { address: 'A1', values: [['中文']], formulas: [['=1+1']] }], [3, 'not-a-tool', {}]] as const) {
    socket.onmessage({ data: JSON.stringify({ id, method, params }) });
    await new Promise(resolve => setImmediate(resolve));
  }
  return JSON.parse(JSON.stringify({ sent, values: range.values, formulas: range.formulas }));
}

test('发布 Office 副本保留注册、零值读取、写入与未知方法错误', async () => {
  const source = fs.readFileSync(ADDONS[1], 'utf8');
  const expected = await officeRpc(source);
  const actual = await officeRpc(protectAddon(source, ADDONS[1]));
  assert.deepEqual(actual, expected);
  assert.equal(actual.sent[0].type, 'register');
  assert.equal(actual.sent.find((p: any) => p.id === 1).result.values[0][0], 0);
  assert.equal(actual.values[0][0], '中文');
  assert.ok(actual.sent.find((p: any) => p.id === 3).error);
});

test('发布门禁拒绝源码、备份、旧入口和缺失的 Agent 使用资源', () => {
  const names = [...POLICY.files, 'package.json', 'protection-manifest.json'];
  for (const leak of ['dist/main/addon-installer.js', 'dist/preload/index.js', 'dist/bridge/cli-full.cjs',
    'wps-addon/src/excel.js', 'skills/AGENTS.md', 'wps-addon/addon-core.js.map', 'wps-addon/addon-core.js.backup']) {
    assert.ok(checkFiles([...names, leak], () => Buffer.from('')).some((e: string) => e.includes(leak)));
  }
  assert.ok(checkFiles(names.filter(n => n !== 'skills/office-agent-bridge/SKILL.md'), () => Buffer.from(''))
    .some((e: string) => e.includes('缺少运行资源')));
});
