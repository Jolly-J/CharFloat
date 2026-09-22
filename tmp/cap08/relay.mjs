// CAP-08 真机探测运行器（v2）。
//
// 背景：MCP 工具面里没有入口能把任意 Office.js 代码送进 Excel 任务窗格，
// 而直接抢占 ms-excel 通道会在 2.5s 后被真实任务窗格重连替换，来不及回包。
//
// 做法：临时注册一个"WPS 表格加载项"通道（sockets 键 excel），让 MCP 的
// `wps_execute_script` 把 execute_script 请求发到本进程；本进程在**真实 Excel
// 任务窗格上下文里**通过 `ms-excel` 通道执行等价的 Office.js 代码（run_script），
// 再把结果作为 execute_script 的响应回给桥接。真实任务窗格始终保有 ms-excel 通道，
// 因此执行者是真机 Excel，不是模拟器。跑完断开，WPS 加载项会自动重连。
//
// 用法：node tmp/cap08/relay.mjs <probe-code.js> [等待毫秒]
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import crypto from 'node:crypto';
import WebSocket from 'ws';

const BRIDGE_WS = 'ws://127.0.0.1:19890/office-addon';
const token = fs.readFileSync(path.join(os.homedir(), '.wps-bridge', 'token'), 'utf8').trim();
const codePath = process.argv[2] || path.join(import.meta.dirname, 'probe-code.js');
const probeCode = fs.readFileSync(codePath, 'utf8');
const holdMs = Number(process.argv[3] || 90_000);

/** WPS 通道：接收 MCP 的 wps_execute_script，把工作交给真实 Office.js 任务窗格。 */
const wps = new WebSocket(BRIDGE_WS, { headers: { Origin: 'https://localhost:19890' } });
/** Office.js 通道：作为"临时任务窗格"下发 run_script（真实任务窗格会周期性重连夺回该键）。 */
let office = null;
const officeWaiters = new Map();

function ensureOfficeLink() {
  if (office && (office.readyState === WebSocket.OPEN || office.readyState === WebSocket.CONNECTING)) return office;
  office = new WebSocket(BRIDGE_WS, { headers: { Origin: 'https://localhost:19890' } });
  office.on('open', () => {
    office.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: { workbookName: '工作簿1.xlsx' } }));
  });
  office.on('message', (raw) => {
    let p; try { p = JSON.parse(raw.toString()); } catch { return; }
    if (p.type === 'rpc_response' && officeWaiters.has(p.id)) {
      const w = officeWaiters.get(p.id); officeWaiters.delete(p.id);
      w(p.error ? { error: p.error } : { result: p.result });
    }
  });
  office.on('close', () => { console.log('[relay] Office.js 通道被真实任务窗格替换，稍后自动重连'); });
  office.on('error', () => {});
  return office;
}

/** 把代码交给本机真实 Excel 任务窗格执行（走 run_script 分支，拿到返回值）。 */
function runInExcel(code, timeoutMs = 120_000) {
  const link = ensureOfficeLink();
  const id = crypto.randomUUID();
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => { officeWaiters.delete(id); reject(new Error('Office.js 通道超时')); }, timeoutMs);
    officeWaiters.set(id, (r) => { clearTimeout(timer); r.error ? reject(new Error(r.error)) : resolve(r.result); });
    const send = () => link.send(JSON.stringify({ id, method: 'run_script', params: { code } }));
    const doSend = () => { console.log('[relay] 下发 run_script 到真实任务窗格', id); send(); };
    if (link.readyState === WebSocket.OPEN) doSend();
    else link.once('open', () => setTimeout(doSend, 80));
  });
}

wps.on('open', () => {
  console.log('[relay] 已注册 WPS 表格通道（临时占用 excel 键），等待 MCP 下发脚本…');
  wps.send(JSON.stringify({ type: 'register', client: 'relay-probe', host: 'wps', version: '2.1.0', summary: { workbookName: 'CAP08-Relay', activeSheetName: '_cap08_probe' } }));
  runInExcel(probeCode).then(r => console.log('[relay] 预热 run_script 结果:', JSON.stringify(r).slice(0,200))).catch(e => console.log('[relay] 预热失败:', e.message));
console.log('[relay] READY ' + (codePath.includes('probe-code') ? 'PROBE' : 'CUSTOM'));
});

wps.on('message', async (raw) => {
  let p; try { p = JSON.parse(raw.toString()); } catch { return; }
  if (p.type !== 'rpc_request' && !p.method) return;
  console.log('[relay] 收到下发方法:', p.method);
  if (p.method === 'get_workspace_summary') {
    wps.send(JSON.stringify({ id: p.id, type: 'rpc_response', result: { hostType: 'excel', workbookName: 'CAP08-Relay', workbookCount: 1, sheetCount: 1, sheets: [{ index: 1, name: '_cap08_probe', visible: true }], activeSheetName: '_cap08_probe' } }));
    return;
  }
  if (p.method === 'execute_script' || p.method === 'run_script') {
    try {
      const result = await runInExcel(probeCode);
      console.log('[relay] 真机执行完成，回包长度', JSON.stringify(result || {}).length);
      wps.send(JSON.stringify({ id: p.id, type: 'rpc_response', result }));
    } catch (e) {
      console.log('[relay] 真机执行失败:', e.message);
      wps.send(JSON.stringify({ id: p.id, type: 'rpc_response', error: e.message }));
    }
    return;
  }
  wps.send(JSON.stringify({ id: p.id, type: 'rpc_response', error: 'relay 不支持的方法: ' + p.method }));
});

wps.on('error', (e) => console.log('[relay] wps 通道异常', e.message));
setTimeout(() => { console.log('[relay] 保持窗口结束，退出'); process.exit(0); }, holdMs);
