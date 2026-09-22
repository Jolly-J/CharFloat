// CAP-08 真机探测运行器。
//
// 作用：本机 MCP 工具面里没有任何入口能把任意 Office.js 代码送进 Excel 任务窗格
// （excel_* 路由表不含形状方法；office_execute_script 走 macOS JXA / Windows COM）。
// 这里用 bridge 的「加载项连接」协议注册一个**临时 ms-excel 通道**，把探测代码
// 作为 run_script 方法发过去，由真实 Office.js 任务窗格执行（run_script 分支本来就存在），
// 再把结果写成 JSON 供后续读取。执行完毕即断开，真实任务窗格会自动重连。
//
// 用法：node tmp/cap08/probe-runner.mjs [probe-code.js 路径]
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import WebSocket from 'ws';

const BRIDGE = 'ws://127.0.0.1:19890/office-addon';
const token = fs.readFileSync(path.join(os.homedir(), '.wps-bridge', 'token'), 'utf8').trim();

const codePath = process.argv[2] || path.join(import.meta.dirname, 'probe-code.js');
const code = fs.readFileSync(codePath, 'utf8');

const ws = new WebSocket(BRIDGE, { headers: { Origin: 'https://localhost:19890' } });

const done = (payload, exitCode) => {
  try { ws.close(); } catch { /* already closed */ }
  const out = JSON.stringify(payload, null, 2);
  fs.writeFileSync(path.join(import.meta.dirname, 'probe-result.json'), out, 'utf8');
  console.log(out);
  process.exit(exitCode);
};

const timer = setTimeout(() => done({ ok: false, error: '整体超时 180s，未收到 rpc_response' }, 1), 180_000);

ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: { workbookName: 'CAP08-Probe', activeSheetName: '_cap08_probe' } }));
  setTimeout(() => {
    ws.send(JSON.stringify({ id: 'cap08-probe-1', method: 'run_script', params: { code } }));
    console.log('[probe] 已通过 Office.js 通道下发 run_script');
  }, 400);
});

ws.on('message', (raw) => {
  let p;
  try { p = JSON.parse(raw.toString()); } catch { return; }
  if (p.type !== 'rpc_response') return;
  clearTimeout(timer);
  done({ ok: !p.error, token, response: p }, p.error ? 1 : 0);
});

ws.on('error', (err) => { clearTimeout(timer); done({ ok: false, error: 'WS 异常: ' + err.message }, 1); });
