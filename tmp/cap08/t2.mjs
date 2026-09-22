import fs from 'node:fs'; import path from 'node:path'; import os from 'node:os'; import crypto from 'node:crypto'; import WebSocket from 'ws';
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
const id = crypto.randomUUID();
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: {} }));
  setTimeout(() => { ws.send(JSON.stringify({ id, method: 'run_script', params: { code: "return {probe:'cap08-t2'};" } })); console.log('SENT-RUNSCRIPT'); }, 40);
  // 2.2s 后（仍在窗口内）再要一次 outline 诊断
  setTimeout(() => { ws.send(JSON.stringify({ id: crypto.randomUUID(), method: 'get_sheet_outline', params: { sheetName: 'Sheet1' } })); console.log('SENT-OUTLINE'); }, 2200);
});
ws.on('message', r => console.log('MSG', String(r).slice(0,700)));
ws.on('close', (c,r) => console.log('closed', c, r?.toString?.()));
setTimeout(()=>process.exit(0), 20000);
