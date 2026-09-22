import fs from 'node:fs'; import path from 'node:path'; import os from 'node:os'; import WebSocket from 'ws';
const token = fs.readFileSync(path.join(os.homedir(), '.wps-bridge', 'token'), 'utf8').trim();
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
const t0 = Date.now();
ws.on('open', () => {
  console.log('open', Date.now()-t0);
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: { workbookName: 'x', activeSheetName: 'y' } }));
  ws.send(JSON.stringify({ id: 'w1', method: 'get_workspace_summary', params: {} }));
});
ws.on('message', r => console.log('recv', Date.now()-t0, String(r).slice(0,200)));
ws.on('close', (c, r) => { console.log('closed', Date.now()-t0, c, r?.toString?.()); process.exit(0); });
setTimeout(() => { console.log('still open at', Date.now()-t0); process.exit(0); }, 20000);
