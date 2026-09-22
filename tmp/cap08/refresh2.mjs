import fs from 'node:fs'; import path from 'node:path'; import os from 'node:os'; import WebSocket from 'ws';
const token = fs.readFileSync(path.join(os.homedir(),'.wps-bridge','token'),'utf8').trim();
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: {} }));
  ws.on('message', r => { if (String(r).includes('hot-reload')) { ws.close(); } });
  setTimeout(async () => { await fetch('http://127.0.0.1:19890/api/v1/office/reload', { method:'POST', headers: { Authorization: 'Bearer ' + token } }).then(r=>r.text()).then(t=>console.log('reload:', t.slice(0,60))); }, 150);
});
setTimeout(()=>process.exit(0), 12000);
