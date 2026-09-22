import fs from 'node:fs'; import path from 'node:path'; import os from 'node:os'; import WebSocket from 'ws';
const token = fs.readFileSync(path.join(os.homedir(),'.wps-bridge','token'),'utf8').trim();
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '9.9.9', summary: {} }));
});
ws.on('message', r => console.log('MSG', String(r).slice(0,120)));
ws.on('close', (c,r) => console.log('closed', c, r?.toString?.()));
// 4s 后（真实任务窗格已夺回 ms-excel 键）再触发重载
setTimeout(async () => {
  const r = await fetch('http://127.0.0.1:19890/api/v1/office/reload', { method: 'POST', headers: { Authorization: 'Bearer '+token } });
  console.log('reload(4s):', (await r.text()).slice(0,90));
}, 4000);
setTimeout(()=>process.exit(0), 20000);
