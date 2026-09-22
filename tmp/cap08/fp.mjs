import fs from 'node:fs'; import path from 'node:path'; import os from 'node:os'; import WebSocket from 'ws';
const token = fs.readFileSync(path.join(os.homedir(),'.wps-bridge','token'),'utf8').trim();
async function status() {
  const r = await fetch('http://127.0.0.1:19890/api/v1/status');
  const t = await r.text();
  return t;
}
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '9.9.9', summary: {} }));
  setTimeout(async () => {
    const r = await fetch('http://127.0.0.1:19890/api/v1/office/reload', { method: 'POST', headers: { Authorization: 'Bearer '+token } });
    console.log('reload:', (await r.text()).slice(0,80));
  }, 400);
});
fs.writeFileSync('/tmp/status.json', await status());
console.log('status keys:', Object.keys(JSON.parse(fs.readFileSync('/tmp/status.json','utf8'))).slice(0,20));
setTimeout(()=>process.exit(0), 12000);
