import fs from 'node:fs'; import path from 'node:path'; import os from 'node:os'; import WebSocket from 'ws';
const token = fs.readFileSync(path.join(os.homedir(),'.wps-bridge','token'),'utf8').trim();
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: {} }));
  ws.on('message', r => { const s = String(r); console.log('MSG', s.slice(0,200));
    if (s.includes('hot-reload') || s.includes('run_script')) { ws.close(); }
  });
  setTimeout(async () => {
    const t = await fetch('http://127.0.0.1:19890/api/v1/office/reload', { method:'POST', headers: { Authorization: 'Bearer ' + token } }).then(r=>r.text());
    console.log('reload resp:', t.slice(0,100));
  }, 200);
});
ws.on('close', (c,r) => console.log('closed', c, r?.toString?.()));
setTimeout(()=>process.exit(0), 15000);
