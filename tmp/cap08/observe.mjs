import fs from 'node:fs'; import path from 'node:path'; import os from 'node:os'; import crypto from 'node:crypto'; import WebSocket from 'ws';
const token = fs.readFileSync(path.join(os.homedir(),'.wps-bridge','token'),'utf8').trim();
const label = process.argv[2] || 'OBS';
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
let n = 0;
async function triggerReload() {
  await fetch('http://127.0.0.1:19890/api/v1/office/reload', { method: 'POST', headers: { Authorization: 'Bearer ' + token } }).then(r=>r.text()).then(t=>console.log('[obs] reload:', t.slice(0,80))).catch(e=>console.log('[obs] reload err', e.message));
}
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: {} }));
  triggerReload();
  setTimeout(() => {
    const id = crypto.randomUUID();
    ws.send(JSON.stringify({ id, method: 'run_script', params: {} }));  // 故意不带 code → 应回"缺少必要参数"
    console.log('[obs] 发送 run_script(空)，id=', id);
  }, 3000);
});
ws.on('message', r => console.log('[obs] MSG', ++n, String(r).slice(0,300)));
ws.on('close', (c,r) => console.log('[obs] closed', c, r?.toString?.()));
setTimeout(async () => {
  const s = await fetch('http://127.0.0.1:19890/api/v1/tool/call', { method:'POST', headers:{'Content-Type':'application/json',Authorization:'Bearer '+token}, body: JSON.stringify({name:'excel_read_range',arguments:{host:'microsoft',address:'A1:B1'},clientName:'CAP08'})});
  console.log('[obs] readback', (await s.text()).slice(0,200));
  process.exit(0);
}, 12000);
