import fs from 'node:fs'; import path from 'node:path'; import os from 'node:os'; import WebSocket from 'ws';
const token = fs.readFileSync(path.join(os.homedir(),'.wps-bridge','token'),'utf8').trim();
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: {} }));
  setTimeout(() => ws.send(JSON.stringify({ id: 'tr1', method: 'run_script', params: { code: "return await (async()=>{ const r = await Excel.run(async(c)=>{ c.workbook.worksheets.getActiveWorksheet().getRange('A1').values=[['TRACE-' + Date.now()]]; await c.sync(); return 'done'; }); return {ok:true, r}; })();" } })), 300);
});
ws.on('message', r => console.log('MSG', String(r).slice(0,300)));
ws.on('close', (c,r) => console.log('closed', c, r?.toString?.()));
setTimeout(async () => {
  const s = await fetch('http://127.0.0.1:19890/api/v1/status', { headers: { Authorization: 'Bearer ' + token } });
  const j = await s.json();
  console.log('msExcel status:', JSON.stringify(j.components?.msExcel || {}).slice(0,500));
  process.exit(0);
}, 12000);
