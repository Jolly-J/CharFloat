import fs from 'node:fs'; import path from 'node:path'; import os from 'node:os'; import WebSocket from 'ws';
const token = fs.readFileSync(path.join(os.homedir(),'.wps-bridge','token'),'utf8').trim();
const code = `await Excel.run(async (context) => { context.workbook.worksheets.getActiveWorksheet().getRange('A1').values = [['SOLO-' + Date.now()]]; await context.sync(); }); return {ok:true};`;
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: {} }));
  setTimeout(() => ws.send(JSON.stringify({ id: 'solo1', method: 'run_script', params: { code } })), 100);
});
ws.on('message', r => console.log('MSG', String(r).slice(0,300)));
ws.on('close', (c,r) => console.log('closed', c, r?.toString?.()));
setTimeout(async () => {
  const s = await fetch('http://127.0.0.1:19890/api/v1/tool/call', { method:'POST', headers:{'Content-Type':'application/json',Authorization:'Bearer '+token}, body: JSON.stringify({name:'excel_read_range',arguments:{host:'microsoft',address:'A1'},clientName:'CAP08'})});
  console.log('READBACK', (await s.text()).slice(0,250));
  process.exit(0);
}, 25000);
