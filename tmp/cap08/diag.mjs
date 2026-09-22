import fs from 'node:fs'; import path from 'node:path'; import os from 'node:os'; import WebSocket from 'ws';
const token = fs.readFileSync(path.join(os.homedir(), '.wps-bridge', 'token'), 'utf8').trim();
const code = `return (async()=>{ const r = await Excel.run(async (ctx)=>{ const s = ctx.workbook.worksheets.getActiveWorksheet(); s.getRange('A1').values = [['C08M-' + Date.now()]]; await ctx.sync(); return s.name; }); return {ok:true, sheet:r}; })();`;
let tries = 0; let replied = false;
function attempt(n) {
  const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
  const id = 'diag-' + n;
  ws.on('open', () => {
    ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: {} }));
    setTimeout(() => ws.send(JSON.stringify({ id, method: 'run_script', params: { code } })), 50);
  });
  ws.on('message', r => { if (String(r).includes(id)) { replied = true; console.log('REPLY', n, String(r).slice(0,300)); } else console.log('OTHER', String(r).slice(0,120)); });
  ws.on('close', () => { if (!replied) { if (n < 4) setTimeout(() => attempt(n+1), 2600); else { console.log('no reply after 5 tries'); process.exit(1);} } });
  ws.on('error', e => console.log('err', e.message));
}
attempt(1);
setTimeout(async () => {
  const resp = await fetch('http://127.0.0.1:19890/api/v1/tool/call', { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + token }, body: JSON.stringify({ name: 'excel_read_range', arguments: { host: 'microsoft', address: 'A1' }, clientName: 'CAP08' }) });
  console.log('READBACK', (await resp.text()).slice(0, 300));
  process.exit(0);
}, 20000);
