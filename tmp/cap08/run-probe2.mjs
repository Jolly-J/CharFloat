import fs from 'node:fs'; import path from 'node:path'; import os from 'node:os'; import WebSocket from 'ws';
const token = fs.readFileSync(path.join(os.homedir(), '.wps-bridge', 'token'), 'utf8').trim();
const code = `window.__cap08Mark = 'C08-' + Date.now(); return await (async()=>{ const r = await Excel.run(async (ctx)=>{ const s = ctx.workbook.worksheets.getActiveWorksheet(); const c = s.getRange('A1'); c.values = [[window.__cap08Mark]]; await ctx.sync(); return {sheet: s.name}; }); return {mark: window.__cap08Mark, r}; })();`;
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
let closed = null;
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: {} }));
  setTimeout(() => { ws.send(JSON.stringify({ id: 'rp1', method: 'run_script', params: { code } })); }, 60);
});
ws.on('message', r => console.log('MSG', String(r).slice(0,400)));
ws.on('close', (c, r) => { closed = c + ' ' + (r?.toString?.() || ''); console.log('closed', closed); });
setTimeout(async () => {
  console.log('--- 6s 后读回 A1（走 excel_read_range host=microsoft）---');
  const resp = await fetch('http://127.0.0.1:19890/api/v1/tool/call', { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + token }, body: JSON.stringify({ name: 'excel_read_range', arguments: { host: 'microsoft', address: 'A1:B2' }, clientName: 'CAP08' }) });
  console.log((await resp.text()).slice(0, 800));
  process.exit(0);
}, 6500);
