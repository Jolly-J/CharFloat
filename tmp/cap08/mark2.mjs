import fs from 'node:fs'; import path from 'node:path'; import os from 'node:os'; import WebSocket from 'ws';
const code = `var c = context.workbook.worksheets.getItem('Sheet1'); c.getRange('C1').values = [['C08MARK-' + Date.now()]]; await context.sync(); return {ok:true};`;
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: {} }));
  setTimeout(() => ws.send(JSON.stringify({ id: 'm2', method: 'run_script', params: { code, sheetName: 'Sheet1' } })), 300);
});
ws.on('message', r => console.log('MSG', String(r).slice(0,400)));
ws.on('close', (c,r) => console.log('closed', c, r?.toString?.()));
setTimeout(()=>process.exit(0), 30000);
