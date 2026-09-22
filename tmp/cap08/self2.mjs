import crypto from 'node:crypto';
import WebSocket from 'ws';
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '9.9.9', summary: { workbookName: 'CLAIMTEST2' } }));
  setTimeout(() => { ws.send(JSON.stringify({ id: 'rs-1', method: 'run_script', params: { code: "return {probe:'x'};" } })); console.log('t=300 run_script'); }, 300);
  setTimeout(() => { ws.send(JSON.stringify({ id: 'os-1', method: 'get_sheet_outline', params: { sheetName: 'Sheet1' } })); console.log('t=1400 outline'); }, 1400);
});
ws.on('message', r => console.log('MSG', String(r).slice(0,900)));
ws.on('close', (c,r) => console.log('closed', c, r?.toString?.()));
setTimeout(()=>process.exit(0), 12000);
