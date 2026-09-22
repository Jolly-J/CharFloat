import WebSocket from 'ws';
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: {} }));
  setTimeout(() => { ws.send(JSON.stringify({ id: 'self-1', method: 'get_sheet_outline', params: { sheetName: 'Sheet1' } })); console.log('sent self outline'); }, 60);
});
ws.on('message', r => console.log('MSG', String(r).slice(0,800)));
ws.on('close', (c,r) => console.log('closed', c, r?.toString?.()));
setTimeout(()=>process.exit(0), 15000);
