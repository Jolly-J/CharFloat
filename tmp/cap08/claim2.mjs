import WebSocket from 'ws';
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '9.9.9', summary: { workbookName: 'CLAIMTEST' } }));
  setTimeout(async () => {
    const t = await (await fetch('http://127.0.0.1:19890/api/v1/status')).json();
    const c = t.components || t.connection?.components || {};
    console.log('keys:', Object.keys(t));
    console.log('msExcel:', JSON.stringify(c.msExcel).slice(0,300));
  }, 1200);
});
ws.on('close', (c,r) => console.log('closed', c, r?.toString?.()));
setTimeout(()=>process.exit(0), 8000);
