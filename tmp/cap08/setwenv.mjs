import WebSocket from 'ws';
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: {} }));
  setTimeout(() => { ws.send(JSON.stringify({ id: 'we1', method: 'run_script', params: { code: "window.__cap08Probe = 'CAP08-PROBE-MARKER-' + Date.now(); return window.__cap08Probe;" } })); }, 60);
});
ws.on('message', r => { console.log('recv', String(r).slice(0,300)); process.exit(0); });
ws.on('close', (c, r) => { console.log('closed', c, r?.toString?.()); process.exit(1); });
setTimeout(() => { console.log('timeout'); process.exit(1); }, 8000);
