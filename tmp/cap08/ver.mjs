import WebSocket from 'ws';
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
ws.on('message', r => console.log('MSG', String(r).slice(0,600)));
ws.on('close', (c,r) => console.log('closed', c, r?.toString?.()));
ws.on('open', () => {
  console.log('open');
  setTimeout(() => {
    ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '2.1.0', summary: {}, buildFingerprint: 'verifier' }));
  }, 3000);  // 等真实任务窗格先注册
  setTimeout(() => { console.log('t=3500 sending run_script'); ws.send(JSON.stringify({ id: 'v1', method: 'run_script', params: { code: "return {ver:'v1'};" } })); }, 3500);
});
setTimeout(()=>process.exit(0), 30000);
