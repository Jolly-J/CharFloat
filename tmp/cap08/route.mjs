import fs from 'node:fs'; import path from 'node:path'; import os from 'node:os'; import WebSocket from 'ws';
const token = fs.readFileSync(path.join(os.homedir(),'.wps-bridge','token'),'utf8').trim();
const ws = new WebSocket('ws://127.0.0.1:19890/office-addon', { headers: { Origin: 'https://localhost:19890' } });
ws.on('open', () => {
  ws.send(JSON.stringify({ type: 'register', client: 'ms-excel-addon', host: 'microsoft', version: '9.9.9', summary: { workbookName: 'ROUTETEST' } }));
  setTimeout(async () => {
    const r = await fetch('http://127.0.0.1:19890/api/v1/tool/call', { method:'POST', headers:{'Content-Type':'application/json',Authorization:'Bearer '+token}, body: JSON.stringify({name:'excel_get_sheet_outline',arguments:{host:'microsoft',sheetName:'Sheet1'},clientName:'CAP08'})});
    console.log('OFFICE call ->', (await r.text()).slice(0,200));
  }, 600);
  setTimeout(async () => {
    const r = await fetch('http://127.0.0.1:19890/api/v1/tool/call', { method:'POST', headers:{'Content-Type':'application/json',Authorization:'Bearer '+token}, body: JSON.stringify({name:'wps_get_workspace_summary',arguments:{},clientName:'CAP08'})});
    console.log('WPS call ->', (await r.text()).slice(0,160));
  }, 1200);
});
ws.on('message', r => console.log('FAKE-PANE RECEIVED:', String(r).slice(0,300)));
ws.on('close', (c,r) => console.log('closed', c, r?.toString?.()));
setTimeout(()=>process.exit(0), 12000);
