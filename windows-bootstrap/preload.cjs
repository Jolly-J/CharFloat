const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('setup', {
  start: () => ipcRenderer.invoke('setup:start'),
  launch: () => ipcRenderer.invoke('setup:launch'),
  close: () => ipcRenderer.invoke('setup:close'),
  painted: token => ipcRenderer.invoke('setup:painted', token),
  telemetry: (event, data) => ipcRenderer.invoke('setup:telemetry', event, data),
  listen: callback => { ipcRenderer.on('setup:state', (_event, state) => callback(state)); }
});
