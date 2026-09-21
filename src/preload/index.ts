import { contextBridge, ipcRenderer } from 'electron';
const invoke = (name: string) => (...args: any[]) => ipcRenderer.invoke(name, ...args);
contextBridge.exposeInMainWorld('api', {
  platform: process.platform, getStatus: invoke('get-status'), startService: invoke('service-start'), stopService: invoke('service-stop'),
  diagnose: invoke('diagnose'), officeStatus: invoke('office-status'), getAppInfo: invoke('get-app-info'),
  getAuditRecords: invoke('get-audit-records'), getAuditRecord: invoke('get-audit-record'), rollbackRecord: invoke('rollback-record'), clearAuditRecords: invoke('clear-audit-records'),
  checkAddonStatus: invoke('check-addon-status'), installAddon: invoke('install-addon'),
  checkOfficeAddonStatus: invoke('check-office-addon-status'), installOfficeAddon: invoke('install-office-addon'),
  detectEnvironment: invoke('installer:detect'), executeInstall: invoke('installer:execute'),
  setTheme: invoke('set-theme'), setLogin: invoke('set-login'), copyText: invoke('copy-text'), openLog: invoke('open-log'), exitApp: invoke('exit-app'),
  onStatusChange: (callback: (value: any) => void) => { const handler = (_: any, value: any) => callback(value); ipcRenderer.on('status-changed', handler); return () => ipcRenderer.removeListener('status-changed', handler); }
});
