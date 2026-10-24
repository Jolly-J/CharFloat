import { contextBridge, ipcRenderer } from 'electron';
const invoke = (name: string) => (...args: any[]) => ipcRenderer.invoke(name, ...args);
contextBridge.exposeInMainWorld('api', {
  platform: process.platform, getStatus: invoke('get-status'), startService: invoke('service-start'), stopService: invoke('service-stop'),
  diagnose: invoke('diagnose'), officeStatus: invoke('office-status'), getAppInfo: invoke('get-app-info'),
  getAuditRecords: invoke('get-audit-records'), getAuditRecord: invoke('get-audit-record'), rollbackRecord: invoke('rollback-record'), clearAuditRecords: invoke('clear-audit-records'),
  checkAddonStatus: invoke('check-addon-status'), installAddon: invoke('install-addon'),
  checkOfficeAddonStatus: invoke('check-office-addon-status'), installOfficeAddon: invoke('install-office-addon'),
  detectEnvironment: invoke('installer:detect'), executeInstall: invoke('installer:execute'),
  permissionInfo: invoke('permission:full-disk-access-info'), openFullDiskAccess: invoke('permission:open-full-disk-access'),
  setTheme: invoke('set-theme'), setLogin: invoke('set-login'), copyText: invoke('copy-text'), openLog: invoke('open-log'), exitApp: invoke('exit-app'),
  markDoubaoConfigured: invoke('mark-doubao-configured'),
  openAgentApp: invoke('open-agent-app'),
  openClientDocumentation: invoke('client:documentation'),
  updateStatus: invoke('update:status'), takeUpdateFocus: invoke('update:take-focus'), checkUpdate: invoke('update:check'), downloadUpdate: invoke('update:download'), cancelUpdate: invoke('update:cancel'), installUpdate: invoke('update:install'),
  onUpdateOpen: (callback: () => void) => { const handler = () => callback(); ipcRenderer.on('update-open', handler); return () => ipcRenderer.removeListener('update-open', handler); },
  updateNoticeInfo: invoke('update-notice:info'), updateNoticeReady: invoke('update-notice:ready'), dismissUpdateNotice: invoke('update-notice:dismiss'), openUpdateFromNotice: invoke('update-notice:open'),
  onUpdateChange: (callback: (value: any) => void) => { const handler = (_: any, value: any) => callback(value); ipcRenderer.on('update-changed', handler); return () => ipcRenderer.removeListener('update-changed', handler); },
  startAppDrag: () => ipcRenderer.send('permission:drag-start'),
  appIcon: invoke('permission:app-icon'),
  openPermissionGuide: invoke('permission:open-guide'),
  closePermissionWindow: invoke('permission:close-window'),
  retryPermissionInstall: invoke('permission:retry-install'),
  onPermissionIssue: (callback: (value: any) => void) => { const handler = (_: any, value: any) => callback(value); ipcRenderer.on('permission:issue', handler); return () => ipcRenderer.removeListener('permission:issue', handler); },
  onStatusChange: (callback: (value: any) => void) => { const handler = (_: any, value: any) => callback(value); ipcRenderer.on('status-changed', handler); return () => ipcRenderer.removeListener('status-changed', handler); }
});
