import { app, BrowserWindow, ipcMain, Tray, Menu, nativeImage, clipboard, nativeTheme, shell } from 'electron';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { ensureService, serviceRequest, probeService } from '../bridge/service-client.js';
import { atomicWrite, runtimeHome, runtimePort, VERSION, resourcePath } from '../bridge/runtime.js';
import { InstallerEngine } from './installer-engine.js';
import { AddonInstaller, OfficeAddonInstaller } from './addon-installer.js';
const here = path.dirname(fileURLToPath(import.meta.url));
process.env.WPS_BRIDGE_RESOURCES ||= app.getAppPath();
const root = app.getAppPath();
app.setPath('userData', path.join(runtimeHome(), 'desktop-state'));
const cliPath = path.join(root, 'dist/bridge/cli.cjs').replace(/app\.asar([\\/])/, 'app.asar.unpacked$1');
InstallerEngine.runtimeEntry = cliPath;
let window: BrowserWindow | null = null, tray: Tray | null = null;
let timer: NodeJS.Timeout | undefined;
let notice = '';
const prefsPath = path.join(runtimeHome(), 'desktop.json');
let prefs: any = {};
try { prefs = JSON.parse(fs.readFileSync(prefsPath, 'utf8')); } catch {}
if (['system', 'light', 'dark'].includes(prefs.theme)) nativeTheme.themeSource = prefs.theme;
function persist() { atomicWrite(prefsPath, JSON.stringify(prefs)); }
function createWindow() {
  if (window) { window.show(); window.focus(); return; }
  window = new BrowserWindow({ width: Math.max(760, Math.min(prefs.width || 880, 1300)), height: Math.max(560, Math.min(prefs.height || 620, 950)), minWidth: 760, minHeight: 560, show: false,
    title: 'Office Agent Bridge', titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    backgroundColor: nativeTheme.shouldUseDarkColors ? '#17191c' : '#f7f8fa',
    webPreferences: { preload: path.join(here, '../preload/index.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true }
  });
  window.setMenuBarVisibility(false);
  if (process.env.VITE_DEV_SERVER_URL) void window.loadURL(process.env.VITE_DEV_SERVER_URL); else void window.loadFile(path.join(here, '../renderer/index.html'));
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  window.webContents.on('will-navigate', event => event.preventDefault());
  window.once('ready-to-show', () => window?.show());
  window.on('resized', () => { if (window && !window.isMaximized()) { const [width, height] = window.getSize(); prefs = { ...prefs, width, height }; persist(); } });
  window.on('closed', () => { window = null; });
}
async function state() {
  try { return { online: true, ...(await serviceRequest('/api/v1/status')), notice: '' }; }
  catch { return { online: false, isWpsConnected: false, components: {}, notice }; }
}
function setupIpc() {
  ipcMain.handle('get-status', state);
  ipcMain.handle('service-start', async () => { try { await ensureService(cliPath); notice = ''; return { success: true }; } catch(e: any) { notice = e.message; return { success: false, message: e.message }; } });
  ipcMain.handle('service-stop', async () => { await serviceRequest('/api/v1/service/stop', {}); return { success: true }; });
  ipcMain.handle('diagnose', () => serviceRequest('/api/v1/tool/call', { name: 'bridge_diagnose', arguments: {} }).then(r => r.data));
  ipcMain.handle('office-status', () => serviceRequest('/api/v1/tool/call', { name: 'office_get_status', arguments: {} }).then(r => r.data));
  ipcMain.handle('get-audit-records', (_e, options = {}) => serviceRequest('/api/v1/tool/call', { name: 'wps_get_audit_history', arguments: { ...options, limit: 20, view: 'summary' } }).then(r => r.data));
  ipcMain.handle('clear-audit-records', () => serviceRequest('/api/v1/tool/call', { name: 'wps_clear_audit_history', arguments: {} }).then(r => r.data));
  ipcMain.handle('get-audit-record', (_e, id: string) => serviceRequest('/api/v1/tool/call', { name: 'wps_get_audit_record', arguments: { auditId: id } }).then(r => r.data));
  ipcMain.handle('rollback-record', (_e, id: string) => serviceRequest('/api/v1/tool/call', { name: 'wps_rollback', arguments: { auditId: id }, sessionId: 'desktop' }).then(r => r.data));
  ipcMain.handle('check-addon-status', () => AddonInstaller.checkStatus());
  ipcMain.handle('install-addon', () => AddonInstaller.install());
  ipcMain.handle('check-office-addon-status', () => OfficeAddonInstaller.checkStatus());
  ipcMain.handle('install-office-addon', () => OfficeAddonInstaller.install());
  ipcMain.handle('installer:detect', () => InstallerEngine.detectEnvironment());
  ipcMain.handle('installer:execute', (_e, options) => InstallerEngine.executeInstall(options));
  ipcMain.handle('get-app-info', () => ({ version: VERSION, port: runtimePort(), home: runtimeHome(), platform: process.platform, theme: prefs.theme || 'system', login: app.getLoginItemSettings().openAtLogin, config: { mcpServers: { 'office-agent-bridge': InstallerEngine.configEntry(), 'wps-bridge': InstallerEngine.configEntry() } } }));
  ipcMain.handle('set-theme', (_e, theme) => { if (!['system','light','dark'].includes(theme)) throw new Error('主题无效'); nativeTheme.themeSource = theme; prefs.theme = theme; persist(); return true; });
  ipcMain.handle('set-login', (_e, enabled) => { if (!app.isPackaged) throw new Error('登录启动仅在安装版中可用'); app.setLoginItemSettings({ openAtLogin: Boolean(enabled), args: ['--background'] }); return app.getLoginItemSettings().openAtLogin; });
  ipcMain.handle('copy-text', (_e, text) => { if (typeof text !== 'string' || text.length > 100000) throw new Error('文本无效'); clipboard.writeText(text); return true; });
  ipcMain.handle('open-log', () => { const file = path.join(runtimeHome(), 'service.log'); if (fs.existsSync(file)) return shell.openPath(file); return '尚无日志'; });
  ipcMain.handle('exit-app', () => app.quit());
}
if (!app.requestSingleInstanceLock()) app.quit();
else {
  app.on('second-instance', createWindow);
  app.whenReady().then(async () => {
    setupIpc();
    atomicWrite(path.join(runtimeHome(), 'installation.json'), JSON.stringify({ executable: process.execPath, cli: cliPath, resources: root, version: VERSION }));
    const icon = nativeImage.createFromPath(resourcePath('resources/trayTemplate.png')).resize({ width: 22, height: 22 }); icon.setTemplateImage(true);
    tray = new Tray(icon); tray.setToolTip('Office Agent Bridge');
    tray.setContextMenu(Menu.buildFromTemplate([{ label: '打开 Office Agent Bridge', click: createWindow }, { label: '停止服务', click: () => { void serviceRequest('/api/v1/service/stop', {}).catch(e => { notice = e.message; }); } }, { type: 'separator' }, { label: '退出管理窗口（后台继续运行）', click: () => app.quit() }]));
    tray.on('click', createWindow);
    if (!process.argv.includes('--background')) createWindow();
    ensureService(cliPath).catch(e => { notice = e.message; });
    timer = setInterval(async () => { if (window && !window.isDestroyed()) window.webContents.send('status-changed', await state()); }, 2000);
    app.on('activate', createWindow);
  });
  app.on('window-all-closed', () => {});
  app.on('before-quit', () => { clearInterval(timer); tray?.destroy(); });
}
