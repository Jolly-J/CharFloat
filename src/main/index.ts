import { app, BrowserWindow, ipcMain, Tray, Menu, nativeImage, clipboard, nativeTheme, shell } from 'electron';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { ensureService, serviceRequest, probeService } from '../bridge/service-client.js';
import { atomicWrite, runtimeHome, runtimePort, VERSION, resourcePath, appendServiceLog, getToken } from '../bridge/runtime.js';
import { FULL_DISK_ACCESS_URL, appToAuthorize } from './permissions.js';
import { openPermissionWindow, closePermissionWindow, startAppDrag, getPermissionIssue, loadAppIcon } from './permission-window.js';
import { InstallerEngine } from './installer-engine.js';
import { AddonInstaller, OfficeAddonInstaller } from './addon-installer.js';
import { ensureAddonUpToDate } from './addon-autoupgrade.js';
const here = path.dirname(fileURLToPath(import.meta.url));
process.env.WPS_BRIDGE_RESOURCES ||= app.getAppPath();
const root = app.getAppPath();
app.setPath('userData', path.join(runtimeHome(), 'desktop-state'));
const cliPath = path.join(root, 'dist/bridge/cli.cjs').replace(/app\.asar([\\/])/, 'app.asar.unpacked$1');
InstallerEngine.runtimeEntry = cliPath;
let window: BrowserWindow | null = null, tray: Tray | null = null;
let timer: NodeJS.Timeout | undefined;
let notice = '';
// 最近一次加载项自动升级的结果（供界面展示）
let addonUpgradeResult: Awaited<ReturnType<typeof ensureAddonUpToDate>> | null = null;
const prefsPath = path.join(runtimeHome(), 'desktop.json');
let prefs: any = {};
try { prefs = JSON.parse(fs.readFileSync(prefsPath, 'utf8')); } catch {}
if (['system', 'light', 'dark'].includes(prefs.theme)) nativeTheme.themeSource = prefs.theme;
function persist() { atomicWrite(prefsPath, JSON.stringify(prefs)); }
function createWindow() {
  if (window) { window.show(); window.focus(); return; }
  let appIcon: Electron.NativeImage | undefined;
  if (process.platform !== 'darwin') {
    try {
      appIcon = nativeImage.createFromPath(resourcePath('resources/icon.png'));
    } catch {}
  }
  window = new BrowserWindow({ width: Math.max(780, Math.min(prefs.width || 880, 1300)), height: Math.max(560, Math.min(prefs.height || 620, 950)), minWidth: 780, minHeight: 560, show: false,
    ...(appIcon ? { icon: appIcon } : {}),
    title: '字浮 CharFloat', titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
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
  // 已部署加载项 vs 客户端包内构建的构建指纹比对（"版本号相同 ≠ 构建相同"）
  ipcMain.handle('check-addon-freshness', () => AddonInstaller.checkBuildFreshness());
  ipcMain.handle('get-addon-upgrade-result', () => addonUpgradeResult);
  // 需要时手动触发一次"检测 → 重新部署 → 触发重载"，与启动时的自动流程同一条路径
  ipcMain.handle('ensure-addon-uptodate', async (_e, opts?: { force?: boolean }) => {
    appendServiceLog('IPC', `收到加载项升级请求 (force=${Boolean(opts?.force)})`);
    addonUpgradeResult = await ensureAddonUpToDate({ force: Boolean(opts?.force) });
    return addonUpgradeResult;
  });
  // 安装器只做**纯识别**（不依赖 electron）；"该授权哪个 App"在这里补上，保证 installer 可在普通 Node 测试里加载
  let lastPermissionRetry: (() => Promise<any>) | null = null;
  const withPermissionTarget = (r: any) => {
    if (!r?.permissionIssue) return r;
    const issue = { ...r.permissionIssue, appToAuthorize: appToAuthorize(), isDev: !app.isPackaged, label: r.message };
    // 方案 B：直接开置顶引导浮窗，它能在系统设置面板上方保持可见，便于拖拽图标
    appendServiceLog('PermissionWindow', `权限失败，准备打开引导浮窗: ${issue?.targetPath || '(无路径)'}`);
    try { openPermissionWindow(issue); appendServiceLog('PermissionWindow', '引导浮窗已创建'); }
    catch (e: any) { appendServiceLog('PermissionWindow', `打开引导浮窗失败: ${e?.message || e}`); }
    return { ...r, permissionIssue: issue };
  };
  ipcMain.handle('install-addon', async () => {
    appendServiceLog('IPC', '收到前端一键安装 / 升级 WPS 加载项请求');
    lastPermissionRetry = async () => withPermissionTarget(await AddonInstaller.install());
    return withPermissionTarget(await AddonInstaller.install());
  });
  ipcMain.handle('check-office-addon-status', () => OfficeAddonInstaller.checkStatus());
  ipcMain.handle('install-office-addon', async () => {
    lastPermissionRetry = async () => withPermissionTarget(await OfficeAddonInstaller.install());
    return withPermissionTarget(await OfficeAddonInstaller.install());
  });
  ipcMain.handle('installer:detect', () => InstallerEngine.detectEnvironment());
  ipcMain.handle('installer:execute', (_e, options) => InstallerEngine.executeInstall(options));
  // macOS 完全磁盘访问权限引导：打开设置面板 + 告知当前运行模式该授权的 App。
  ipcMain.handle('permission:full-disk-access-info', () => ({ url: FULL_DISK_ACCESS_URL, appToAuthorize: appToAuthorize(), isDev: !app.isPackaged, platform: process.platform }));
  ipcMain.handle('permission:open-full-disk-access', async () => { await shell.openExternal(FULL_DISK_ACCESS_URL); return { success: true }; });
  // 必须用 ipcMain.on（配合渲染层 ipcRenderer.send）：invoke 的异步往返会错过拖拽会话
  ipcMain.on('permission:drag-start', event => { try { startAppDrag(event.sender); } catch (e: any) { appendServiceLog('PermissionWindow', `发起拖拽失败: ${e?.message || e}`); } });
  ipcMain.handle('permission:app-icon', () => loadAppIcon());
  ipcMain.handle('permission:open-guide', () => { openPermissionWindow(getPermissionIssue()); return { success: true }; });
  ipcMain.handle('permission:close-window', () => { closePermissionWindow(); return { success: true }; });
  ipcMain.handle('permission:retry-install', async () => lastPermissionRetry ? lastPermissionRetry() : { success: false, message: '没有可重试的安装' });
  ipcMain.handle('get-app-info', () => ({ version: VERSION, port: runtimePort(), home: runtimeHome(), token: (() => { try { return getToken(); } catch { return ''; } })(), platform: process.platform, theme: prefs.theme || 'system', login: app.getLoginItemSettings().openAtLogin, config: { mcpServers: { 'charfloat': InstallerEngine.configEntry() } } }));
  ipcMain.handle('set-theme', (_e, theme) => { if (!['system','light','dark'].includes(theme)) throw new Error('主题无效'); nativeTheme.themeSource = theme; prefs.theme = theme; persist(); return true; });
  ipcMain.handle('set-login', (_e, enabled) => { if (!app.isPackaged) throw new Error('登录启动仅在安装版中可用'); app.setLoginItemSettings({ openAtLogin: Boolean(enabled), args: ['--background'] }); return app.getLoginItemSettings().openAtLogin; });
  ipcMain.handle('copy-text', (_e, text) => { if (typeof text !== 'string' || text.length > 100000) throw new Error('文本无效'); clipboard.writeText(text); return true; });
  ipcMain.handle('open-log', () => {
    const file = path.join(runtimeHome(), 'service.log');
    if (!fs.existsSync(file)) {
      appendServiceLog('System', `日志文件已初始化 (平台: ${process.platform}, 版本: ${VERSION})`);
    }
    shell.openPath(file);
    return true;
  });
  ipcMain.handle('mark-doubao-configured', (_e, configured = true) => {
    prefs.doubaoConfigured = Boolean(configured);
    persist();
    return true;
  });
  ipcMain.handle('open-agent-app', async (_e, agentId: string) => {
    try {
      if (agentId === 'workbuddy') {
        try {
          await shell.openExternal('workbuddy://connectors');
        } catch {}
        if (process.platform === 'darwin') {
          const { execFile } = await import('node:child_process');
          execFile('osascript', ['-e', 'tell application "WorkBuddy" to activate'], () => {});
        }
        return { success: true };
      }
      if (agentId === 'doubao') {
        if (process.platform === 'darwin') {
          await shell.openPath('/Applications/DoubaoWork.app');
          const { execFile } = await import('node:child_process');
          execFile('osascript', ['-e', 'tell application "DoubaoWork" to activate'], () => {});
          return { success: true };
        } else {
          try {
            await shell.openExternal('doubaowork://doubaowork-settings');
            return { success: true };
          } catch {}
        }
      }
      return { success: false, message: '未找到该客户端的快捷唤起入口' };
    } catch (e: any) {
      return { success: false, message: e.message };
    }
  });
  ipcMain.handle('exit-app', () => app.quit());
}
if (!app.requestSingleInstanceLock()) app.quit();
else {
  app.on('second-instance', createWindow);
  app.whenReady().then(async () => {
    setupIpc();
    const isMac = process.platform === 'darwin';
    let trayIcon: Electron.NativeImage;
    if (isMac) {
      trayIcon = nativeImage.createFromPath(resourcePath('resources/trayTemplate.png'));
      trayIcon.setTemplateImage(true);
    } else {
      trayIcon = nativeImage.createFromPath(resourcePath('resources/tray-win.png'));
    }
    tray = new Tray(trayIcon); tray.setToolTip('字浮 CharFloat');
    tray.setContextMenu(Menu.buildFromTemplate([{ label: '打开 字浮 CharFloat', click: createWindow }, { label: '停止服务', click: () => { void serviceRequest('/api/v1/service/stop', {}).catch(e => { notice = e.message; }); } }, { type: 'separator' }, { label: '退出管理窗口（后台继续运行）', click: () => app.quit() }]));
    tray.on('click', createWindow);
    if (!process.argv.includes('--background')) createWindow();
    // 启动即对齐：客户端包内的加载项若比 WPS 里部署的新，自动重新部署并触发重载。
    // 串行在 ensureService 之后——重载要靠桥接把 reload 送到加载项。
    ensureService(cliPath)
      .then(() => ensureAddonUpToDate())
      .then(r => {
        addonUpgradeResult = r;
        appendServiceLog('AutoUpgrade', r.message);
      })
      .catch(e => { notice = e.message; });
    timer = setInterval(async () => { if (window && !window.isDestroyed()) window.webContents.send('status-changed', await state()); }, 2000);
    app.on('activate', createWindow);
  });
  app.on('window-all-closed', () => {});
  app.on('before-quit', () => { clearInterval(timer); tray?.destroy(); });
}
