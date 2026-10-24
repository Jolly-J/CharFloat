import { app, BrowserWindow, ipcMain, Tray, Menu, nativeImage, clipboard, nativeTheme, shell, dialog } from 'electron';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { serviceRequest, probeService, pauseAutomaticServiceStart } from '../bridge/service-client.js';
import { DesktopService } from './desktop-service.js';
import { UpdateService } from './update-service.js';
import { UPDATE_TRUST } from './update-trust.js';
import { handoffUpdate } from './update-handoff.js';
import { revokeInstallAuthorization } from './update-install.js';
import { retargetUpdateConfig } from './update-config.js';
import { showUpdateNotice, closeUpdateNotice, setupUpdateNoticeIpc } from './update-notice-window.js';
import { refreshManagedSkills } from './skill-links.js';
import { runDesktopDiagnostics } from './desktop-diagnostics.js';
import { atomicWrite, runtimeHome, runtimePort, VERSION, resourcePath, appendServiceLog, getToken } from '../bridge/runtime.js';
import { FULL_DISK_ACCESS_URL, appToAuthorize } from './permissions.js';
import { openPermissionWindow, closePermissionWindow, startAppDrag, getPermissionIssue, loadAppIcon } from './permission-window.js';
import { InstallerEngine } from './installer-engine.js';
import { findClient } from './client-registry.js';
import { AddonInstaller, OfficeAddonInstaller } from './addon-installer.js';
import { ensureAddonUpToDate } from './addon-autoupgrade.js';
const here = path.dirname(fileURLToPath(import.meta.url));
// A separate Windows process runs only our bundled updater. It never takes the GUI lock,
// creates the desktop window, starts Bridge, or loads a script supplied by the caller.
if (process.argv.includes('--charfloat-update-helper')) {
  if (process.platform !== 'win32') app.exit(1);
  else createRequire(import.meta.url)(path.join(here, 'update-helper.cjs'));
} else {
process.env.WPS_BRIDGE_RESOURCES ||= app.getAppPath();
const root = app.getAppPath();
app.setPath('userData', path.join(runtimeHome(), 'desktop-state'));
const cliPath = path.join(root, 'dist/bridge/cli.cjs').replace(/app\.asar([\\/])/, 'app.asar.unpacked$1');
InstallerEngine.runtimeEntry = cliPath;
let window: BrowserWindow | null = null, tray: Tray | null = null;
let timer: NodeJS.Timeout | undefined;
let updateTimer: NodeJS.Timeout | undefined;
let updateFocusPending = false;
let notice = '';
const desktopService = new DesktopService();
const updates = new UpdateService({ trust: UPDATE_TRUST, context: { version: VERSION, platform: process.platform, arch: process.arch }, cacheDirectory: path.join(runtimeHome(), 'desktop-state', 'updates'), onState: value => {
  if (window && !window.isDestroyed()) window.webContents.send('update-changed', value);
  if (value.phase === 'available') void announceUpdate(value.availableVersion).catch(() => { notice = '更新提醒未能显示，请在客户端查看版本更新。'; });
} });
let quitCompleted = false;
let quitPending: Promise<void> | undefined;
let installing = false;
async function prepareShutdown() {
  clearInterval(timer);
  clearInterval(updateTimer);
  desktopService.beginClosing();
  await pauseAutomaticServiceStart();
  await updates.shutdown();
  await desktopService.stop();
}
// 最近一次加载项自动升级的结果（供界面展示）
let addonUpgradeResult: Awaited<ReturnType<typeof ensureAddonUpToDate>> | null = null;
const prefsPath = path.join(runtimeHome(), 'desktop.json');
let prefs: any = {};
try { prefs = JSON.parse(fs.readFileSync(prefsPath, 'utf8')); } catch {}
if (['system', 'light', 'dark'].includes(prefs.theme)) nativeTheme.themeSource = prefs.theme;
function persist() { atomicWrite(prefsPath, JSON.stringify(prefs)); }
function openUpdateArea() {
  updateFocusPending = true;
  createWindow();
  if (window && !window.webContents.isLoading()) window.webContents.send('update-open');
}
async function announceUpdate(version?: string) {
  if (!version || prefs.notifiedUpdateVersion === version || desktopService.closing) return;
  // Persist before showing: closing the popup or restarting the GUI cannot re-prompt this version.
  prefs.notifiedUpdateVersion = version;
  try { persist(); } catch { notice = '更新提醒记录未能保存，请检查用户目录权限。'; }
  await showUpdateNotice({ version, currentVersion: VERSION, notes: updates.getState().notes || '' });
}
function startUpdateChecks() {
  clearInterval(updateTimer);
  const check = () => {
    if (!desktopService.closing && !updates.isWorking && updates.getState().phase !== 'ready') void updates.check().catch(() => {});
  };
  check();
  updateTimer = setInterval(check, 6 * 60 * 60 * 1000); updateTimer.unref();
}
function createWindow() {
  if (quitCompleted) return;
  if (window) { if (window.isMinimized()) window.restore(); window.show(); window.focus(); return; }
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
  window.webContents.on('did-finish-load', () => { if (updateFocusPending) window?.webContents.send('update-open'); });
  window.once('ready-to-show', () => window?.show());
  window.on('close', event => { if (!quitCompleted) { event.preventDefault(); window?.hide(); } });
  window.on('resized', () => { if (window && !window.isMaximized()) { const [width, height] = window.getSize(); prefs = { ...prefs, width, height }; persist(); } });
  window.on('closed', () => { window = null; });
}
async function state() {
  try { return { online: true, ...(await serviceRequest('/api/v1/status')), notice, addonStartupChecked: addonUpgradeResult !== null }; }
  catch { return { online: false, isWpsConnected: false, components: {}, notice, addonStartupChecked: addonUpgradeResult !== null }; }
}
function setupIpc() {
  setupUpdateNoticeIpc(openUpdateArea, () => desktopService.closing);
  ipcMain.handle('client:documentation', async (event, id: string) => {
    if (!window || event.sender !== window.webContents || event.senderFrame !== window.webContents.mainFrame) throw new Error('无效的客户端请求来源。');
    const url = typeof id === 'string' ? findClient(id)?.documentation : undefined;
    if (!url || !url.startsWith('https://')) return { success: false, message: '请在该客户端中查看接入说明。' };
    await shell.openExternal(url); return { success: true };
  });
  const ownUpdate = (event: Electron.IpcMainInvokeEvent) => {
    if (!window || event.sender !== window.webContents || event.senderFrame !== window.webContents.mainFrame) throw new Error('无效的更新请求来源。');
    if (installing || quitPending || desktopService.closing) throw new Error('字浮正在退出或准备安装。');
  };
  ipcMain.handle('update:status', event => { ownUpdate(event); return updates.restoreInstallReport(); });
  ipcMain.handle('update:take-focus', event => { ownUpdate(event); const pending = updateFocusPending; updateFocusPending = false; return pending; });
  ipcMain.handle('update:check', event => { ownUpdate(event); return updates.check(); });
  ipcMain.handle('update:download', event => { ownUpdate(event); return updates.download(); });
  ipcMain.handle('update:cancel', event => { ownUpdate(event); updates.cancel(); return { success: true }; });
  ipcMain.handle('update:install', async event => {
    ownUpdate(event);
    if (installing || quitPending || desktopService.closing) throw new Error('字浮正在退出或准备安装。');
    if (!app.isPackaged || !['darwin', 'win32'].includes(process.platform) || !['arm64', 'x64'].includes(process.arch)) throw new Error('在线安装仅适用于正式安装的 Mac/Windows 应用。');
    installing = true;
    let handoff: Awaited<ReturnType<typeof handoffUpdate>> | undefined;
    try {
      const artifact = await updates.verifiedDownload();
      updates.preparingInstall();
      const probe = await probeService();
      if (probe.ready && !desktopService.ownsProcess(probe.info.pid)) throw new Error('检测到独立启动的 Bridge，不能接管或停止它。请先通过它所属的入口退出后台，再重试更新。');
      prefs.updateLogin = app.getLoginItemSettings().openAtLogin; persist();
      await prepareShutdown();
      const target = process.platform === 'darwin' ? path.resolve(process.execPath, '../../..') : path.dirname(process.execPath);
      handoff = await handoffUpdate({ cache: path.join(runtimeHome(), 'desktop-state', 'updates'), artifact, currentVersion: VERSION, target, helper: path.join(here, 'update-helper.cjs'), platform: process.platform as 'darwin' | 'win32', arch: process.arch as 'arm64' | 'x64' });
      quitCompleted = true; closePermissionWindow(); closeUpdateNotice(); tray?.destroy(); app.quit();
      return { success: true };
    } catch (error) {
      if (handoff) { await revokeInstallAuthorization(handoff.file, handoff.ticket.nonce); quitCompleted = false; }
      if (desktopService.closing) {
        try { desktopService.resumeAfterFailedHandoff(); updates.resumeAfterFailedHandoff(); startStatusTimer(); startUpdateChecks(); } catch { /* keep shutdown guard when a child is still alive */ }
      }
      updates.failedInstall(error instanceof Error ? error.message : '更新未完成，请重试。');
      throw error;
    } finally { installing = false; }
  });
  ipcMain.handle('get-status', state);
  ipcMain.handle('service-start', async () => { try { await desktopService.start(cliPath); notice = ''; return { success: true }; } catch(e: any) { notice = e.message; return { success: false, message: e.message }; } });
  ipcMain.handle('service-stop', async () => { await serviceRequest('/api/v1/service/stop', {}); return { success: true }; });
  ipcMain.handle('diagnose', () => runDesktopDiagnostics());
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
function startStatusTimer() {
  clearInterval(timer);
  timer = setInterval(async () => {
    if (!desktopService.closing && window && !window.isDestroyed()) {
      const current = await state();
      if (!desktopService.closing && window && !window.isDestroyed()) window.webContents.send('status-changed', current);
    }
  }, 2000);
}
function migrateUpdatedWindowsPaths() {
  if (!app.isPackaged || process.platform !== 'win32') return;
  const directory = path.dirname(process.execPath);
  try {
    const marker = path.join(directory, '.charfloat-install.json'), st = fs.lstatSync(marker);
    if (!st.isFile() || st.isSymbolicLink() || st.size > 16384) return;
    const receipt = JSON.parse(fs.readFileSync(marker, 'utf8'));
    const previous = receipt.previousDirectory;
    if (receipt.product !== 'charfloat-user-bootstrap-v1' || receipt.version !== VERSION || typeof previous !== 'string') return;
    const root = path.join(fs.realpathSync(process.env.LOCALAPPDATA!), 'Programs', 'CharFloat');
    if (path.dirname(directory) !== root || path.dirname(previous) !== root || !/^\d+\.\d+\.\d+-[a-f0-9]{12}$/.test(path.basename(previous)) || fs.lstatSync(previous).isSymbolicLink()) throw new Error('无法确认旧安装目录，未迁移客户端路径。');
    if (prefs.updateMigratedFrom === previous && prefs.updateMigratedTo === directory) return;
    const errors: string[] = [];
    for (const agent of InstallerEngine.detectEnvironment().agents) {
      if (agent.configKind === 'manual') continue;
      for (const target of agent.configTargets) {
        try { retargetUpdateConfig(target.file, target.kind, previous, directory); }
        catch { errors.push(agent.name); }
      }
    }
    if (prefs.updateLogin === true) app.setLoginItemSettings({ openAtLogin: true, args: ['--background'] });
    if (errors.length) notice = `应用已更新，部分客户端路径未迁移：${[...new Set(errors)].join('、')}。原配置与旧应用均保留，可在授权中心重新配置。`;
    else { prefs.updateMigratedFrom = previous; prefs.updateMigratedTo = directory; delete prefs.updateLogin; persist(); }
  } catch (error: any) { notice = error.message || '更新后的客户端路径迁移未完成，原配置已保留。'; }
}
if (!app.requestSingleInstanceLock()) app.quit();
else {
  app.on('second-instance', createWindow);
  app.whenReady().then(async () => {
    migrateUpdatedWindowsPaths();
    try { refreshManagedSkills(); } catch (error: any) { notice = error.message || '技能资源更新未完成，请重新连接客户端。'; }
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
    tray.setContextMenu(Menu.buildFromTemplate([{ label: '打开 字浮 CharFloat', click: createWindow }, { label: '停止服务', click: () => { void serviceRequest('/api/v1/service/stop', {}).catch(e => { notice = e.message; }); } }, { type: 'separator' }, { label: '退出字浮（停止本应用后台）', click: () => app.quit() }]));
    tray.on('click', createWindow);
    if (!process.argv.includes('--background')) createWindow();
    // 启动即对齐：客户端包内的加载项若比 WPS 里部署的新，自动重新部署并触发重载。
    // 串行在 ensureService 之后——重载要靠桥接把 reload 送到加载项。
    desktopService.start(cliPath)
      .then(() => desktopService.closing ? null : ensureAddonUpToDate())
      .then(r => {
        if (!r || desktopService.closing) return;
        addonUpgradeResult = r;
        appendServiceLog('AutoUpgrade', r.message);
      })
      .catch(e => { notice = e.message; });
    startStatusTimer();
    startUpdateChecks();
    app.on('activate', createWindow);
  });
  app.on('window-all-closed', () => {});
  app.on('before-quit', event => {
    if (quitCompleted) return;
    event.preventDefault();
    if (installing) return;
    quitPending ||= prepareShutdown().then(() => {
      quitCompleted = true; closePermissionWindow(); closeUpdateNotice(); tray?.destroy(); app.quit();
    }).catch(error => {
      quitPending = undefined; notice = error.message;
      dialog.showErrorBox('字浮未能完成退出', error.message);
    });
  });
}
}
