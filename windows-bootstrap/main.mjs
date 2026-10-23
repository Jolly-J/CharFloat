import { app, BrowserWindow, ipcMain, shell, dialog } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import { once } from 'node:events';
import { randomUUID } from 'node:crypto';
import { install } from './core.mjs';
import { createShortcut } from './shortcut.mjs';
import { runStages } from './stages.mjs';
import { launchApplication } from './launch.mjs';
import { createProgressSender } from './progress.mjs';
import { createDiagnostics } from './diagnostics.mjs';

let window, busy = false, ready, launched = false, started = false;
const paintWaiters = new Map();
const resources = process.resourcesPath;
const diagnostics = createDiagnostics();
let lastProgress = {}, lastProgressAt = Date.now(), lastRendererAt = Date.now();
diagnostics.record('main-start', { electron: process.versions.electron, node: process.versions.node, arch: process.arch, platform: process.platform });
if (diagnostics.enabled) {
  process.on('uncaughtExceptionMonitor', error => { diagnostics.record('uncaught-exception', { message: error?.message, code: error?.code }); diagnostics.flush(); });
  process.on('unhandledRejection', error => { diagnostics.record('unhandled-rejection', { message: error?.message || String(error) }); diagnostics.flush(); app.exit(1); });
  process.on('exit', exitCode => diagnostics.close({ exitCode }));
  setInterval(() => diagnostics.record('main-heartbeat', { ...lastProgress, idleMs: Date.now()-lastProgressAt, rendererSilentMs: Date.now()-lastRendererAt }), 5000).unref();
}
app.setName('CharFloat Setup');
let profileError = false;
try {
  const profile = path.join(app.getPath('appData'), 'CharFloat Setup');
  fs.mkdirSync(profile, { recursive: true }); app.setPath('userData', profile);
} catch (error) { profileError = true; diagnostics.record('profile-failed', { message: error?.message, code: error?.code }); diagnostics.flush(); }
if (profileError) app.whenReady().then(async () => { await dialog.showMessageBox({ type: 'error', message: '无法写入当前用户配置目录，请使用手动解压版或联系企业管理员。' }); app.exit(1); });
else if (!app.requestSingleInstanceLock()) app.quit();
else {
  app.on('second-instance', () => { if (window) { if (window.isMinimized()) window.restore(); window.show(); window.focus(); } });
  app.whenReady().then(async () => {
    if (process.platform !== 'win32') { await dialog.showMessageBox({ message: '此引导仅适用于 Windows x64。' }); app.quit(); return; }
    window = new BrowserWindow({ width: 720, height: 550, resizable: false, maximizable: false, frame: false, transparent: true,
      show: false, roundedCorners: true, title: '字浮 CharFloat', webPreferences: { preload: path.join(__dirname, 'preload.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true, backgroundThrottling: false } });
    window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
    window.webContents.on('will-navigate', e => e.preventDefault());
    window.webContents.on('did-finish-load', () => diagnostics.record('renderer-loaded'));
    window.webContents.on('did-fail-load', (_e, code, message) => diagnostics.record('renderer-load-failed', { code, message }));
    window.webContents.on('render-process-gone', (_e, details) => { diagnostics.record('renderer-gone', details); diagnostics.flush(); });
    window.webContents.on('unresponsive', () => diagnostics.record('renderer-unresponsive'));
    window.webContents.on('responsive', () => diagnostics.record('renderer-responsive'));
    window.webContents.on('console-message', (_e, level, message) => { if (level >= 2) diagnostics.record('renderer-console', { code: level, message }); });
    window.once('ready-to-show', () => window.show());
    window.on('close', e => { diagnostics.record('window-close', { enabled: !busy }); if (busy) e.preventDefault(); });
    await window.loadFile(path.join(__dirname, 'index.html'));
  }).catch(async error => { diagnostics.record('window-start-failed', { message: error?.message, code: error?.code }); diagnostics.flush(); await dialog.showMessageBox({ type: 'error', message: '配置引导无法打开，请使用手动解压版。' }); app.exit(1); });
}
function send(value) { if (window && !window.isDestroyed()) window.webContents.send('setup:state', value); }
function ownEvent(e) { if (!window || e.sender !== window.webContents || e.senderFrame !== window.webContents.mainFrame) throw new Error('无效请求'); }
async function sendStage(state) {
  diagnostics.record('stage-sent', state);
  const token = randomUUID();
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => { paintWaiters.delete(token); reject(new Error('引导界面没有完成显示，请关闭后重试')); }, 20000);
    paintWaiters.set(token, () => { clearTimeout(timer); paintWaiters.delete(token); diagnostics.record('stage-painted', state); resolve(); });
    send({ ...state, paintToken: token, diagnosticsEnabled: diagnostics.enabled });
  });
}
ipcMain.handle('setup:painted', (e, token) => { ownEvent(e); if (typeof token === 'string') paintWaiters.get(token)?.(); });
ipcMain.handle('setup:telemetry', (e, event, data) => {
  ownEvent(e); if (!diagnostics.enabled || !['ready','heartbeat','window-error','window-rejection'].includes(event)) return;
  lastRendererAt = Date.now(); diagnostics.record(`renderer-${event}`, data);
});
ipcMain.handle('setup:start', async e => {
  ownEvent(e); if (started) return; started = true; busy = true;
  diagnostics.record('configuration-start');
  try {
    if (!window.isVisible()) await once(window, 'show');
    ready = await runStages(async () => {
      const manifest = JSON.parse(await fs.promises.readFile(path.join(resources, 'payload-manifest.json'), 'utf8'));
      const progress = createProgressSender(value => { lastProgress = value; lastProgressAt = Date.now(); diagnostics.record('progress', value); send({ stage: 'work', ...value }); });
      try {
        const result = await install({ localAppData: process.env.LOCALAPPDATA, payload: path.join(resources, 'payload.zip'), manifest,
          shortcut: args => createShortcut({ ...args, desktop: app.getPath('desktop'), shell }), onProgress: progress.push });
        return { manifest, ...result };
      } finally { progress.dispose(); }
    }, sendStage);
    busy = false;
    diagnostics.record('configuration-complete'); diagnostics.flush();
  } catch (error) {
    busy = false; ready = undefined;
    diagnostics.record('configuration-failed', { ...lastProgress, message: error?.message, code: error?.code }); diagnostics.flush();
    send({ stage: 'error', message: `配置没有完成。${error?.message || '请使用手动解压版或联系支持。'}` });
  }
});
ipcMain.handle('setup:launch', async e => {
  ownEvent(e); if (!ready || busy || launched) return { ok: false }; busy = true;
  diagnostics.record('launch-click');
  try {
    await launchApplication(ready);
    launched = true; busy = false; diagnostics.record('application-spawned'); setTimeout(() => app.quit(), 150); return { ok: true };
  } catch (error) { busy = false; diagnostics.record('launch-failed', { message: error?.message, code: error?.code }); diagnostics.flush(); return { ok: false, message: '无法启动应用，文件已保留。请从桌面快捷方式重试。' }; }
});
ipcMain.handle('setup:close', e => { ownEvent(e); diagnostics.record('close-click', { enabled: !busy }); if (!busy) app.quit(); });
app.on('window-all-closed', () => { if (!busy) app.quit(); });
