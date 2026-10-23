import { app, BrowserWindow, ipcMain, shell, dialog } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import { once } from 'node:events';
import { randomUUID } from 'node:crypto';
import { install } from './core.mjs';
import { createShortcut } from './shortcut.mjs';
import { runStages } from './stages.mjs';
import { launchApplication } from './launch.mjs';

let window, busy = false, ready, launched = false, started = false;
const paintWaiters = new Map();
const resources = process.resourcesPath;
app.setName('CharFloat Setup');
let profileError = false;
try {
  const profile = path.join(app.getPath('appData'), 'CharFloat Setup');
  fs.mkdirSync(profile, { recursive: true }); app.setPath('userData', profile);
} catch { profileError = true; }
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
    window.once('ready-to-show', () => window.show());
    window.on('close', e => { if (busy) e.preventDefault(); });
    await window.loadFile(path.join(__dirname, 'index.html'));
  }).catch(async () => { await dialog.showMessageBox({ type: 'error', message: '配置引导无法打开，请使用手动解压版。' }); app.exit(1); });
}
function send(value) { if (window && !window.isDestroyed()) window.webContents.send('setup:state', value); }
function ownEvent(e) { if (!window || e.sender !== window.webContents || e.senderFrame !== window.webContents.mainFrame) throw new Error('无效请求'); }
async function sendStage(state) {
  const token = randomUUID();
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => { paintWaiters.delete(token); reject(new Error('引导界面没有完成显示，请关闭后重试')); }, 20000);
    paintWaiters.set(token, () => { clearTimeout(timer); paintWaiters.delete(token); resolve(); });
    send({ ...state, paintToken: token });
  });
}
ipcMain.handle('setup:painted', (e, token) => { ownEvent(e); if (typeof token === 'string') paintWaiters.get(token)?.(); });
ipcMain.handle('setup:start', async e => {
  ownEvent(e); if (started) return; started = true; busy = true;
  try {
    if (!window.isVisible()) await once(window, 'show');
    ready = await runStages(async () => {
      const manifest = JSON.parse(await fs.promises.readFile(path.join(resources, 'payload-manifest.json'), 'utf8'));
      const result = await install({ localAppData: process.env.LOCALAPPDATA, payload: path.join(resources, 'payload.zip'), manifest,
        shortcut: args => createShortcut({ ...args, desktop: app.getPath('desktop'), shell }),
        onProgress: progress => send({ stage: 'work', ...progress }) });
      return { manifest, ...result };
    }, sendStage);
    busy = false;
  } catch (error) {
    busy = false; ready = undefined;
    send({ stage: 'error', message: `配置没有完成。${error?.message || '请使用手动解压版或联系支持。'}` });
  }
});
ipcMain.handle('setup:launch', async e => {
  ownEvent(e); if (!ready || busy || launched) return { ok: false }; busy = true;
  try {
    await launchApplication(ready);
    launched = true; busy = false; setTimeout(() => app.quit(), 150); return { ok: true };
  } catch { busy = false; return { ok: false, message: '无法启动应用，文件已保留。请从桌面快捷方式重试。' }; }
});
ipcMain.handle('setup:close', e => { ownEvent(e); if (!busy) app.quit(); });
app.on('window-all-closed', () => { if (!busy) app.quit(); });
