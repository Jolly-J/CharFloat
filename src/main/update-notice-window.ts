import { BrowserWindow, ipcMain, nativeTheme } from 'electron';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
let noticeWindow: BrowserWindow | null = null;
let content: { version: string; currentVersion: string; notes: string; theme: string } | undefined;

export function setupUpdateNoticeIpc(openUpdates: () => void, isClosing: () => boolean) {
  const own = (event: Electron.IpcMainInvokeEvent) => {
    if (!noticeWindow || event.sender !== noticeWindow.webContents || event.senderFrame !== noticeWindow.webContents.mainFrame) throw new Error('无效的更新提醒请求来源。');
  };
  ipcMain.handle('update-notice:info', event => { own(event); return content; });
  ipcMain.handle('update-notice:ready', event => { own(event); if (!isClosing()) noticeWindow?.show(); else closeUpdateNotice(); });
  ipcMain.handle('update-notice:dismiss', event => { own(event); closeUpdateNotice(); });
  ipcMain.handle('update-notice:open', event => {
    own(event);
    if (!isClosing()) openUpdates();
    closeUpdateNotice();
    // No download or install is reachable through the reminder's actions.
  });
}
export async function showUpdateNotice(value: { version: string; currentVersion: string; notes: string }) {
  closeUpdateNotice();
  content = { ...value, theme: nativeTheme.shouldUseDarkColors ? 'dark' : 'light' };
  const popup = new BrowserWindow({ width: 420, height: 480, show: false, frame: false, roundedCorners: true,
    resizable: false, minimizable: false, maximizable: false, fullscreenable: false, title: '字浮有新版本',
    backgroundColor: nativeTheme.shouldUseDarkColors ? '#1a1c20' : '#ffffff',
    webPreferences: { preload: path.join(here, '../preload/index.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true, backgroundThrottling: false },
  });
  noticeWindow = popup;
  popup.setMenuBarVisibility(false);
  popup.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  popup.webContents.on('will-navigate', event => event.preventDefault());
  popup.once('closed', () => { if (noticeWindow === popup) { noticeWindow = null; content = undefined; } });
  try {
    if (process.env.VITE_DEV_SERVER_URL) await popup.loadURL(`${process.env.VITE_DEV_SERVER_URL}#update-notice`);
    else await popup.loadFile(path.join(here, '../renderer/index.html'), { hash: 'update-notice' });
  } catch (error) { if (!popup.isDestroyed()) popup.destroy(); throw error; }
}
export function closeUpdateNotice() {
  if (noticeWindow && !noticeWindow.isDestroyed()) noticeWindow.close();
  noticeWindow = null; content = undefined;
}
