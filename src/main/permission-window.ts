/**
 * 完全磁盘访问权限的**置顶引导浮窗**（方案 B）。
 *
 * 为什么需要独立窗口而不是应用内弹窗：引导的第 2 步是"把应用图标拖进设置面板的列表"，
 * 用户必须**同时看到系统设置面板和这个图标**。应用内弹窗会被系统设置窗口盖住，拖拽无从下手；
 * 置顶浮窗才能停在设置面板上方。
 *
 * 复用同一个渲染包（`index.html#permission-guide`），不新增构建入口。
 */
import { BrowserWindow, app, nativeImage, nativeTheme } from 'electron';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { appToAuthorize } from './permissions.js';
import { appendServiceLog, resourcePath } from '../bridge/runtime.js';

let guideWindow: BrowserWindow | null = null;
let pendingIssue: any = null;

// 必须用 fileURLToPath：new URL(...).pathname 不解码 %20，
// 仓库路径带空格时 preload 路径会失效，window.api 直接不存在（曾导致浮窗所有按钮点了没反应）。
const here = path.dirname(fileURLToPath(import.meta.url));

export function getPermissionIssue() { return pendingIssue; }

/** 打开（或聚焦）引导浮窗，并把本次失败信息交给它渲染。 */
export function openPermissionWindow(issue: any): void {
  pendingIssue = issue;
  if (guideWindow && !guideWindow.isDestroyed()) {
    guideWindow.webContents.send('permission:issue', issue);
    guideWindow.show();
    guideWindow.focus();
    return;
  }
  // 窗口选项保持**保守**：不用 transparent、不用自定义 titleBarStyle。
  // 曾用 `transparent: true` + `titleBarStyle: 'hiddenInset'`，结果新开的渲染进程连不上主进程的
  // Mach rendezvous 服务（bootstrap_look_up … Permission denied 1100），整个应用 SIGTRAP 退出。
  guideWindow = new BrowserWindow({
    width: 430, height: 560, resizable: false, minimizable: false, maximizable: false,
    fullscreenable: false, alwaysOnTop: true,
    title: '需要完全磁盘访问权限',
    backgroundColor: nativeTheme.shouldUseDarkColors ? '#17191c' : '#ffffff',
    webPreferences: { preload: path.join(here, '../preload/index.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true }
  });
  guideWindow.setMenuBarVisibility(false);
  guideWindow.setAlwaysOnTop(true, 'floating');
  if (process.env.VITE_DEV_SERVER_URL) void guideWindow.loadURL(`${process.env.VITE_DEV_SERVER_URL}#permission-guide`);
  else void guideWindow.loadFile(path.join(here, '../renderer/index.html'), { hash: 'permission-guide' });
  guideWindow.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  guideWindow.webContents.on('will-navigate', event => event.preventDefault());
  guideWindow.webContents.once('did-finish-load', () => guideWindow?.webContents.send('permission:issue', pendingIssue));
  // 预加载失败时窗口会变成"点哪都没反应"，必须写出日志便于定位
  guideWindow.webContents.on('preload-error', (_e, preloadPath, error) => {
    appendServiceLog('PermissionWindow', `预加载脚本失败: ${preloadPath} — ${error?.message || error}`);
  });
  guideWindow.on('closed', () => { guideWindow = null; });
}

export function closePermissionWindow(): void {
  if (guideWindow && !guideWindow.isDestroyed()) guideWindow.close();
  guideWindow = null;
}

/**
 * 应用图标（用于浮窗展示与拖拽预览）。
 *
 * **刻意不用 `app.getFileIcon()`**：它在 Electron 里跑在 `ThreadPoolForegroundWorker` 线程池上，
 * 而所有"点重新部署就闪退"的崩溃报告都恰好落在该线程（EXC_BREAKPOINT）。
 * 这里改用随包资源 `resources/icon.png`，纯 Node 读文件转 data URL，不触发任何原生异步调用；
 * 拖拽预览用 `nativeImage.createFromPath`（同步，主窗口图标已在用同一 API，属已验证路径）。
 */
const ICON_RELATIVE = 'resources/icon.png';

export function loadAppIcon(): string {
  try {
    const iconPath = resourcePath(ICON_RELATIVE);
    if (!fs.existsSync(iconPath)) return '';
    return `data:image/png;base64,${fs.readFileSync(iconPath).toString('base64')}`;
  } catch (error: any) {
    appendServiceLog('PermissionWindow', `读取图标失败: ${error?.message || error}`);
    return '';
  }
}

/**
 * 发起拖拽：把当前该授权的 App 包拖出去。
 *
 * 两个关键约束（写错就会"拖不动"）：
 * 1. 必须由渲染层的 **`dragstart`** 事件触发，不能用 `mousedown`——`startDrag` 要在拖拽会话内调用；
 * 2. 必须用 `ipcRenderer.send` 而非 `invoke`——`invoke` 的异步往返会让 drag 会话先结束，等于没拖。
 */
export function startAppDrag(sender: Electron.WebContents): void {
  const file = appToAuthorize();
  const payload: any = { file };
  try {
    const icon = nativeImage.createFromPath(resourcePath(ICON_RELATIVE));
    if (!icon.isEmpty()) payload.icon = icon.resize({ width: 128, height: 128 });
  } catch {}
  sender.startDrag(payload);
}

/** 当前运行模式是否为打包版（影响引导里显示哪一个 App）。 */
export function isPackagedRun(): boolean { return app.isPackaged; }
