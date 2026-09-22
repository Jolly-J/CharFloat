/**
 * macOS 完全磁盘访问权限（Full Disk Access, FDA）的识别与引导。
 *
 * 背景：WPS 与 Microsoft Office 的加载项目录位于**别的 App 的沙盒容器**内
 * （`~/Library/Containers/com.kingsoft.wpsoffice.mac/...`、`com.microsoft.Excel/...`），
 * 这类路径受 TCC「App Data / Data Vault」保护。
 *
 * 为什么不能"自动弹权限申请"：**FDA 是 TCC 里唯一既没有申请 API、也没有系统弹窗的类别**。
 * Documents/Desktop/Downloads 有 `NS*FolderUsageDescription`、相机麦克风有专门的 requestAccess，
 * 而 FDA 权限范围太大，Apple 要求用户必须在「系统设置 → 隐私与安全性 → 完全磁盘访问权限」里
 * 显式勾选。因此 App 唯一能做的是：检测到 EPERM/EACCES → 主动引导用户去设置。
 *
 * 本模块只做三件事：识别权限类错误、给出当前运行模式下该授权的 App 路径、给出设置深链。
 */
import { app } from 'electron';
import path from 'node:path';
import { detectPermissionIssueCore } from './permission-detect.js';

/** FDA 设置面板深链。属未公开但长期可用的 URL scheme，失败时界面仍给出文字步骤。 */
export const FULL_DISK_ACCESS_URL = 'x-apple.systempreferences:com.apple.preference.security?Privacy_AllFiles';

/**
 * 当前运行模式下用户应当授权的 App 路径。
 *
 * - 打包版：`.../Office Agent Bridge.app`
 * - 开发版：`.../node_modules/electron/dist/Electron.app`
 *
 * 两者是**不同的 TCC 主体**，授权一个不会让另一个生效，因此界面必须显示当前这一个。
 */
export function appToAuthorize(): string {
  const exe = process.execPath;
  const appIndex = exe.indexOf('.app/');
  if (appIndex !== -1) return exe.slice(0, appIndex + 4);
  return app.isPackaged ? path.dirname(exe) : exe;
}
