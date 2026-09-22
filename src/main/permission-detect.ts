/**
 * macOS 磁盘权限错误的**纯识别逻辑**（不依赖 electron，可单元测试）。
 *
 * 为什么要独立成模块：识别错误的正确性直接决定用户体验——
 * 漏判会让用户继续看到裸 EPERM；误判会把无关失败说成权限问题，误导用户去改系统设置。
 * 因此这部分必须有测试，而 electron 无法在普通 Node 测试里加载。
 */

export interface PermissionIssueCore {
  kind: 'macos-full-disk-access';
  targetPath: string;
  detail: string;
}

/** 从错误信息里提取可能的目标路径。 */
export function extractProtectedPath(message: string): string {
  const quoted = message.match(/['"]([^'"]*\/Library\/[^'"]*)['"]/);
  if (quoted) return quoted[1];
  const bare = message.match(/(\/Users\/[^\s'"]+|\/Library\/[^\s'"]+)/);
  // 裸路径会连上结尾标点（如 cp 报错里的冒号），去掉以免展示错误路径
  return bare ? bare[1].replace(/[.,:;)\]、。，；）]+$/, '') : '';
}

/** 判断错误是否属于 macOS 磁盘权限受限；非权限类错误返回 null。 */
export function detectPermissionIssueCore(
  error: unknown,
  platform: NodeJS.Platform = process.platform
): PermissionIssueCore | null {
  if (platform !== 'darwin') return null;
  const code = (error as any)?.code;
  const detail = error instanceof Error ? error.message : String(error ?? '');
  const isPermission = code === 'EPERM' || code === 'EACCES'
    || /Operation not permitted|not permitted|Permission denied|EPERM|EACCES/i.test(detail);
  if (!isPermission) return null;
  return { kind: 'macos-full-disk-access', targetPath: extractProtectedPath(detail), detail };
}
