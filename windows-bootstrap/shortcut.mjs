import fs from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { regular, PRODUCT, EXE } from './core.mjs';

export async function createShortcut({ desktop, executable, root, directory, shell }) {
  const final = path.join(desktop, '字浮 CharFloat.lnk');
  let previous;
  try {
    const stat = await regular(final); if (stat.size > 1024 * 1024) throw new Error('已有快捷方式大小异常，未覆盖');
    const old = shell.readShortcutLink(final), relative = path.relative(root, old.target);
    if (!relative || relative.startsWith('..') || path.isAbsolute(relative) || path.basename(old.target) !== EXE || path.dirname(path.dirname(old.target)) !== root) throw new Error('桌面已有同名快捷方式，未覆盖；请先另行保存该快捷方式后重试');
    const oldDir = path.dirname(old.target), dirStat = await fs.promises.lstat(oldDir);
    if (!dirStat.isDirectory() || dirStat.isSymbolicLink()) throw new Error('旧快捷方式目标无法确认，未覆盖');
    const receiptFile = path.join(oldDir, '.charfloat-install.json'); await regular(receiptFile);
    const receipt = JSON.parse(await fs.promises.readFile(receiptFile, 'utf8'));
    if (receipt.product !== PRODUCT) throw new Error('旧快捷方式不属于本引导，未覆盖');
    previous = await fs.promises.readFile(final);
  } catch (e) { if (e.code !== 'ENOENT') throw e; if (fs.existsSync(final)) throw new Error('旧快捷方式目标不完整，未覆盖'); }
  const temporary = path.join(desktop, `.CharFloat-${randomUUID()}.lnk`);
  try {
    if (!shell.writeShortcutLink(temporary, 'create', { target: executable, cwd: directory, icon: executable, iconIndex: 0, description: '字浮 CharFloat', appUserModelId: 'com.antigravity.charfloat' })) throw new Error('桌面快捷方式创建失败；应用文件已保留，可重试');
    if (path.resolve(shell.readShortcutLink(temporary).target).toLowerCase() !== path.resolve(executable).toLowerCase()) throw new Error('桌面快捷方式读回失败');
    if (previous) {
      const backups = path.join(root, 'shortcut-backups');
      try { await fs.promises.mkdir(backups); } catch (e) { if (e.code !== 'EEXIST') throw e; }
      const st = await fs.promises.lstat(backups); if (!st.isDirectory() || st.isSymbolicLink()) throw new Error('快捷方式备份目录不安全');
      await fs.promises.writeFile(path.join(backups, `${Date.now()}-${randomUUID()}.lnk`), previous, { flag: 'wx' });
      await regular(final);
      if (!(await fs.promises.readFile(final)).equals(previous)) throw new Error('桌面快捷方式已被其他操作改变，未覆盖');
      await fs.promises.rename(temporary, final);
    } else {
      // 排他复制兼容重定向桌面；预检后出现同名文件时不覆盖。
      await fs.promises.copyFile(temporary, final, fs.constants.COPYFILE_EXCL);
    }
  } finally { await fs.promises.rm(temporary, { force: true }); }
}
