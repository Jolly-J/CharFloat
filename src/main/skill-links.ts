import fs from 'node:fs';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import { atomicWrite, resourcePath, runtimeHome, VERSION } from '../bridge/runtime.js';

const PRODUCT = 'charfloat-managed-skills-v1';
const MARKER = '.charfloat-skills.json';
const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
const safeName = (name: string) => /^[a-z0-9][a-z0-9-]*$/.test(name);
export const managedSkillsDirectory = () => path.join(runtimeHome(), 'managed-skills');
function writeBytes(file: string, bytes: Buffer, mode = 0o600) {
  const temp = `${file}.${randomUUID()}.tmp`;
  try { fs.writeFileSync(temp, bytes, { flag: 'wx', mode }); fs.renameSync(temp, file); }
  finally { try { fs.unlinkSync(temp); } catch (e: any) { if (e.code !== 'ENOENT') throw e; } }
}
function directory(dir: string) {
  fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
  const st = fs.lstatSync(dir);
  if (!st.isDirectory() || st.isSymbolicLink()) throw new Error('技能目录不是普通目录，已停止。');
}
function walk(root: string, prefix = ''): string[] {
  return fs.readdirSync(path.join(root, prefix), { withFileTypes: true }).flatMap(entry => {
    const relative = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isSymbolicLink() || (!entry.isDirectory() && !entry.isFile())) throw new Error('技能资源包含链接或特殊文件，未同步。');
    return entry.isDirectory() ? walk(root, relative) : [relative];
  });
}
function marker(root: string): { product: string; version: string; names: string[]; files: Record<string, string> } {
  const file = path.join(root, MARKER), st = fs.lstatSync(file);
  if (!st.isFile() || st.isSymbolicLink() || st.size > 1024 ** 2) throw new Error('技能源归属记录无效。');
  const value = JSON.parse(fs.readFileSync(file, 'utf8'));
  if (value.product !== PRODUCT || !Array.isArray(value.names) || !value.names.every(safeName) || !value.files || typeof value.files !== 'object') throw new Error('技能源不属于字浮，未覆盖。');
  return value;
}
/** One stable, app-owned source. Client links never point at a versioned installation directory. */
export function syncManagedSkills(source = resourcePath('skills'), root = managedSkillsDirectory()) {
  const existed = fs.existsSync(root);
  if (existed) { directory(root); marker(root); }
  else { directory(root); fs.writeFileSync(path.join(root, MARKER), JSON.stringify({ product: PRODUCT, version: VERSION, names: [], files: {} }), { flag: 'wx', mode: 0o600 }); }
  const lock = path.join(root, '.sync-lock');
  let handle: number;
  try { handle = fs.openSync(lock, 'wx', 0o600); } catch { throw new Error('技能正在同步或上次同步未结束，请稍后检查。'); }
  try {
    const previous = existed ? marker(root) : { files: {} as Record<string, string> };
    const names = fs.readdirSync(source).filter(name => safeName(name) && fs.existsSync(path.join(source, name, 'SKILL.md')));
    if (!names.length) throw new Error('应用没有可同步的技能资源。');
    const files: Record<string, string> = {};
    const writes: { target: string; bytes: Buffer; before?: Buffer; mode: number }[] = [];
    for (const name of names) {
      const sourceDir = path.join(source, name);
      if (fs.lstatSync(sourceDir).isSymbolicLink()) throw new Error('技能源目录包含链接。');
      directory(path.join(root, name));
      for (const relative of walk(sourceDir)) {
        const key = `${name}/${relative}`, original = path.join(source, key), target = path.join(root, key);
        let parent = path.join(root, name);
        for (const part of relative.split('/').slice(0, -1)) { parent = path.join(parent, part); directory(parent); }
        const bytes = fs.readFileSync(original); files[key] = hash(bytes);
        let before: Buffer | undefined;
        if (fs.existsSync(target)) {
          const st = fs.lstatSync(target);
          if (!st.isFile() || st.isSymbolicLink() || st.nlink !== 1) throw new Error('现有技能文件类型异常，未覆盖。');
          before = fs.readFileSync(target); const currentHash = hash(before);
          if (currentHash === files[key]) continue;
          if (!previous.files[key] || previous.files[key] !== currentHash) throw new Error('技能源含有自定义修改，已保留；请先另存修改后重新连接。');
        }
        writes.push({ target, bytes, before, mode: fs.statSync(original).mode & 0o111 ? 0o700 : 0o600 });
      }
    }
    const written: typeof writes = [];
    try {
      for (const item of writes) { writeBytes(item.target, item.bytes, item.mode); written.push(item); }
      atomicWrite(path.join(root, MARKER), JSON.stringify({ product: PRODUCT, version: VERSION, names, files }));
    } catch (error) {
      for (const item of written.reverse()) {
        if (hash(fs.readFileSync(item.target)) !== hash(item.bytes)) throw new Error('技能源被同时修改，原内容已保留，请检查后重试。');
        if (item.before) writeBytes(item.target, item.before, item.mode); else fs.unlinkSync(item.target);
      }
      throw error;
    }
    return { root: fs.realpathSync(root), names };
  } finally { fs.closeSync(handle); fs.unlinkSync(lock); }
}
/** Startup refreshes only an already-managed source; it never injects into new clients. */
export function refreshManagedSkills() {
  if (fs.existsSync(path.join(managedSkillsDirectory(), MARKER))) syncManagedSkills();
}

export type SkillLinkInspection = { status: 'linked' | 'copied' | 'missing' | 'conflict' | 'unsupported'; count: number; detail: string };
export function inspectSkillLinks(destination?: string, root = managedSkillsDirectory()): SkillLinkInspection {
  if (!destination) return { status: 'unsupported', count: 0, detail: '在客户端管理技能' };
  try {
    const source = marker(root), realRoot = fs.realpathSync(root); let count = 0, copied = false;
    if (!source.names.length) return { status: 'missing', count: 0, detail: '尚未准备技能' };
    for (const name of source.names) {
      const target = path.join(destination, name);
      let st: fs.Stats; try { st = fs.lstatSync(target); } catch (e: any) { if (e.code === 'ENOENT') return { status: 'missing', count, detail: '尚未链接技能' }; throw e; }
      if (!st.isSymbolicLink()) { if (st.isDirectory()) { copied = true; continue; } throw new Error(); }
      if (fs.realpathSync(target) !== path.join(realRoot, name) || !fs.statSync(path.join(target, 'SKILL.md')).isFile()) throw new Error();
      count++;
    }
    return copied ? { status: 'copied', count, detail: '检测到旧技能副本' } : { status: 'linked', count, detail: `${count} 项技能已链接` };
  } catch { return { status: fs.existsSync(destination) ? 'conflict' : 'missing', count: 0, detail: '技能链接尚未验证' }; }
}
export function linkClientSkills(client: string, destination: string, source: { root: string; names: string[] }, platform = process.platform) {
  if (!safeName(client)) throw new Error('客户端标识无效。');
  fs.mkdirSync(destination, { recursive: true, mode: 0o700 });
  const realDestination = fs.realpathSync(destination);
  const relative = path.relative(source.root, realDestination);
  if (!relative || (!relative.startsWith('..' + path.sep) && relative !== '..' && !path.isAbsolute(relative))) throw new Error('客户端技能目录不能位于管理源内。');
  let backups = 0;
  for (const name of source.names) {
    const origin = path.join(source.root, name), target = path.join(realDestination, name);
    let existing: fs.Stats | undefined;
    try { existing = fs.lstatSync(target); } catch (e: any) { if (e.code !== 'ENOENT') throw e; }
    if (existing?.isSymbolicLink()) { try { if (fs.realpathSync(target) === origin) continue; } catch {} }
    const temp = path.join(realDestination, `.charfloat-link-${randomUUID()}`);
    let backup: string | undefined;
    let installed = false;
    try {
      fs.symlinkSync(origin, temp, platform === 'win32' ? 'junction' : 'dir');
      if (fs.realpathSync(temp) !== origin) throw new Error('技能链接读回失败。');
      if (existing) {
        // Preserve a user's old copy or link outside the directory scanned for skills.
        const backupBase = path.join(path.dirname(realDestination), '.charfloat-skill-backups'); directory(backupBase);
        const backupRoot = path.join(backupBase, `${Date.now()}-${client}-${randomUUID()}`);
        directory(backupRoot); backup = path.join(backupRoot, name);
        fs.renameSync(target, backup); backups++;
      }
      fs.renameSync(temp, target);
      installed = true;
      if (fs.realpathSync(target) !== origin || !fs.statSync(path.join(target, 'SKILL.md')).isFile()) throw new Error('技能链接未通过读回。');
    } catch (error) {
      // Roll back only the exact link created by this operation, never a concurrent edit.
      if (installed) {
        const st = fs.lstatSync(target);
        if (!st.isSymbolicLink() || fs.realpathSync(target) !== origin) throw new Error('技能目标被同时修改，原副本仍在备份目录，请检查后重试。');
        fs.unlinkSync(target);
      }
      if (backup) fs.renameSync(backup, target);
      throw error;
    } finally { try { fs.unlinkSync(temp); } catch (e: any) { if (e.code !== 'ENOENT') throw e; } }
  }
  const verified = inspectSkillLinks(destination, source.root);
  if (verified.status !== 'linked') throw new Error('技能链接未全部通过验证，原副本已保留。');
  return { count: verified.count, backups };
}
