import nodeFs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pipeline } from 'node:stream/promises';
import { Transform } from 'node:stream';
import { randomUUID } from 'node:crypto';
import { DOMParser } from '@xmldom/xmldom';
import yauzl from 'yauzl';
import { openZipStream } from '../../windows-bootstrap/zip-stream.mjs';
import { canonicalDirectory, plainDirectory } from './update-install.js';

const fs: typeof nodeFs = process.versions.electron ? createRequire(process.execPath)('original-fs') : nodeFs;
const BUNDLE_ID = 'com.antigravity.charfloat';
const MAX_BYTES = 4 * 1024 ** 3;
function relativeName(value: string) {
  if (!value || value.includes('\\') || value.includes(':') || /[\x00-\x1f]/.test(value) || value.startsWith('/')) throw new Error('Mac 更新包路径无效。');
  const parts = value.replace(/\/$/, '').split('/');
  if (parts.some(p => !p || p === '.' || p === '..')) throw new Error('Mac 更新包包含越界路径。');
  return parts.join('/');
}
function inside(root: string, candidate: string) {
  const rel = path.relative(root, candidate);
  return rel !== '' && !rel.startsWith(`..${path.sep}`) && rel !== '..' && !path.isAbsolute(rel);
}
/** Extract regular entries first and create links last; no archive entry can write through a link. */
export async function extractMacArchive(archive: string, destination: string) {
  await canonicalDirectory(destination);
  const zip: any = await new Promise((resolve, reject) => yauzl.open(archive, { lazyEntries: true, validateEntrySizes: true }, (e, z) => e ? reject(e) : resolve(z)));
  const seen = new Set<string>(), links: { target: string; value: string }[] = [];
  let count = 0, declared = 0, written = 0, appName = '';
  await new Promise<void>((resolve, reject) => {
    let finished = false;
    const fail = (e: unknown) => { if (!finished) { finished = true; zip.close(); reject(e); } };
    zip.on('error', fail);
    zip.on('end', () => { if (!finished) { finished = true; resolve(); } });
    zip.on('entry', (entry: any) => { void (async () => {
      if (++count > 40000 || (entry.generalPurposeBitFlag & 1)) throw new Error('Mac 更新包加密或条目过多。');
      const name = relativeName(entry.fileName), top = name.split('/')[0];
      if (top === '__MACOSX') { zip.readEntry(); return; }
      if (!top.endsWith('.app') || (appName && top !== appName)) throw new Error('Mac 更新包必须仅包含一个应用。');
      appName = top;
      const lower = name.normalize('NFC').toLowerCase();
      if (seen.has(lower)) throw new Error('Mac 更新包含有重复路径。'); seen.add(lower);
      const target = path.join(destination, name), mode = (entry.externalFileAttributes >>> 16), type = mode & 0xf000;
      if (!inside(destination, target) || (type && ![0x8000, 0x4000, 0xa000].includes(type))) throw new Error('Mac 更新包含有特殊文件。');
      declared += entry.uncompressedSize;
      if (!Number.isSafeInteger(declared) || declared > MAX_BYTES) throw new Error('Mac 更新包展开大小超出限制。');
      if (entry.fileName.endsWith('/')) {
        if (entry.uncompressedSize || (type && type !== 0x4000)) throw new Error('Mac 更新目录条目无效。');
        await fs.promises.mkdir(target, { recursive: true, mode: 0o755 }); zip.readEntry(); return;
      }
      if (type === 0x4000) throw new Error('Mac 更新目录标记无效。');
      const input = await openZipStream(zip, entry);
      if (type === 0xa000) {
        if (entry.uncompressedSize > 4096) throw new Error('Mac 更新链接过长。');
        const parts: Buffer[] = []; let size = 0;
        for await (const chunk of input) { size += chunk.length; if (size > 4096) throw new Error('Mac 更新链接过长。'); parts.push(chunk); }
        const value = new TextDecoder('utf-8', { fatal: true }).decode(Buffer.concat(parts));
        if (size !== entry.uncompressedSize || !value || value.includes('\0') || value.includes('\\') || path.isAbsolute(value) || !inside(path.join(destination, appName), path.resolve(path.dirname(target), value))) throw new Error('Mac 更新链接越出应用目录。');
        links.push({ target, value });
      } else {
        await fs.promises.mkdir(path.dirname(target), { recursive: true, mode: 0o755 });
        let size = 0;
        await pipeline(input, new Transform({ transform(chunk, _enc, cb) {
          size += chunk.length; written += chunk.length;
          cb(size > entry.uncompressedSize || written > MAX_BYTES ? new Error('Mac 解压大小超出限制。') : null, chunk);
        } }), fs.createWriteStream(target, { flags: 'wx', mode: (mode & 0o111) ? 0o755 : 0o644 }));
        if (size !== entry.uncompressedSize) throw new Error('Mac 更新文件不完整。');
      }
      zip.readEntry();
    })().catch(fail); });
    zip.readEntry();
  });
  if (!appName) throw new Error('Mac 更新包为空。');
  // A path reserved for a symlink must not have been created by an entry below it.
  for (const link of links) {
    const prefix = path.relative(destination, link.target).normalize('NFC').toLowerCase() + '/';
    if ([...seen].some(name => name.startsWith(prefix))) throw new Error('Mac 更新包含链接与子路径冲突。');
  }
  for (const link of links) { await fs.promises.mkdir(path.dirname(link.target), { recursive: true }); await fs.promises.symlink(link.value, link.target); }
  const app = path.join(destination, appName), realApp = await fs.promises.realpath(app);
  for (const link of links) {
    const resolved = await fs.promises.realpath(link.target);
    if (!inside(realApp, resolved)) throw new Error('Mac 更新含有悬空、循环或越界链接。');
  }
  return app;
}
function plistStrings(text: string) {
  if (text.length > 1024 * 1024 || /<!ENTITY/i.test(text)) throw new Error('应用标识文件无效。');
  const doc = new DOMParser({ onError: () => { throw new Error('应用标识 XML 无效。'); } }).parseFromString(text, 'application/xml');
  const dict = doc.documentElement?.getElementsByTagName('dict')[0];
  if (!dict) throw new Error('应用标识缺少字典。');
  const values: Record<string, string> = {};
  for (let n = dict.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1 || n.nodeName !== 'key') continue;
    let value = n.nextSibling; while (value && value.nodeType !== 1) value = value.nextSibling;
    if (value?.nodeName === 'string') { const key = n.textContent || ''; if (Object.hasOwn(values, key)) throw new Error('应用标识含重复字段。'); values[key] = value.textContent || ''; }
  }
  return values;
}
export async function validateMacBundle(app: string, version: string, arch: string) {
  await canonicalDirectory(app);
  for (const dir of ['Contents', 'Contents/MacOS']) await plainDirectory(path.join(app, dir));
  const plistFile = path.join(app, 'Contents/Info.plist'), stat = await fs.promises.lstat(plistFile);
  if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 1024 * 1024) throw new Error('应用标识文件不安全。');
  const values = plistStrings(await fs.promises.readFile(plistFile, 'utf8'));
  if (values.CFBundleIdentifier !== BUNDLE_ID || values.CFBundleShortVersionString !== version || !values.CFBundleExecutable || /[\\/\x00-\x1f]/.test(values.CFBundleExecutable) || ['.', '..'].includes(values.CFBundleExecutable)) throw new Error('应用身份或版本与已签名清单不符。');
  const executable = path.join(app, 'Contents/MacOS', values.CFBundleExecutable), es = await fs.promises.lstat(executable);
  if (!es.isFile() || es.isSymbolicLink() || !(es.mode & 0o111)) throw new Error('应用主程序类型无效。');
  const handle = await fs.promises.open(executable, 'r');
  try {
    const header = Buffer.alloc(4096); const { bytesRead } = await handle.read(header, 0, 4096, 0);
    const expected = arch === 'arm64' ? 0x0100000c : arch === 'x64' ? 0x01000007 : 0;
    const magic = header.readUInt32BE(0); let matching = false;
    if (magic === 0xcffaedfe && bytesRead >= 8) matching = header.readUInt32LE(4) === expected;
    if ([0xcafebabe, 0xcafebabf].includes(magic)) {
      const count = header.readUInt32BE(4), stride = magic === 0xcafebabe ? 20 : 32;
      if (count > 32 || 8 + count * stride > bytesRead) throw new Error('应用架构头无效。');
      for (let i = 0; i < count; i++) if (header.readUInt32BE(8 + i * stride) === expected) matching = true;
    }
    if (!expected || !matching) throw new Error('应用主程序架构不匹配。');
  } finally { await handle.close(); }
  return { executable };
}
export type MacTransaction = { schema: 1; target: string; staged: string; backup: string; oldVersion: string; version: string; arch: string; phase: 'prepared' | 'moving' | 'committed' | 'rolled-back' };
async function save(file: string, tx: MacTransaction) {
  const temp = `${file}.${randomUUID()}.tmp`;
  await fs.promises.writeFile(temp, JSON.stringify(tx), { flag: 'wx', mode: 0o600 });
  const h = await fs.promises.open(temp, 'r+'); try { await h.sync(); } finally { await h.close(); }
  await fs.promises.rename(temp, file);
}
/** Offline recovery only restores a missing original path; it never replaces a present app or launches one. */
export async function recoverMacUpdate(journal: string) {
  const dir = path.dirname(journal); await canonicalDirectory(dir);
  const ds = await fs.promises.lstat(dir), js = await fs.promises.lstat(journal);
  if (!path.basename(dir).startsWith('.charfloat-update-') || path.basename(journal) !== 'transaction.json' || (ds.mode & 0o077) || (process.getuid && ds.uid !== process.getuid()) || !js.isFile() || js.isSymbolicLink() || js.size > 16384) throw new Error('更新恢复记录不安全。');
  const tx = JSON.parse(await fs.promises.readFile(journal, 'utf8')) as MacTransaction;
  if (tx.schema !== 1 || typeof tx.target !== 'string' || typeof tx.staged !== 'string' || tx.backup !== path.join(dir, 'previous.app') || path.dirname(tx.staged) !== dir || path.dirname(tx.target) !== path.dirname(dir) || !tx.target.endsWith('.app') || !tx.staged.endsWith('.app')) throw new Error('恢复记录路径无效。');
  await canonicalDirectory(path.dirname(dir));
  try { await fs.promises.lstat(tx.target); throw new Error('原应用路径仍存在，未覆盖；请保留应用和备份后手动检查。'); } catch (error: any) { if (error.code !== 'ENOENT') throw error; }
  await validateMacBundle(tx.backup, tx.oldVersion, tx.arch);
  await fs.promises.rename(tx.backup, tx.target); tx.phase = 'rolled-back'; await save(journal, tx);
  return tx.target;
}
/** Only renames same-volume bundles; a failed commit restores the previous bundle without deleting either. */
export async function commitMacUpdate(journal: string, tx: MacTransaction, afterMove?: () => Promise<void>, authorize?: () => Promise<void>) {
  await canonicalDirectory(path.dirname(tx.target)); await canonicalDirectory(path.dirname(tx.staged));
  if (path.dirname(tx.backup) !== path.dirname(tx.staged) || !tx.staged.endsWith('.app') || !tx.backup.endsWith('.app') || path.dirname(tx.target) !== path.dirname(path.dirname(tx.staged))) throw new Error('Mac 更新事务路径无效。');
  if (fs.existsSync(tx.backup) && !fs.existsSync(tx.target)) {
    await validateMacBundle(tx.backup, tx.oldVersion, tx.arch);
    await fs.promises.rename(tx.backup, tx.target); tx.phase = 'rolled-back'; await save(journal, tx);
    throw new Error('上次更新中断，旧应用已恢复，请重新检查更新。');
  }
  if (fs.existsSync(tx.backup)) throw new Error('已有更新备份，未覆盖；请保留事务目录。');
  await validateMacBundle(tx.target, tx.oldVersion, tx.arch);
  await validateMacBundle(tx.staged, tx.version, tx.arch);
  if ((await fs.promises.stat(tx.target)).dev !== (await fs.promises.stat(tx.staged)).dev) throw new Error('更新暂存与应用不在同一磁盘。');
  await authorize?.();
  tx.phase = 'moving'; await save(journal, tx);
  await fs.promises.rename(tx.target, tx.backup);
  try {
    await afterMove?.();
    await authorize?.();
    await fs.promises.rename(tx.staged, tx.target);
    await validateMacBundle(tx.target, tx.version, tx.arch);
    tx.phase = 'committed'; await save(journal, tx);
  } catch (error) {
    if (fs.existsSync(tx.target)) await fs.promises.rename(tx.target, tx.staged);
    await fs.promises.rename(tx.backup, tx.target);
    tx.phase = 'rolled-back'; await save(journal, tx);
    throw error;
  }
}
export async function stageMacUpdate(archive: string, target: string, oldVersion: string, version: string, arch: string) {
  if (!target.endsWith('.app') || target.includes('/AppTranslocation/') || target.startsWith('/Volumes/')) throw new Error('请先把字浮移到本机可写目录后更新；不支持磁盘映像或隔离运行路径。');
  await validateMacBundle(target, oldVersion, arch);
  const parent = path.dirname(target); await canonicalDirectory(parent); await fs.promises.access(parent, fs.constants.W_OK);
  const directory = await fs.promises.mkdtemp(path.join(parent, '.charfloat-update-'));
  await fs.promises.chmod(directory, 0o700);
  try {
    const staged = await extractMacArchive(archive, directory);
    await validateMacBundle(staged, version, arch);
    const tx: MacTransaction = { schema: 1, target, staged, backup: path.join(directory, 'previous.app'), oldVersion, version, arch, phase: 'prepared' };
    const journal = path.join(directory, 'transaction.json'); await save(journal, tx);
    return { directory, journal, tx };
  } catch (e) { await fs.promises.rm(directory, { recursive: true, force: true }); throw e; }
}
