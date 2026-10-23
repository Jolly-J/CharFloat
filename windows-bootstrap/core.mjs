// 用户级解压引擎；不启动应用、不修改真实客户端配置。宿主操作通过适配器注入。
import fs from 'node:fs';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import { pipeline } from 'node:stream/promises';
import { Transform } from 'node:stream';
import yauzl from 'yauzl';

export const PRODUCT = 'charfloat-user-bootstrap-v1';
export const MIN_STAGE_MS = 5000;
export const EXE = '字浮 CharFloat.exe';
const MARKER = '.charfloat-install.json';
const wait = ms => new Promise(resolve => setTimeout(resolve, ms));

export async function minimumStage(start, { now = Date.now, sleep = wait, minimum = MIN_STAGE_MS } = {}) {
  await sleep(Math.max(0, minimum - (now() - start)));
}
export function safeName(name) {
  if (typeof name !== 'string' || !name || name.includes('\\') || name.startsWith('/') || name.includes(':') || /[\x00-\x1f]/.test(name)) throw new Error('压缩包包含不安全路径');
  const parts = name.replace(/\/$/, '').split('/');
  if (parts.some(p => !p || p === '.' || p === '..' || /[. ]$/.test(p) || /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(p))) throw new Error('压缩包包含不安全文件名');
  return parts.join('/');
}
export async function fileHash(file) {
  const hash = createHash('sha256');
  for await (const chunk of fs.createReadStream(file)) hash.update(chunk);
  return hash.digest('hex');
}
export function validateManifest(m) {
  if (m?.product !== PRODUCT || !/^\d+\.\d+\.\d+$/.test(m.version) || !/^[a-f0-9]{64}$/.test(m.payloadSha256) || !Array.isArray(m.files) || !m.files.length || m.files.length > 15000) throw new Error('安装清单无效');
  const seen = new Set(); let total = 0;
  for (const f of m.files) {
    const name = safeName(f.name), lower = name.toLowerCase();
    if (seen.has(lower) || name === MARKER || !Number.isSafeInteger(f.size) || f.size < 0 || !/^[a-f0-9]{64}$/.test(f.sha256)) throw new Error('安装文件清单无效');
    seen.add(lower); total += f.size;
  }
  if (!seen.has(EXE.toLowerCase()) || !seen.has('resources/app.asar') || total > 4 * 1024 ** 3 || total !== m.totalBytes) throw new Error('安装清单缺少应用或超出限制');
  return m;
}
export async function regular(file) {
  const stat = await fs.promises.lstat(file);
  if (!stat.isFile() || stat.isSymbolicLink()) throw new Error('目标文件类型不安全，已停止以保留原文件');
  return stat;
}
async function directory(dir) {
  try { await fs.promises.mkdir(dir); } catch (e) { if (e.code !== 'EEXIST') throw e; }
  const stat = await fs.promises.lstat(dir);
  if (!stat.isDirectory() || stat.isSymbolicLink()) throw new Error('安装目录包含链接或非目录对象，已停止');
}
async function readOwned(dir, m) {
  await regular(path.join(dir, MARKER));
  const receipt = JSON.parse(await fs.promises.readFile(path.join(dir, MARKER), 'utf8'));
  if (receipt.product !== PRODUCT || (m && receipt.payloadSha256 !== m.payloadSha256)) throw new Error('目录不是本引导创建的安装，已保留原文件');
}
export async function verifyInstallation(dir, m) {
  await readOwned(dir, m);
  for (const f of m.files) {
    const parts = f.name.split('/'); let current = dir;
    for (const part of parts.slice(0, -1)) { current = path.join(current, part); const st = await fs.promises.lstat(current); if (!st.isDirectory() || st.isSymbolicLink()) throw new Error('现有安装包含链接，未覆盖'); }
    const file = path.join(dir, ...parts), st = await regular(file);
    if (st.size !== f.size || await fileHash(file) !== f.sha256) throw new Error('现有安装文件已改变或不完整，未覆盖；请保留文件并使用手动解压版');
  }
}
export async function openZip(file) {
  return new Promise((resolve, reject) => yauzl.open(file, { lazyEntries: true, autoClose: true, validateEntrySizes: true }, (e, zip) => e ? reject(e) : resolve(zip)));
}
export async function extractPayload(file, destination, manifest, onProgress = () => {}) {
  const zip = await openZip(file), expected = new Map(manifest.files.map(f => [f.name, f]));
  const seen = new Set(); let bytes = 0, entries = 0;
  await new Promise((resolve, reject) => {
    let ended = false;
    const fail = e => { if (!ended) { ended = true; zip.close(); reject(e); } };
    zip.on('error', fail);
    zip.on('end', () => { if (!ended) { ended = true; seen.size === expected.size && bytes === manifest.totalBytes ? resolve() : reject(new Error('压缩包文件不完整')); } });
    zip.on('entry', entry => { (async () => {
      if (++entries > 20000) throw new Error('压缩包条目超出限制');
      const name = safeName(entry.fileName), mode = (entry.externalFileAttributes >>> 16) & 0xf000;
      if (mode && mode !== 0x8000 && mode !== 0x4000) throw new Error('压缩包包含链接或特殊文件');
      const target = path.join(destination, ...name.split('/'));
      if (entry.fileName.endsWith('/')) { if (entry.uncompressedSize) throw new Error('目录条目无效'); await fs.promises.mkdir(target, { recursive: true }); zip.readEntry(); return; }
      const f = expected.get(name);
      if (!f || seen.has(name) || f.size !== entry.uncompressedSize || (entry.generalPurposeBitFlag & 1)) throw new Error('压缩包条目与清单不符');
      seen.add(name); await fs.promises.mkdir(path.dirname(target), { recursive: true });
      const input = await new Promise((res, rej) => zip.openReadStream(entry, (e, s) => e ? rej(e) : res(s)));
      const hash = createHash('sha256'); let size = 0;
      await pipeline(input, new Transform({ transform(chunk, _encoding, callback) {
        size += chunk.length; hash.update(chunk); callback(null, chunk);
      } }), fs.createWriteStream(target, { flags: 'wx' }));
      if (size !== f.size || hash.digest('hex') !== f.sha256) throw new Error('解压文件校验失败');
      bytes += size; onProgress({ phase: 'extract', bytes, totalBytes: manifest.totalBytes });
      zip.readEntry();
    })().catch(fail); });
    zip.readEntry();
  });
}
async function acquireLock(root) {
  const file = path.join(root, '.setup-lock'), nonce = randomUUID();
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const handle = await fs.promises.open(file, 'wx');
      try { await handle.writeFile(JSON.stringify({ product: PRODUCT, pid: process.pid, nonce })); } finally { await handle.close(); }
      return async () => { try { await regular(file); const value = JSON.parse(await fs.promises.readFile(file, 'utf8')); if (value.nonce === nonce) await fs.promises.unlink(file); } catch {} };
    } catch (e) {
      if (e.code !== 'EEXIST' || attempt) throw new Error('另一个配置正在进行，或锁文件需要检查；请关闭其他引导后重试');
      await regular(file); const old = JSON.parse(await fs.promises.readFile(file, 'utf8'));
      if (old.product !== PRODUCT || !Number.isInteger(old.pid) || old.pid <= 0) throw new Error('配置锁无法确认，未覆盖');
      try { process.kill(old.pid, 0); throw new Error('另一个配置正在进行'); } catch (alive) { if (alive.code !== 'ESRCH') throw alive; }
      // 只移除已确认进程不存在的本产品锁；不结束任何进程。
      await fs.promises.unlink(file);
    }
  }
}
export async function install({ localAppData, payload, manifest, shortcut, onProgress = () => {} }) {
  const m = validateManifest(manifest);
  if (!localAppData || !path.isAbsolute(localAppData)) throw new Error('无法确认当前用户的本地应用目录');
  const base = await fs.promises.realpath(localAppData);
  const programs = path.join(base, 'Programs'); await directory(programs);
  const root = path.join(programs, 'CharFloat');
  let created = false;
  try { await fs.promises.mkdir(root); created = true; } catch (e) { if (e.code !== 'EEXIST') throw e; }
  await directory(root);
  if (created) await fs.promises.writeFile(path.join(root, MARKER), JSON.stringify({ product: PRODUCT }), { flag: 'wx' });
  await readOwned(root);
  const release = await acquireLock(root);
  const target = path.join(root, `${m.version}-${m.payloadSha256.slice(0, 12)}`);
  let staging;
  try {
    onProgress({ phase: 'verify', bytes: 0, totalBytes: m.totalBytes });
    await regular(payload);
    if (await fileHash(payload) !== m.payloadSha256) throw new Error('安装包校验失败，请重新下载');
    try {
      const stat = await fs.promises.lstat(target);
      if (!stat.isDirectory() || stat.isSymbolicLink()) throw new Error('目标安装目录不安全');
      await verifyInstallation(target, m);
    } catch (e) {
      if (e.code !== 'ENOENT') throw e;
      // ENOENT 只能表示目标目录不存在；已有目录的缺失文件不允许覆盖。
      if (fs.existsSync(target)) throw new Error('现有安装不完整，未覆盖原文件');
      staging = await fs.promises.mkdtemp(path.join(root, '.pending-'));
      await extractPayload(payload, staging, m, onProgress);
      await fs.promises.writeFile(path.join(staging, MARKER), JSON.stringify({ product: PRODUCT, payloadSha256: m.payloadSha256, version: m.version, sourceCommit: m.sourceCommit }), { flag: 'wx' });
      await verifyInstallation(staging, m);
      await fs.promises.rename(staging, target); staging = undefined;
    }
    onProgress({ phase: 'shortcut', bytes: m.totalBytes, totalBytes: m.totalBytes });
    const executable = path.join(target, EXE);
    await shortcut({ executable, root, directory: target });
    return { executable, directory: target, payloadSha256: m.payloadSha256 };
  } finally {
    // 仅清理本次创建的临时目录；已提交安装及旧版本永不自动删除。
    try { if (staging) await fs.promises.rm(staging, { recursive: true, force: true }); }
    finally { await release(); }
  }
}
