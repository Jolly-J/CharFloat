import fs from 'node:fs';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import { readSignedManifest, compareVersion, configuredTrust, UpdateSecurityError, type UpdateTrust } from './update-protocol.js';

export type InstallTicket = { schema: 1; nonce: string; parentPid: number; createdAt: number; platform: 'darwin' | 'win32'; arch: 'arm64' | 'x64'; currentVersion: string; target: string; archive: string; envelope: string };
export async function plainFile(file: string, limit = Infinity) {
  const st = await fs.promises.lstat(file);
  if (!st.isFile() || st.isSymbolicLink() || st.nlink !== 1 || st.size > limit) throw new UpdateSecurityError('更新文件不是独立的普通文件。');
  return st;
}
export async function plainDirectory(dir: string) {
  const st = await fs.promises.lstat(dir);
  if (!st.isDirectory() || st.isSymbolicLink()) throw new UpdateSecurityError('更新目录包含链接或不是目录。');
  return st;
}
export async function canonicalDirectory(dir: string) {
  await plainDirectory(dir);
  if (path.resolve(dir) !== await fs.promises.realpath(dir)) throw new UpdateSecurityError('更新路径包含目录链接，请将应用移到普通可写目录后重试。');
}
export async function writeInstallState(directory: string, value: object) {
  if ('phase' in value && value.phase !== 'failed') await assertNotRevoked(directory);
  const file = path.join(directory, 'result.json'), temp = path.join(directory, `.result-${randomUUID()}`);
  await fs.promises.writeFile(temp, JSON.stringify(value), { flag: 'wx', mode: 0o600 });
  await fs.promises.rename(temp, file);
}
async function assertNotRevoked(directory: string) {
  try { await fs.promises.lstat(path.join(directory, 'revoked.json')); }
  catch (error: any) { if (error.code === 'ENOENT') return; throw error; }
  throw new UpdateSecurityError('本次更新授权已撤销，请重新确认更新。');
}
export async function grantInstallAuthorization(file: string, nonce: string) {
  const directory = path.dirname(file); await canonicalDirectory(directory); await assertNotRevoked(directory);
  try { await fs.promises.lstat(path.join(directory, 'consumed.json')); throw new UpdateSecurityError('本次更新授权已使用，不能重复安装。'); }
  catch (error: any) { if (error.code !== 'ENOENT') throw error; }
  await fs.promises.writeFile(path.join(directory, 'granted.json'), JSON.stringify({ schema: 1, nonce }), { flag: 'wx', mode: 0o600 });
  await assertNotRevoked(directory);
}
export async function revokeInstallAuthorization(file: string, nonce: string) {
  const directory = path.dirname(file); await canonicalDirectory(directory);
  try { await fs.promises.writeFile(path.join(directory, 'revoked.json'), JSON.stringify({ schema: 1, nonce }), { flag: 'wx', mode: 0o600 }); }
  catch (error: any) { if (error.code !== 'EEXIST') throw error; }
  // The permanent tombstone is authoritative even if a late grant write races this unlink.
  await fs.promises.unlink(path.join(directory, 'granted.json')).catch(error => { if (error.code !== 'ENOENT') throw error; });
  await writeInstallState(directory, { phase: 'failed', nonce, message: '更新交接未完成，安装授权已撤销；原应用未替换。请重新确认更新。' });
}
/** Consume exactly once after the old parent exits; the returned guard is checked at every commit boundary. */
export async function consumeInstallAuthorization(file: string, nonce: string) {
  const directory = path.dirname(file); await canonicalDirectory(directory); await assertNotRevoked(directory);
  const granted = path.join(directory, 'granted.json'), consumed = path.join(directory, 'consumed.json');
  try { await fs.promises.lstat(consumed); throw new UpdateSecurityError('本次更新授权已使用，不能重复安装。'); }
  catch (error: any) { if (error.code !== 'ENOENT') throw error; }
  try {
    await plainFile(granted, 4096);
    const value = JSON.parse(await fs.promises.readFile(granted, 'utf8'));
    if (value.schema !== 1 || value.nonce !== nonce) throw new UpdateSecurityError('安装授权回执不匹配。');
    await fs.promises.rename(granted, consumed);
  } catch (error: any) { if (error.code === 'ENOENT') throw new UpdateSecurityError('没有本次 GUI 确认的安装授权，未执行安装。'); throw error; }
  const guard = async () => {
    await assertNotRevoked(directory); await plainFile(consumed, 4096);
    const value = JSON.parse(await fs.promises.readFile(consumed, 'utf8'));
    if (value.schema !== 1 || value.nonce !== nonce) throw new UpdateSecurityError('安装授权已改变，未继续安装。');
  };
  await guard(); return guard;
}
export async function createInstallTicket(cache: string, ticket: Omit<InstallTicket, 'schema' | 'nonce' | 'createdAt'>) {
  await canonicalDirectory(cache);
  const dir = await fs.promises.mkdtemp(path.join(cache, 'install-'));
  await fs.promises.chmod(dir, 0o700);
  const value: InstallTicket = { ...ticket, schema: 1, nonce: randomUUID(), createdAt: Date.now() };
  const file = path.join(dir, 'ticket.json');
  await fs.promises.writeFile(file, JSON.stringify(value), { flag: 'wx', mode: 0o600 });
  return { file, directory: dir, ticket: value };
}
export async function readInstallTicket(file: string, trust: UpdateTrust, platform: string, arch: string) {
  if (!configuredTrust(trust)) throw new UpdateSecurityError('正式更新源或签名公钥尚未配置，不能安装。');
  const directory = path.dirname(file);
  await canonicalDirectory(directory);
  await assertNotRevoked(directory);
  const ds = await plainDirectory(directory);
  if (process.platform !== 'win32' && ((ds.mode & 0o077) || ds.uid !== process.getuid?.())) throw new UpdateSecurityError('更新交接目录权限无效。');
  await plainFile(file, 512 * 1024);
  const ticket = JSON.parse(await fs.promises.readFile(file, 'utf8')) as InstallTicket;
  if (path.basename(file) !== 'ticket.json' || ticket.schema !== 1 || !/^[a-f0-9-]{36}$/.test(ticket.nonce) || !Number.isSafeInteger(ticket.parentPid) || ticket.parentPid < 1 || ticket.parentPid === process.pid || !Number.isFinite(ticket.createdAt) || Math.abs(Date.now() - ticket.createdAt) > 3600000 || ticket.platform !== platform || ticket.arch !== arch || !path.isAbsolute(ticket.target) || !path.isAbsolute(ticket.archive) || path.resolve(ticket.target) !== ticket.target || path.resolve(ticket.archive) !== ticket.archive) throw new UpdateSecurityError('更新交接已过期或与当前系统不匹配。');
  const selected = readSignedManifest(ticket.envelope, trust, { version: ticket.currentVersion, platform, arch });
  if (compareVersion(selected.manifest.version, ticket.currentVersion) <= 0) throw new UpdateSecurityError('只允许安装更高版本。');
  await canonicalDirectory(path.dirname(ticket.archive));
  const before = await plainFile(ticket.archive);
  const handle = await fs.promises.open(ticket.archive, fs.constants.O_RDONLY | (fs.constants.O_NOFOLLOW || 0));
  try {
    const st = await handle.stat();
    if (st.ino !== before.ino || st.dev !== before.dev || st.size !== selected.asset.size) throw new UpdateSecurityError('更新包已改变。');
    const hash = createHash('sha256');
    for await (const chunk of handle.createReadStream({ autoClose: false })) hash.update(chunk);
    const after = await handle.stat();
    if (hash.digest('hex') !== selected.asset.sha256 || after.mtimeMs !== st.mtimeMs || after.size !== st.size) throw new UpdateSecurityError('安装前更新包校验失败。');
  } finally { await handle.close(); }
  return { ticket, selected, directory };
}
export async function waitForParentExit(pid: number, timeoutMs = 60000) {
  const deadline = Date.now() + timeoutMs;
  while (true) {
    try { process.kill(pid, 0); } catch (e: any) { if (e.code === 'ESRCH') return; throw new Error('无法确认旧应用已退出，未执行安装。'); }
    if (Date.now() >= deadline) throw new Error('等待旧应用退出超时，未执行安装；请从托盘退出后重试。');
    await new Promise(r => setTimeout(r, 100));
  }
}
export async function lastInstallReport(cache: string) {
  try {
    await canonicalDirectory(cache);
    const names = (await fs.promises.readdir(cache)).filter(n => /^install-[A-Za-z0-9]+$/.test(n)).slice(-128);
    const reports: { time: number; phase: string; message: string }[] = [];
    for (const name of names) {
      try {
        const dir = path.join(cache, name); await plainDirectory(dir);
        const file = path.join(dir, 'result.json'), st = await plainFile(file, 16384);
        const value = JSON.parse(await fs.promises.readFile(file, 'utf8'));
        if (['installed', 'failed', 'ready'].includes(value.phase)) reports.push({ time: st.mtimeMs, phase: value.phase, message: value.phase === 'ready' ? '上次更新未确认完成。原应用或备份已保留，请查看更新交接目录并重试。' : typeof value.message === 'string' ? value.message.slice(0, 2000) : '上次更新结束，请检查当前版本。' });
      } catch { /* never follow an unsafe or incomplete history entry */ }
    }
    return reports.sort((a, b) => b.time - a.time)[0];
  } catch { return undefined; }
}
