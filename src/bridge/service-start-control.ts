import fs from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { atomicWrite, runtimeHome } from './runtime.js';

type Control = { schema: 1; generation: string; paused: boolean };
const pausedMessage = '字浮已退出并暂停后台自动启动；请打开字浮桌面端，或显式执行 CLI --start 后重新连接。';
const paths = () => ({ lock: path.join(runtimeHome(), '.service-start.lock'), control: path.join(runtimeHome(), 'service-start-control.json'), pause: path.join(runtimeHome(), 'desktop-stopped.json') });
/** Serializes start intent, actual daemon listen, and GUI pause across processes in this runtime home. */
export async function withServiceStartLock<T>(action: () => Promise<T> | T, signal?: AbortSignal) {
  fs.mkdirSync(runtimeHome(), { recursive: true, mode: 0o700 });
  const { lock } = paths(), nonce = randomUUID(), deadline = Date.now() + 20000;
  while (true) {
    signal?.throwIfAborted();
    let handle: fs.promises.FileHandle;
    try { handle = await fs.promises.open(lock, 'wx', 0o600); }
    catch (error: any) {
      if (error.code !== 'EEXIST') throw error;
      if (Date.now() >= deadline) throw new Error('后台启动协调锁超时；未启动或结束任何其他进程。请关闭相关字浮实例后检查运行目录的锁文件。');
      await new Promise(r => setTimeout(r, 25)); continue;
    }
    try {
      await handle.writeFile(JSON.stringify({ schema: 1, pid: process.pid, nonce }));
      signal?.throwIfAborted();
      return await action();
    } finally {
      await handle.close();
      // Never remove an unknown/stale owner's lock automatically: PID reuse is not ownership proof.
      const st = await fs.promises.lstat(lock);
      if (!st.isFile() || st.isSymbolicLink() || JSON.parse(await fs.promises.readFile(lock, 'utf8')).nonce !== nonce) throw new Error('后台启动协调锁已改变，未删除未知文件。');
      await fs.promises.unlink(lock);
    }
  }
}
function readControl(): Control {
  const { control, pause } = paths();
  try {
    const st = fs.lstatSync(control);
    if (!st.isFile() || st.isSymbolicLink() || st.size > 4096) throw new Error();
    const value = JSON.parse(fs.readFileSync(control, 'utf8'));
    if (value.schema !== 1 || !/^[a-f0-9-]{36}$/.test(value.generation) || typeof value.paused !== 'boolean') throw new Error();
    return { ...value, paused: value.paused || fs.existsSync(pause) };
  } catch (error: any) {
    if (error.code !== 'ENOENT') throw new Error('后台启动协调记录无效，未启动服务。');
    const value: Control = { schema: 1, generation: randomUUID(), paused: fs.existsSync(pause) };
    atomicWrite(control, JSON.stringify(value)); return value;
  }
}
export function captureServiceStartIntent(resume = false, signal?: AbortSignal) {
  return withServiceStartLock(() => {
    const value = readControl();
    if (!resume) return value.generation;
    const next: Control = { schema: 1, generation: randomUUID(), paused: false };
    atomicWrite(paths().control, JSON.stringify(next));
    try { fs.unlinkSync(paths().pause); } catch (error: any) { if (error.code !== 'ENOENT') throw new Error('无法恢复字浮后台启动，请检查运行目录。'); }
    return next.generation;
  }, signal);
}
export function pauseServiceStarts() {
  return withServiceStartLock(() => {
    const value: Control = { schema: 1, generation: randomUUID(), paused: true };
    atomicWrite(paths().pause, JSON.stringify({ schema: 1, reason: 'desktop-quit', at: Date.now(), generation: value.generation }));
    atomicWrite(paths().control, JSON.stringify(value));
  });
}
export function commitServiceStartIntent<T>(generation: string, action: () => Promise<T> | T, signal?: AbortSignal) {
  return withServiceStartLock(() => {
    const value = readControl();
    if (value.paused || value.generation !== generation) throw new Error(pausedMessage);
    signal?.throwIfAborted();
    return action();
  }, signal);
}
