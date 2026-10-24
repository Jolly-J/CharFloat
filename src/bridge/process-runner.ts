import { spawn } from 'node:child_process';
import type { ChildProcess } from 'node:child_process';
import { stopOwnedChild } from './owned-process.js';

const ownedJobs = new Set<ChildProcess>();
let stopping = false;
export async function stopNativeJobs() {
  stopping = true;
  const results = await Promise.allSettled([...ownedJobs].map(child => stopOwnedChild(child)));
  if (results.some(r => r.status === 'rejected')) throw new Error('部分字浮原生辅助进程未能退出。');
}

/** No shell interpolation. Native jobs run asynchronously so HTTP and heartbeat stay responsive. */
export function runProcess(command: string, args: string[], input = '', timeout = 20000): Promise<string> {
  if (stopping) return Promise.reject(new Error('字浮后台正在退出，未执行新的原生操作。'));
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] });
    ownedJobs.add(child);
    child.once('close', () => ownedJobs.delete(child));
    let out = '', err = '', settled = false;
    const finish = (error?: Error) => {
      if (settled) return;
      settled = true; clearTimeout(timer);
      if (error) reject(error); else resolve(out.trim());
    };
    const timer = setTimeout(() => { child.kill(); finish(new Error('原生调用超时；结果未知。请读取目标状态，勿自动重试写入。')); }, timeout);
    child.stdout.on('data', b => {
      out += b.toString('utf8');
      if (Buffer.byteLength(out) > 32 * 1024 * 1024) { child.kill(); finish(new Error('原生返回超过 32 MB，请缩小范围')); }
    });
    child.stderr.on('data', b => { err = (err + b.toString('utf8')).slice(-4000); });
    child.on('error', finish);
    child.on('close', (code, signal) => finish(code !== 0 ? new Error(err || `原生进程已中断（${signal || code}）；执行结果未知，请先读回。`) : undefined));
    child.stdin.on('error', () => {});
    child.stdin.end(input);
  });
}
