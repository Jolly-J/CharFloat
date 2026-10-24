import type { ChildProcess } from 'node:child_process';

/** Only accepts a live handle obtained from our own spawn; never discovers/kills by name or PID file. */
export async function stopOwnedChild(child: ChildProcess, options: { graceful?: () => void; graceMs?: number; forceMs?: number; force?: boolean } = {}) {
  let exited = child.exitCode !== null || child.signalCode !== null || !child.pid;
  const markExited = () => { exited = true; };
  child.once('exit', markExited); child.once('close', markExited);
  const wait = async (ms: number) => {
    const until = Date.now() + ms;
    while (!exited && Date.now() < until) await new Promise(resolve => setTimeout(resolve, Math.min(25, ms)));
    return exited;
  };
  try {
    if (exited) return;
    if (options.graceful) { try { options.graceful(); } catch { /* bounded handle-based fallback below */ } }
    else child.kill('SIGTERM');
    if (await wait(options.graceMs ?? 2500)) return;
    if (options.force === false) throw new Error('字浮后台未确认完成退出，已取消退出或更新；不会强制结束无法确认的子进程树，请稍后重试。');
    child.kill('SIGTERM');
    if (await wait(options.forceMs ?? 1000)) return;
    child.kill('SIGKILL');
    if (!await wait(options.forceMs ?? 1000)) throw new Error('字浮拥有的子进程未能退出，已停止后续操作，请保留日志后重试。');
  } finally { child.removeListener('exit', markExited); child.removeListener('close', markExited); }
}
