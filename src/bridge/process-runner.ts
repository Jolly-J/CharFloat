import { spawn } from 'node:child_process';

/** No shell interpolation. Native jobs run asynchronously so HTTP and heartbeat stay responsive. */
export function runProcess(command: string, args: string[], input = '', timeout = 20000): Promise<string> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] });
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
    child.on('close', code => finish(code ? new Error(err || `原生进程退出 ${code}`) : undefined));
    child.stdin.on('error', () => {});
    child.stdin.end(input);
  });
}
