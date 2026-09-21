import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { getToken, runtimeHome, runtimePort, PROTOCOL } from './runtime.js';

export async function serviceRequest(route: string, body?: unknown) {
  const response = await fetch(`http://127.0.0.1:${runtimePort()}${route}`, {
    method: body === undefined ? 'GET' : 'POST',
    headers: { Authorization: `Bearer ${getToken()}`, 'Content-Type': 'application/json' },
    body: body === undefined ? undefined : JSON.stringify(body), signal: AbortSignal.timeout(35000)
  });
  const result: any = await response.json();
  if (!response.ok || result.success === false) throw new Error(result.error || `Bridge HTTP ${response.status}`);
  return result;
}
export async function probeService() {
  try {
    const r = await fetch(`http://127.0.0.1:${runtimePort()}/health`, { signal: AbortSignal.timeout(1200) });
    if (!r.ok) return { ready: false, occupied: true, message: '端口由旧版 Bridge 或其他程序占用；请先退出旧版 Bridge。' };
    const info: any = await r.json();
    if (info.service !== 'wps-bridge' || info.protocol !== PROTOCOL) return { ready: false, occupied: true, message: '端口上的服务版本不兼容。' };
    await serviceRequest('/api/v1/status');
    return { ready: true, occupied: true, info };
  } catch (e: any) {
    if (String(e.message).includes('401')) return { ready: false, occupied: true, message: 'Bridge 凭据不匹配，请检查运行目录。' };
    return { ready: false, occupied: false, message: e.message };
  }
}
let starting: Promise<void> | undefined;
export function ensureService(entry: string) {
  if (starting) return starting;
  starting = (async () => {
    const state = await probeService();
    if (state.ready) return;
    if (state.occupied) throw new Error(state.message);
    getToken(true);
    const logPath = path.join(runtimeHome(), 'service.log');
    if (fs.existsSync(logPath) && fs.statSync(logPath).size > 2 * 1024 * 1024) fs.renameSync(logPath, `${logPath}.previous`);
    const log = fs.openSync(logPath, 'a', 0o600);
    const child = spawn(process.execPath, [entry, '--serve'], {
      detached: true, windowsHide: true, stdio: ['ignore', log, log],
      env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }
    });
    fs.closeSync(log);
    let launchError: Error | undefined;
    child.on('error', error => { launchError = error; });
    child.unref();
    for (let n = 0; n < 40; n++) {
      await new Promise(r => setTimeout(r, 150));
      if (launchError) throw launchError;
      if ((await probeService()).ready) return;
    }
    throw new Error(`后台启动失败。查看 ${logPath}，不要反复启动多个实例。`);
  })().finally(() => { starting = undefined; });
  return starting;
}
