import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { getToken, runtimeHome, runtimePort, PROTOCOL } from './runtime.js';

/**
 * 统一的服务不可用文案（ISS-50）。
 *
 * 后台没起来时，原来直接把 `fetch failed` 或原始 `urlopen` 错误抛给调用方——
 * 既不说服务名、也不说地址，更不说怎么恢复。这里统一成"服务名 + 地址 + 恢复动作"，
 * 并把原始错误单独放在 `error` 字段里，不与正常字段混在一起。
 */
const SERVICE_NAME = 'wps-bridge';
export function serviceAddress() {
  return `http://127.0.0.1:${runtimePort()}`;
}
export function serviceRecovery(): string[] {
  return [
    '确认 Bridge 后台正在运行：执行 `npx office-agent-bridge --status` 查看 ready 字段',
    '未运行时启动：`npx office-agent-bridge --start`（stdio 客户端首次连接也会自动启动）',
    `仍失败时查看日志 ${path.join(runtimeHome(), 'service.log')}，并用 \`npx office-agent-bridge --doctor\` 复核凭据与端口`,
    '不要反复启动多个实例；端口被旧版占用时先 `npx office-agent-bridge --stop`'
  ];
}
/** 网络层不可达（区别于 HTTP 已返回但业务失败）。 */
function describeUnreachable(error: any) {
  const raw = error?.message || String(error);
  const detail = error?.cause?.code ? `${raw}（${error.cause.code}）` : raw;
  return `${SERVICE_NAME} 后台服务未运行或不可达（地址 ${serviceAddress()}）。恢复步骤：${serviceRecovery().join('；')}。原始错误：${detail}`;
}

export async function serviceRequest(route: string, body?: unknown) {
  let response: Response;
  try {
    response = await fetch(`${serviceAddress()}${route}`, {
      method: body === undefined ? 'GET' : 'POST',
      headers: { Authorization: `Bearer ${getToken()}`, 'Content-Type': 'application/json' },
      body: body === undefined ? undefined : JSON.stringify(body), signal: AbortSignal.timeout(35000)
    });
  } catch (error: any) {
    // fetch 只在"连不上/超时"时抛错，此时后台一定没有处理请求。
    throw new Error(describeUnreachable(error));
  }
  const result: any = await response.json().catch(() => ({}));
  if (!response.ok || result.success === false) throw new Error(result.error || `${SERVICE_NAME} HTTP ${response.status}（地址 ${serviceAddress()}）`);
  return result;
}
export async function probeService() {
  try {
    const r = await fetch(`${serviceAddress()}/health`, { signal: AbortSignal.timeout(1200) });
    if (!r.ok) return { ready: false, occupied: true, service: SERVICE_NAME, address: serviceAddress(), recovery: serviceRecovery(), message: '端口由旧版 Bridge 或其他程序占用；请先退出旧版 Bridge。' };
    const info: any = await r.json();
    if (info.service !== 'wps-bridge' || info.protocol !== PROTOCOL) return { ready: false, occupied: true, service: SERVICE_NAME, address: serviceAddress(), recovery: serviceRecovery(), message: '端口上的服务版本不兼容。' };
    await serviceRequest('/api/v1/status');
    return { ready: true, occupied: true, service: SERVICE_NAME, address: serviceAddress(), info };
  } catch (e: any) {
    const raw = e?.message || String(e);
    if (String(raw).includes('401')) return { ready: false, occupied: true, service: SERVICE_NAME, address: serviceAddress(), recovery: serviceRecovery(), message: 'Bridge 凭据不匹配，请检查运行目录。', error: raw };
    return {
      ready: false,
      occupied: false,
      service: SERVICE_NAME,
      address: serviceAddress(),
      message: `${SERVICE_NAME} 后台服务未运行（地址 ${serviceAddress()}）。`,
      recovery: serviceRecovery(),
      error: raw
    };
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
