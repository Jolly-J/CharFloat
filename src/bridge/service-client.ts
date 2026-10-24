import fs from 'node:fs';
import { toolFailureResponse } from './errors.js';
import path from 'node:path';
import { spawn, type ChildProcess } from 'node:child_process';
import { getToken, runtimeHome, runtimePort, PROTOCOL } from './runtime.js';
import { captureServiceStartIntent, commitServiceStartIntent, pauseServiceStarts } from './service-start-control.js';

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
    '在字浮客户端检查后台状态；源码环境可执行 `node dist/bridge/cli.cjs --status` 查看 ready 字段',
    '未运行时从字浮客户端启动后台；源码环境执行 `node dist/bridge/cli.cjs --start`（stdio 首次连接也会自动启动）',
    `仍失败时查看日志 ${path.join(runtimeHome(), 'service.log')}；源码环境用 \`node dist/bridge/cli.cjs --doctor\` 复核凭据与端口`,
    '不要反复启动多个实例；端口被旧版占用时先从字浮客户端停止后台，或在源码环境执行 `node dist/bridge/cli.cjs --stop`'
  ];
}
/** 网络层不可达（区别于 HTTP 已返回但业务失败）。 */
function describeUnreachable(error: any) {
  const raw = error?.message || String(error);
  const detail = error?.cause?.code ? `${raw}（${error.cause.code}）` : raw;
  return `${SERVICE_NAME} 后台服务未运行或不可达（地址 ${serviceAddress()}）。恢复步骤：${serviceRecovery().join('；')}。原始错误：${detail}`;
}

export async function serviceRequest(route: string, body?: unknown, signal?: AbortSignal) {
  let response: Response;
  try {
    response = await fetch(`${serviceAddress()}${route}`, {
      method: body === undefined ? 'GET' : 'POST',
      headers: { Authorization: `Bearer ${getToken()}`, 'Content-Type': 'application/json' },
      body: body === undefined ? undefined : JSON.stringify(body), signal: signal ? AbortSignal.any([signal, AbortSignal.timeout(35000)]) : AbortSignal.timeout(35000)
    });
  } catch (error: any) {
    // 连接/超时不能证明后台未处理请求；保持 unknown，调用方先读回。
    throw new Error(describeUnreachable(error));
  }
  const result: any = await response.json().catch(() => ({}));
  if (!response.ok || result.success === false) {
    const failure = toolFailureResponse({ ...result, error: result.error || `${SERVICE_NAME} HTTP ${response.status}（地址 ${serviceAddress()}）` });
    throw Object.assign(new Error(failure.error), failure);
  }
  return result;
}
export async function probeService(signal?: AbortSignal) {
  try {
    const r = await fetch(`${serviceAddress()}/health`, { signal: signal ? AbortSignal.any([signal, AbortSignal.timeout(1200)]) : AbortSignal.timeout(1200) });
    if (!r.ok) return { ready: false, occupied: true, service: SERVICE_NAME, address: serviceAddress(), recovery: serviceRecovery(), message: '端口由旧版 Bridge 或其他程序占用；请先退出旧版 Bridge。' };
    const info: any = await r.json();
    if (info.service !== 'wps-bridge' || info.protocol !== PROTOCOL) return { ready: false, occupied: true, service: SERVICE_NAME, address: serviceAddress(), recovery: serviceRecovery(), message: '端口上的服务版本不兼容。' };
    await serviceRequest('/api/v1/status', undefined, signal);
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
export function pauseAutomaticServiceStart() {
  return pauseServiceStarts();
}
export function ensureService(entry: string, options: { signal?: AbortSignal; onSpawn?: (child: ChildProcess) => void; resumeAutomaticStart?: boolean } = {}) {
  const checkCancelled = () => { if (options.signal?.aborted) throw new Error('字浮正在退出，已取消后台启动。'); };
  checkCancelled();
  if (starting) return starting;
  starting = (async () => {
    const generation = await captureServiceStartIntent(options.resumeAutomaticStart, options.signal);
    const state = await probeService(options.signal);
    checkCancelled();
    if (state.ready) return;
    if (state.occupied) throw new Error(state.message);
    const logPath = path.join(runtimeHome(), 'service.log');
    let launchError: Error | undefined;
    let childExited = false;
    let exitCode: number | null = null;
    const child = await commitServiceStartIntent(generation, () => {
      checkCancelled(); getToken(true);
      if (fs.existsSync(logPath) && fs.statSync(logPath).size > 2 * 1024 * 1024) fs.renameSync(logPath, `${logPath}.previous`);
      const log = fs.openSync(logPath, 'a', 0o600);
      try {
        const child = spawn(process.execPath, [entry, '--serve'], {
          detached: true, windowsHide: true, stdio: options.onSpawn ? ['ignore', log, log, 'ipc'] : ['ignore', log, log],
          env: { ...process.env, ELECTRON_RUN_AS_NODE: '1', CHARFLOAT_SERVICE_START_GENERATION: generation }
        });
        child.on('error', error => { launchError = error; });
        child.on('exit', code => { childExited = true; exitCode = code; });
        options.onSpawn?.(child);
        return child;
      } finally { fs.closeSync(log); }
    }, options.signal);
    child.unref();
    child.channel?.unref();
    const readLogTail = () => {
      try {
        if (fs.existsSync(logPath)) {
          return fs.readFileSync(logPath, 'utf8').trim().split('\n').slice(-5).join(' ');
        }
      } catch {}
      return '';
    };
    for (let n = 0; n < 80; n++) {
      await new Promise(r => setTimeout(r, 150));
      checkCancelled();
      if (launchError) throw launchError;
      if (childExited) {
        const tail = readLogTail();
        throw new Error(`后台进程过早退出（退出码 ${exitCode}）。${tail ? `日志: ${tail}` : `查看 ${logPath}`}`);
      }
      const ready = await probeService(options.signal);
      checkCancelled();
      if (ready.ready) {
        if (options.onSpawn && ready.info.pid !== child.pid) throw new Error('后台端口已由其他实例占用，未接管该实例。');
        return;
      }
    }
    const tail = readLogTail();
    throw new Error(`后台启动超时。${tail ? `日志: ${tail}` : `查看 ${logPath}，不要反复启动多个实例。`}`);
  })().finally(() => { starting = undefined; });
  return starting;
}
