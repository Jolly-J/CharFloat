import { bridgeServer } from '../ws-server.js';
import { currentHost } from '../context.js';
import { MsOfficeDriver } from './ms-office-driver.js';
import { EXCEL_METHODS } from '../contracts/host-methods.js';
import { previewPath } from '../runtime.js';
import { normalizeOfficeRequest, normalizeOfficeResponse } from './normalizer.js';
import { BridgeError, routeOfficeFailure, toBridgeError } from '../errors.js';
import type { ChannelParams, ToolResult } from '../contracts/boundary.js';

const queues = new Map<string, Promise<unknown>>();

/**
 * Office.js 通道失败后是否改走 Windows 原生 COM 通道（P1.3）。
 *
 * 判定集中在 errors.routeOfficeFailure：只有确认前一通道未执行，或该方法不修改文档时才回退；
 * 写入类方法在超时、断连、宿主报错时一律拒绝回退并要求先读回，避免跨通道重复写入。
 */
function fallbackToNative<T>(method: string, params: ChannelParams, failure: BridgeError): Promise<T> | never {
  const route = routeOfficeFailure(method, failure, { platform: process.platform, supportedMethods: EXCEL_METHODS });
  if (route.action === 'reject') throw route.error;
  return MsOfficeDriver.windows({
    action: 'excel',
    method,
    params: { ...params, outputPath: method === 'capture_sheet_preview' ? previewPath('png') : undefined }
  }) as Promise<T>;
}

export async function callOffice<T = ToolResult>(method: string, params: ChannelParams = {}, timeout?: number): Promise<T> {
  const host = currentHost();
  const previous = queues.get(host) || Promise.resolve();
  const job = previous.catch(() => {}).then(async () => {
    if (host === 'wps') return bridgeServer.callWps<T>(method, params, timeout);

    // 1. 客户端参数转换：发生在任何宿主调用之前，可确认宿主未执行 → 明确拒绝。
    let normalized: { method: string; params: ChannelParams };
    try {
      normalized = normalizeOfficeRequest(method, params);
    } catch (error: any) {
      return fallbackToNative<T>(method, params, toBridgeError(error, { kind: 'rejected', channel: 'microsoft-officejs', method, executed: 'no' }));
    }

    // 2. 宿主调用：失败种类由 ws-server 在通道边界分类（unavailable / failed / unknown）。
    let rawResult: any;
    try {
      rawResult = await bridgeServer.callOfficeAddon<ToolResult>(normalized.method, normalized.params, timeout);
    } catch (error: any) {
      return fallbackToNative<T>(method, params, toBridgeError(error, { channel: 'microsoft-officejs', method }));
    }

    // 3. 响应转换：宿主已经执行完毕，此时失败属于结果不可判定，禁止回退重放。
    try {
      return normalizeOfficeResponse(method, rawResult, params) as T;
    } catch (error: any) {
      throw toBridgeError(error, { kind: 'failed', channel: 'microsoft-officejs', method, executed: 'unknown' });
    }
  });
  queues.set(host, job);
  try { return await job as T; } finally { if (queues.get(host) === job) queues.delete(host); }
}
