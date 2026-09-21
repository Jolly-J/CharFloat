import { bridgeServer } from '../ws-server.js';
import { currentHost } from '../context.js';
import { MsOfficeDriver } from './ms-office-driver.js';
import { EXCEL_METHODS } from '../catalog.js';
import { previewPath } from '../runtime.js';
import { normalizeOfficeRequest, normalizeOfficeResponse } from './normalizer.js';
const queues = new Map<string, Promise<unknown>>();
export async function callOffice<T = any>(method: string, params: any = {}, timeout?: number): Promise<T> {
  const host = currentHost();
  const previous = queues.get(host) || Promise.resolve();
  const job = previous.catch(() => {}).then(async () => {
    if (host === 'wps') return bridgeServer.callWps<T>(method, params, timeout);
    try {
      const normalized = normalizeOfficeRequest(method, params);
      const rawResult = await bridgeServer.callOfficeAddon<any>(normalized.method, normalized.params, timeout);
      return normalizeOfficeResponse(method, rawResult, params) as T;
    } catch (addonError: any) {
      if (process.platform === 'win32') {
        if (!EXCEL_METHODS.includes(method as any)) throw new Error(`Microsoft Excel 尚不支持 ${method}，请查询能力清单。`);
        return MsOfficeDriver.windows({ action: 'excel', method, params: { ...params, outputPath: method === 'capture_sheet_preview' ? previewPath('png') : undefined } });
      }
      throw addonError;
    }

  });
  queues.set(host, job);
  try { return await job as T; } finally { if (queues.get(host) === job) queues.delete(host); }
}

