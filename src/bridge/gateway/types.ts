/**
 * 网关执行器上下文（P3.1）。
 *
 * 处理器按类拆分到本目录各模块，但**不各自 import 执行器**：`callOffice`、`auditStore`、
 * `MsOfficeDriver`、`TargetLockStore`、`bridgeServer`、`requestContext`、`previewPath` 等
 * 共享依赖全部由门面 `gateway.ts` 组装进本上下文传入，避免处理器与门面之间形成隐式耦合或回环。
 *
 * 依赖方向：本文件只有类型声明与类型级 import，不产生运行时依赖边；
 * 处理器模块只依赖本文件与纯类型，不依赖 `contracts/`、`tools/`、`ws-server.ts`、`mcp-server.ts`。
 */
import type { AuditStore } from '../audit-store.js';
import type { Host } from '../context.js';
import type { MsOfficeDriver } from '../office/ms-office-driver.js';
import type { WpsBridgeServer } from '../ws-server.js';
import type { AuditRecord, PatchResult, RangeData } from '../types.js';
import type { TargetLockStore } from './locks.js';

/** 与 `office/adapter.ts` 的 `callOffice` 同签名（含泛型返回值）。 */
export type CallOffice = <T = any>(method: string, params?: Record<string, any>, timeout?: number) => Promise<T>;

/**
 * 单次工具调用的全部输入与依赖。
 *
 * `args` 是 `executeTool` 在锁定注入后传入的**同一个**对象引用：
 * 原 switch 也直接读写该对象，处理器沿用 `ctx.args` 可保持该语义不变（无参数拷贝/归一化）。
 */
export interface GatewayContext {
  /** 正在执行的工具名（已由 catalog 把 `excel_*` 重写为 `wps_*`）。 */
  name: string;
  args: any;
  clientName: string;

  /** 已 `await` 的锁定注入结果，仅供需要回读目录名称的处理器使用。 */
  locks: { word?: string; excel?: string; ppt?: string };

  callOffice: CallOffice;
  auditStore: AuditStore;
  MsOfficeDriver: typeof MsOfficeDriver;
  TargetLockStore: typeof TargetLockStore;
  bridgeServer: WpsBridgeServer;
  requestContext: { run<T>(store: { sessionId: string; host: Host }, callback: () => T): T };
  currentHost: () => Host;
  currentSession: () => string;
  previewPath: (extension: string) => string;
  /** 剪贴板取图（实现在 script.ts，经上下文注入，避免 excel.ts 反向依赖门面）。 */
  extractClipboardImageBase64: () => Promise<string>;
}

/** 处理器只接收上下文，不接收位置参数；返回值即工具结果。 */
export type Handler = (ctx: GatewayContext) => Promise<unknown>;

/** 处理器模块导出形状：工具名 → 处理器。 */
export type HandlerTable = Record<string, Handler>;

export type { AuditRecord, PatchResult, RangeData };
