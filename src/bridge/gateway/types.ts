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

  /**
   * 本次调用的表格目标来自哪里（ISS-02 / ISS-77）。
   *
   * `request` = 调用方显式传了 workbookName；`session-lock` = 回落到本会话的目标锁
   * （可能是更早会话留下的陈旧目标）；`none` = 两者都没有，宿主会用活动文稿。
   * 只影响返回体里的如实说明，不改变目标解析结果。
   */
  targetSource?: 'request' | 'session-lock' | 'none';

  /**
   * 校验层为"统一参数约定"容忍并忽略的参数名（ISS-03）。
   *
   * `wps_rollback` 只认 auditId、`excel_save_workbook` 不认 sheetName，调用方按同一套约定
   * 批量调用会踩坑。装配层现在容忍 `host`/`workbookName`/`sheetName` 这类目标参数：
   * 该工具不适用时忽略，并把被忽略的参数名放进这里，由处理器在返回体里如实回报，
   * 不做"静默吞掉"。
   */
  ignoredParams?: string[];

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
