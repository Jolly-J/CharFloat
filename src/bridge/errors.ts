/**
 * 失败分类与跨通道回退策略。
 *
 * 目的（改造计划 P1.2 / P1.3）：
 * 1. 把失败显式分为「执行前不可用 / 明确拒绝 / 执行失败 / 结果未知」四类，并记录前一通道是否已把请求交给宿主。
 * 2. 据此决定是否允许把同一请求转交另一个宿主通道重放：只有**确认前一通道未执行**，或**该方法不修改文档内容**时才允许。
 *    超时、断连、宿主已报错或响应转换失败一律视为结果不可判定，不跨通道重放，先读回。
 *
 * 依赖：仅契约层 `contracts/host-methods.ts`。不引用 gateway、adapter、catalog 或服务实例，
 * 因此不会加深 catalog → gateway → adapter → catalog 的既有依赖回环（P2.1）。
 */

import { isReplaySafeMethod } from './contracts/host-methods.js';

export type BridgeErrorKind =
  /** 执行前不可用：通道未连接、未启动。可确认宿主未执行。 */
  | 'unavailable'
  /** 明确拒绝：参数或目标在执行前被拒。可确认宿主未执行。 */
  | 'rejected'
  /** 执行失败：宿主已接收请求并返回失败；是否已部分生效未知。 */
  | 'failed'
  /** 结果未知：超时、断连、响应丢失或响应转换失败；宿主可能已执行。 */
  | 'unknown';

export type BridgeChannel = 'wps-addon' | 'microsoft-officejs' | 'microsoft-native' | 'bridge';

/** 前一通道是否已把请求交给宿主执行。 */
export type ExecutedState = 'no' | 'unknown';

export interface BridgeErrorOptions {
  channel: BridgeChannel;
  /** 宿主方法名，例如 patch_cells、ppt_manage_slides。 */
  method?: string;
  executed?: ExecutedState;
  cause?: unknown;
}

const NOT_EXECUTED: ReadonlySet<BridgeErrorKind> = new Set<BridgeErrorKind>(['unavailable', 'rejected']);

export class BridgeError extends Error {
  readonly kind: BridgeErrorKind;
  readonly channel: BridgeChannel;
  readonly method?: string;
  readonly executed: ExecutedState;

  constructor(kind: BridgeErrorKind, message: string, options: BridgeErrorOptions) {
    super(message, options.cause === undefined ? undefined : { cause: options.cause });
    this.name = 'BridgeError';
    this.kind = kind;
    this.channel = options.channel;
    this.method = options.method;
    // 保守默认：只有 unavailable / rejected 才允许断言"未执行"，其余一律按可能已执行处理。
    this.executed = options.executed ?? (NOT_EXECUTED.has(kind) ? 'no' : 'unknown');
  }

  /** 结构化描述，供日志与响应使用；不包含宿主返回的文档内容。 */
  toJSON() {
    return { name: this.name, kind: this.kind, channel: this.channel, method: this.method, executed: this.executed, message: this.message };
  }
}

export function isBridgeError(value: unknown): value is BridgeError {
  return value instanceof BridgeError;
}

/**
 * 把任意抛出物转成已分类错误。已分类的原样返回，避免二次包装丢失种类。
 * 未分类错误保守归为 unknown（结果未知），因为调用方无法证明宿主未执行。
 */
export function toBridgeError(value: unknown, options: BridgeErrorOptions & { kind?: BridgeErrorKind }): BridgeError {
  if (isBridgeError(value)) return value;
  const message = value instanceof Error ? value.message : String(value);
  return new BridgeError(options.kind ?? 'unknown', message, { ...options, cause: options.cause ?? value });
}

/**
 * 只读与重放安全性的判定已上移到契约层（`contracts/host-methods.ts`），
 * 这里只保留失败分类与回退策略，避免两处各维护一份能力集合。
 */
export { isReadOnlyTool, isReplaySafeMethod, MCP_READ_ONLY_TOOLS, REPLAY_SAFE_METHODS } from './contracts/host-methods.js';

/**
 * 是否允许把这个方法转交另一个宿主通道重放。
 * @param method 宿主方法名（不含 wps_/excel_ 前缀也可）
 * @param failure 已分类或原始失败
 */
export function mayReplayOnAnotherChannel(method: string, failure: unknown): boolean {
  if (isReplaySafeMethod(method)) return true;
  return toBridgeError(failure, { channel: 'bridge' }).executed === 'no';
}

export interface OfficeFailureRouting {
  platform: NodeJS.Platform;
  /** 替代通道（Windows 原生 COM）支持的方法白名单。 */
  supportedMethods: readonly string[];
}

export type OfficeFailureRoute = { action: 'fallback' } | { action: 'reject'; error: BridgeError };

/**
 * 决定 Microsoft Office.js 通道失败后是否改走 Windows 原生通道。
 *
 * 这是 P1.3 的唯一判定入口，抽成纯函数以便在不具备 Windows 桌面环境的机器上做故障注入验证（P1.4）。
 */
export function routeOfficeFailure(method: string, failure: BridgeError, routing: OfficeFailureRouting): OfficeFailureRoute {
  if (routing.platform !== 'win32') return { action: 'reject', error: failure };
  if (!routing.supportedMethods.includes(method)) {
    // 替代通道不支持该方法：只补充"为什么没回退"，不改写前一通道的错误种类与"是否已执行"。
    // 尤其不能把 unknown（可能已执行）写成 no（确认未执行）——那会让调用方错误地重放。
    return {
      action: 'reject',
      error: new BridgeError(
        failure.kind,
        `${failure.message}｜未回退：${method} 不在 Windows 原生通道支持列表内，可用 bridge_get_capabilities 查询能力清单。`,
        { channel: failure.channel, method, executed: failure.executed }
      )
    };
  }
  if (!mayReplayOnAnotherChannel(method, failure)) {
    return {
      action: 'reject',
      error: new BridgeError(
        failure.kind,
        `${failure.message}｜已阻止自动改用 Windows 原生 COM 通道重放：上一通道结果不可判定且该方法可能修改文档，重放可能造成重复写入。请先读回目标确认实际状态；确需重试时由调用方显式发起。`,
        { channel: failure.channel, method, executed: failure.executed }
      )
    };
  }
  return { action: 'fallback' };
}
