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
  /**
   * 参数安全护栏（CAP-50）：拒绝原因不是"当前通道不支持"，而是"这组参数会让**任何**通道
   * 做出破坏性动作"（例：空搜索串的批量替换、空串检索）。
   *
   * 为什么要单独标记：`routeOfficeFailure` 允许把"执行前拒绝"转交 Windows 原生 COM 通道重放，
   * 因为多数拒绝属于**能力缺口**（Office.js 没实现，COM 有独立实现，例如多级排序、只有筛选按钮）。
   * 安全护栏不属此列——原生脚本没有等价护栏，重放等于绕过护栏：
   * `excel.ps1` 的 find_and_replace 对空查询走 `IndexOf('')`（恒真）并按空正则逐格替换，
   * 会把请求里"已判定为破坏性"的动作原样执行一遍。
   */
  unsafe?: boolean;
}

const NOT_EXECUTED: ReadonlySet<BridgeErrorKind> = new Set<BridgeErrorKind>(['unavailable', 'rejected']);

export class BridgeError extends Error {
  readonly kind: BridgeErrorKind;
  readonly channel: BridgeChannel;
  readonly method?: string;
  readonly executed: ExecutedState;
  /** 见 {@link BridgeErrorOptions.unsafe}：参数安全护栏，禁止换通道重放。 */
  readonly unsafe: boolean;

  constructor(kind: BridgeErrorKind, message: string, options: BridgeErrorOptions) {
    super(message, options.cause === undefined ? undefined : { cause: options.cause });
    this.name = 'BridgeError';
    this.kind = kind;
    this.channel = options.channel;
    this.method = options.method;
    this.unsafe = options.unsafe === true;
    // 保守默认：只有 unavailable / rejected 才允许断言"未执行"，其余一律按可能已执行处理。
    this.executed = options.executed ?? (NOT_EXECUTED.has(kind) ? 'no' : 'unknown');
  }

  /** 结构化描述，供日志与响应使用；不包含宿主返回的文档内容。 */
  toJSON() {
    return { name: this.name, kind: this.kind, channel: this.channel, method: this.method, executed: this.executed, unsafe: this.unsafe, message: this.message };
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
  if (failure.unsafe) {
    // CAP-50：参数安全护栏的拒绝**不得**换通道。原生脚本没有等价校验，重放会把
    // "已判定为破坏性"的动作在另一个通道原样执行（空搜索串 → 全区域替换）。
    // 这条必须排在"不在支持列表内"之前：否则会把安全拒绝说成"COM 不支持该方法"，指向错误的自救方向。
    return {
      action: 'reject',
      error: new BridgeError(
        failure.kind,
        `${failure.message}｜已阻止改用 Windows 原生 COM 通道重放：本次拒绝是**参数安全护栏**（不是当前通道能力不足），` +
        `原生脚本没有等价校验，重放会做出同样的破坏性动作（如空搜索串命中全区域并逐格替换）。请修正参数后重试。`,
        // 保留原"是否已执行"：安全护栏默认是执行前拒绝（executed='no'），
        // 但若调用方给出了更保守的判定，不得在此改写成"未执行"（那会诱发重试）。
        { channel: failure.channel, method, executed: failure.executed, unsafe: true }
      )
    };
  }
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

/** 布局失败的实际宿主读回投影：只保留数值/已应用标记及已知告警类别。 */
function projectPptLayoutFailure(result: Record<string, any>) {
  function numbers(source: any, keys: string[]) {
    const out: Record<string, number | null> = {};
    if (source && typeof source === 'object') for (const key of keys) {
      if (source[key] === null || (typeof source[key] === 'number' && Number.isFinite(source[key]))) out[key] = source[key];
    }
    return out;
  }
  const snapshotKeys = ['slideWidth', 'slideHeight', 'slideSize', 'slideOrientation', 'designCount', 'layoutCount', 'slideCount'];
  const applied: Record<string, any> = numbers(result.applied, ['slideWidth', 'slideHeight', 'slideOrientation']);
  if (typeof result.applied?.templatePath === 'string' || result.applied?.templateApplied === true) applied.templateApplied = true;
  if (result.applied?.layout && typeof result.applied.layout === 'object') {
    applied.layout = numbers(result.applied.layout, ['appliedPages']);
  }
  const categories: [RegExp, string, string, string][] = [
    [/^设置页宽失败/, 'slideWidth', 'write_failed', '设置页宽失败'],
    [/^设置页高失败/, 'slideHeight', 'write_failed', '设置页高失败'],
    [/^设置页面方向失败/, 'slideOrientation', 'write_failed', '设置页面方向失败'],
    [/^页宽未生效/, 'slideWidth', 'readback_mismatch', '页宽未生效'],
    [/^页高未生效/, 'slideHeight', 'readback_mismatch', '页高未生效'],
    [/^页面方向未生效/, 'slideOrientation', 'readback_mismatch', '页面方向未生效'],
    [/^套用模板失败/, 'template', 'apply_failed', '套用模板失败'],
    [/^(套用版式失败|版式对象已找到，但没有页面成功套用)/, 'layout', 'apply_failed', '套用版式失败'],
    [/^(不认识的 preset|页面预设无效)/, 'preset', 'invalid_argument', '页面预设无效'],
    [/^(不认识的 orientation|页面方向无效)/, 'slideOrientation', 'invalid_argument', '页面方向无效']
  ];
  const rawWarnings = Array.isArray(result.warnings) ? result.warnings : [];
  const warningDetails = rawWarnings.slice(0, 200).map((warning: unknown) => {
    const text = typeof warning === 'string' ? warning : '';
    const category = categories.find(([pattern]) => pattern.test(text));
    const reason = /host_rejected|permission|denied|权限|拒绝/i.test(text) ? 'host_rejected'
      : /unsupported|不支持|不可写/i.test(text) ? 'unsupported'
      : /not_found|找不到|未找到/i.test(text) ? 'not_found'
      : /越界|invalid|无效|out.of.range/i.test(text) ? 'invalid_argument' : 'host_failure';
    return { field: category?.[1] || 'unknown', code: category?.[2] || 'unclassified_warning', reason,
      summary: category?.[3] || '未分类布局告警' };
  });
  return { diagnosticType: 'ppt_layout' as const,
    applied, before: numbers(result.before, snapshotKeys), after: numbers(result.after, snapshotKeys),
    warnings: warningDetails.map(x => `${x.summary}（${x.reason}）`), warningDetails,
    warningsTruncated: rawWarnings.length > 200 || (result.diagnosticType === 'ppt_layout' && result.warningsTruncated === true),
    recovery: '已有操作可能生效；先按 after 读回目标，不要重放整批布局操作。' };
}

/** 协议边界只导出经过审查的标量字段，绝不序列化 cause/stack/宿主对象。 */
export function toolFailureResponse(value: unknown, toolName?: string) {
  const result = value && typeof value === 'object' ? value as Record<string, any> : undefined;
  const supplied = result?.errorDetails;
  let failure = toBridgeError(value instanceof Error ? value : new Error(
    typeof result?.error === 'string' ? result.error : typeof result?.message === 'string' ? result.message : '操作失败'
  ), { channel: 'bridge' });
  if (!(value instanceof BridgeError) && supplied && typeof supplied === 'object') {
    const kind: BridgeErrorKind = ['unavailable', 'rejected', 'failed', 'unknown'].includes(supplied.kind) ? supplied.kind : 'unknown';
    const channel: BridgeChannel = ['wps-addon', 'microsoft-officejs', 'microsoft-native', 'bridge'].includes(supplied.channel) ? supplied.channel : 'bridge';
    failure = new BridgeError(kind, failure.message, {
      channel, method: typeof supplied.method === 'string' ? supplied.method : undefined,
      // 只有提供明确的 no 才透传未执行；不能由 kind 自动推断远端状态。
      executed: supplied.executed === 'no' ? 'no' : 'unknown', unsafe: supplied.unsafe === true
    });
  }
  const { kind, channel, method, executed, unsafe, message } = failure.toJSON();
  const errorDetails = { kind, channel, method, executed, unsafe, message };
  // 部分完成证据只接受对象 ID/计数；不把任意结果或文档文字带进错误响应。
  const partial: Record<string, number | number[]> = {};
  const source = result?.partial;
  if (source && typeof source === 'object') {
    for (const key of ['slideId', 'shapeId', 'completedCount', 'failedIndex']) {
      if (typeof source[key] === 'number' && Number.isFinite(source[key])) partial[key] = source[key];
    }
    for (const key of ['completedShapeIds', 'completedSlideIds']) {
      if (Array.isArray(source[key])) partial[key] = source[key].filter((x: unknown) => typeof x === 'number' && Number.isFinite(x)).slice(0, 2000);
    }
  }
  const layout = result && (toolName === 'wps_ppt_configure_layout' || result.diagnosticType === 'ppt_layout')
    ? projectPptLayoutFailure(result) : {};
  const radius: Record<string, unknown> = {};
  if(result && (toolName === 'wps_ppt_set_corner_radius' || result.diagnosticType === 'ppt_corner_radius')) {
    radius.diagnosticType='ppt_corner_radius';
    if(typeof result.operationId==='string' && /^[A-Za-z0-9_-]{1,40}$/.test(result.operationId)) radius.operationId=result.operationId;
    if(['planned','applying','applied','partial','unknown','rollingBack','rolledBack','rollbackFailed'].includes(result.state)) radius.state=result.state;
    for(const k of ['liveApplied','saved','ooxmlVerified','previewVerified']) if(typeof result[k]==='boolean') radius[k]=result[k];
    radius.requiresReadbackBeforeRetry=true;
    radius.receipts=Array.isArray(result.receipts)?result.receipts.slice(0,100).map((r:any)=>{
      const receipt:Record<string,unknown>={};
      for(const k of ['oldId','newId']) if(typeof r?.[k]==='number'&&Number.isFinite(r[k])) receipt[k]=r[k];
      if(['pasted','positioned','hidden','replaced'].includes(r?.phase)) receipt.phase=r.phase;
      return receipt;
    }):[];
  }
  return { success: false as const, error: message, errorDetails, ...(Object.keys(partial).length ? { partial } : {}), ...layout, ...radius };
}
