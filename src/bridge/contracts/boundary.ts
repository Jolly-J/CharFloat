/**
 * 请求/响应边界类型（改造计划 P2.4）。
 *
 * 目的：把跨层传递的参数与结果从裸 `any` 收敛为具名类型，让边界可读、可检索、可约束；
 * **不做大规模类型重写**——内部实现仍可保留具体类型，只统一"层与层之间"的形状。
 *
 * 约定：
 * - 进入执行层的参数一律是"字符串键的对象"，值形状由各工具的 JSON Schema 约束；
 * - 跨通道（WPS 加载项 / Office.js / 原生驱动）传输的载荷必须可 JSON 序列化；
 * - 返回值统一为 `ToolResult`，协议层负责决定如何包装成 MCP/HTTP 响应。
 */

/** 工具入参：键为字符串，值的具体形状由该工具的 inputSchema 校验。 */
export type ToolArgs = Record<string, unknown>;

/** 跨宿主通道传输的载荷，必须可 JSON 序列化（换行/对象代理等一律转换后再传）。 */
export type ChannelParams = Record<string, unknown>;

/**
 * 工具返回值。
 *
 * `success` 为 false 表示宿主明确报告失败（此时错误信息在 `error`/`message`）；
 * 省略或为 true 表示成功。注意：**响应丢失或超时不在此表达**——那属于抛出的执行结果未知错误。
 */
export interface ToolResult {
  success?: boolean;
  error?: string;
  message?: string;
  [key: string]: unknown;
}

/** 判断一个值是否为合法工具入参。 */
export function isToolArgs(value: unknown): value is ToolArgs {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/** 把任意入参收敛为 ToolArgs；非对象一律视为空参数，避免把原始值透传到宿主。 */
export function asToolArgs(value: unknown): ToolArgs {
  return isToolArgs(value) ? value : {};
}
