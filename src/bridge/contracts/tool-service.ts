/**
 * 工具服务接口（契约层）。
 *
 * 边界（改造计划 P2.2）：
 * - 本模块只声明**接口与数据结构**，不引用任何执行器、注册表、服务实例或 catalog。
 * - 协议层（`ws-server.ts`、`mcp-server.ts`）通过**注入**的 `ToolService` 查询与执行工具，
 *   不静态反向引用 `catalog.ts`，从而解开协议层与执行层的循环依赖。
 * - 具体实现由运行时组装入口（`compose.ts`）绑定；组装入口是唯一知道"谁实现谁"的地方。
 */

export interface McpToolDescriptor {
  name: string;
  description: string;
  inputSchema: unknown;
  annotations: { readOnlyHint: boolean; openWorldHint: boolean };
}

export type MaybePromise<T> = T | Promise<T>;

/**
 * 一次工具调用的身份信息。
 *
 * `sessionId` 是**会话隔离的载体**：同一客户端内保持稳定，不同客户端必须不同。
 * 它必须随每次调用显式传递——`AsyncLocalStorage` 只在进程内生效，跨 HTTP 转发不会自动携带，
 * 漏传会让后台把所有调用方归入同一个默认会话（P2.2 曾因此出现不同 stdio 客户端文档锁串扰）。
 */
export interface ToolCallContext {
  sessionId: string;
  /** 审计留痕使用的客户端名称。 */
  clientName: string;
}

/**
 * 协议层需要的全部工具能力。
 *
 * 同时由本地实现（`catalog.ts`）与远程代理实现（stdio 客户端经后台服务）提供，
 * 因此方法允许返回 Promise。
 */
export interface ToolService {
  /** 工具清单，形状与 MCP `tools/list` 一致。 */
  list(): MaybePromise<McpToolDescriptor[]>;
  /** 能力与运行状态，形状与 `bridge_get_capabilities` 一致。 */
  capabilities(): MaybePromise<unknown>;
  /** 执行工具；`context` 同时携带会话隔离标识与审计客户端名，实现方不得丢弃。 */
  execute(name: string, args: Record<string, unknown>, context: ToolCallContext): Promise<unknown>;
}

/** 延迟求值的服务提供者，避免组装顺序影响可用性。 */
export type ToolServiceProvider = () => ToolService;
