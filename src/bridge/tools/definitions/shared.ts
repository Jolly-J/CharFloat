/**
 * 工具定义的共享类型与上下文（P2.3 剩余项：工具描述符逐类迁移）。
 *
 * 定位：**纯数据层**。这里只放类型与拼描述用的小工具，不 import gateway、
 * catalog、contracts 或任何执行器；`gateway.ts` 可以单向 import 本目录，
 * 反向依赖会形成回环（由 tests/contracts-boundary.test.ts 的 Tarjan SCC 守卫）。
 *
 * 形状固定在 OpenAI function-calling 协议上（`{ type: 'function', function: {...} }`），
 * 与迁移前 `UniversalGateway.getOpenAiTools()` 的返回形状逐字段一致。
 */

/** 单条对外工具定义（OpenAI function-calling 形状）。 */
export interface GatewayToolDefinition {
  type: 'function';
  function: {
    name: string;
    description: string;
    parameters: Record<string, any>;
  };
}

/**
 * 定义上下文：只承载"按调用参数变化"的文本，避免各定义模块自行拼装导致文案漂移。
 *
 * `activeWbHint` 是 `getOpenAiTools(activeWbHint?)` 的入参，参与 `wps_get_workspace_summary`
 * 的 description 分支；`wbDesc` 是所有表格工具 `workbookName` 参数的描述文本。
 */
export interface DefinitionContext {
  activeWbHint?: string;
  wbDesc: string;
}

/** 构造定义上下文（文案与迁移前 gateway 内联版本逐字节一致）。 */
export function createDefinitionContext(activeWbHint?: string): DefinitionContext {
  return {
    activeWbHint,
    wbDesc: activeWbHint
      ? `目标工作簿名称。当前已连接打开: '${activeWbHint}'。操作该文件时直接省略本参数即可，请先从工作区摘要确认目标。`
      : "目标工作簿名称，例如 '明细.xlsx'，防止多文件焦点漂移，请先从工作区摘要确认目标。"
  };
}
