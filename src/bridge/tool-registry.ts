/**
 * 本地工具服务实现（运行时注册层，P2.2）。
 *
 * 职责：把 `catalog` 的工具定义、入参校验、能力描述与执行器绑定成协议层需要的 `ToolService`。
 * 位置：位于执行层之上、组装入口之下；**只有组装入口**（`compose.ts`）会引用本模块。
 * 协议层（`ws-server.ts`、`mcp-server.ts`）只认接口，不认这个实现。
 *
 * 契约层保持纯定义，不引用本模块。
 */
import { capabilities, executeCatalogTool, getTools } from './catalog.js';
import { requestContext } from './context.js';
import type { Host } from './context.js';
import type { McpToolDescriptor, ToolCallContext, ToolService } from './contracts/tool-service.js';

/** 统一入口只对 excel_* 允许显式选择 Microsoft，其余（含兼容名 wps_*、office_*）按 WPS 处理。 */
function hostFor(name: string, args: Record<string, unknown>): Host {
  return name.startsWith('excel_') && args.host === 'microsoft' ? 'microsoft' : 'wps';
}

/** 用真实的 catalog 实现绑定工具服务。 */
export function createLocalToolService(): ToolService {
  return {
    list: () => getTools() as McpToolDescriptor[],
    capabilities: () => capabilities(),
    execute: (name, args, context: ToolCallContext) => requestContext.run(
      { sessionId: context.sessionId, host: hostFor(name, args) },
      () => executeCatalogTool(name, args, context.clientName)
    )
  };
}
