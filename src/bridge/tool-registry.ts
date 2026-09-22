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
    // ⚠️ 只发 MCP 规范里的字段。
    // `getTools()` 返回的对象还带一个**我们自己的内部字段 `toolClass`**（工具分类，由 assembleTools 附加）。
    // 此前这里用 `as McpToolDescriptor[]` 强制转换，**类型断言把多余字段藏了起来**，
    // 于是非标准字段随 `tools/list` 一起发给了客户端。
    // 真机现象：WorkBuddy 连接器里服务连上了（绿点）却**显示不出工具清单**，
    // 而同界面的其它 MCP（工具少）正常显示。
    // 规范字段：name / title / description / inputSchema / outputSchema / annotations / _meta。
    // `toolClass` 仍保留在内部对象上供工具分类使用，只是**不出现在对外响应里**。
    list: () => getTools().map((tool) => {
      const { toolClass: _internalToolClass, ...standard } = tool as McpToolDescriptor & { toolClass?: string };
      return standard as McpToolDescriptor;
    }),
    capabilities: () => capabilities(),
    execute: (name, args, context: ToolCallContext) => requestContext.run(
      { sessionId: context.sessionId, host: hostFor(name, args) },
      () => executeCatalogTool(name, args, context.clientName)
    )
  };
}
