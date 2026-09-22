/**
 * 统一 Excel 入口派生（P2.3 逐类迁移）。
 *
 * 职责：把网关提供的 `wps_*` 表格工具，按宿主方法路由表派生出统一的 `excel_*` 入口。
 * 兼容契约（不得破坏）：
 * - 原 `wps_*` 兼容名**同时保留**，`excel_*` 是新增的并列入口；
 * - `excel_*` 追加必填 `host`（`wps` | `microsoft`），因此原有的 `required` 必须保留并追加；
 * - 派生顺序与过滤条件必须与迁移前一致（影响 tools/list 与契约快照）。
 */
import { EXCEL_METHODS } from '../contracts/host-methods.js';
import type { ToolDefinition } from './diagnostics.js';

/** 网关输出的工具形状（含 name/description/inputSchema）。 */
export interface GatewayTool {
  name: string;
  description: string;
  inputSchema: any;
}

export function unifiedExcelTools(existing: GatewayTool[]): ToolDefinition[] {
  return existing
    .filter(tool => EXCEL_METHODS.includes(tool.name.replace(/^wps_/, '') as any))
    .map(tool => ({
      ...tool,
      name: tool.name.replace(/^wps_/, 'excel_'),
      description: `${tool.description} 统一表格入口；host 必填。Windows Microsoft Excel 尚待实机验收。`,
      inputSchema: {
        ...tool.inputSchema,
        properties: {
          ...tool.inputSchema.properties,
          host: { type: 'string', enum: ['wps', 'microsoft'], description: '明确选择 WPS 表格或 Microsoft Excel' }
        },
        required: [...(tool.inputSchema.required || []), 'host']
      }
    }));
}
