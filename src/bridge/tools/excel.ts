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
      // 说明后缀只保留真正必要的两项信息，避免 27×N 的样板挤占有效信息（ISS-36）。
      // 原来写的是"Windows Microsoft Excel 尚待实机验收"，**只提 Windows**，读者会推断 macOS 已验证（ISS-33）；
      // 验收状态改挂到必填的 host 参数上，那里每个调用方必然读到。
      description: `${tool.description} 统一表格入口；host 必填。`,
      inputSchema: {
        ...tool.inputSchema,
        properties: {
          ...tool.inputSchema.properties,
          host: {
            type: 'string',
            enum: ['wps', 'microsoft'],
            description: '明确选择 WPS 表格或 Microsoft Excel；Microsoft 通道在 macOS 与 Windows 均尚未实机验收'
          }
        },
        required: [...(tool.inputSchema.required || []), 'host']
      }
    }));
}
