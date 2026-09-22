/**
 * 审计回滚工具定义（P2.3 剩余项：逐类迁移自 gateway.getOpenAiTools）。
 *
 * 已发布顺序中的第 23 条，分类标签为 `audit`。
 * 它的**位置**在 Excel 区间内部（`wps_modify_rows_columns` 与 `wps_auto_fit_columns` 之间），
 * 由 `definitions/index.ts` 显式插回原位——不能"顺手归位"到末尾。
 * 本模块只含网关数组里的回滚定义；审计查询/清理工具在 `tools/audit.ts`（不在网关数组内）。
 */
import type { GatewayToolDefinition } from './shared.js';

/** 审计回滚：已发布顺序第 23–23 条。定义顺序即契约顺序，不得重排。 */
export function auditDefinitionTools(): GatewayToolDefinition[] {
  return [
    {
      type: "function",
      function: {
        name: "wps_rollback",
        description: "根据留痕记录 ID 一键撤销修改，原地恢复表格",
        parameters: {
          type: "object",
          properties: {
            auditId: { type: "string", description: "留痕审计 ID" }
          },
          required: ["auditId"],
          additionalProperties: false
        }
      }
    }
  ];
}
