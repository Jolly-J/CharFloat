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
        description: "按 auditId 回滚一次 patch_cells 写入的值与公式。覆盖边界：只有值与公式；样式、条件格式、图表、行列/工作表结构变更、Word/PPT 操作、脚本执行都不在快照内，无法用本工具撤销。回滚前会比对目标区域当前内容，存在后续修改时拒绝覆盖。参数只接受 auditId（传 host、sheetName 等会报‘未知参数’）。",
        parameters: {
          type: "object",
          properties: {
            auditId: { type: "string", description: "审计记录 ID，取自 patch_cells 返回体的 auditId，或审计历史查询（view='ids'）" }
          },
          required: ["auditId"],
          additionalProperties: false
        }
      }
    }
  ];
}
