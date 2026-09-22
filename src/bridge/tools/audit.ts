/**
 * 审计类工具定义（P2.3 逐类迁移）。
 *
 * 契约：读取/清理本地审计留痕，不触达宿主文档。
 * `wps_clear_audit_history` 会改本地状态，因此**不是**只读工具。
 * 入口名 `wps_get_audit_history`、`wps_get_audit_record`、`wps_clear_audit_history` 为已发布入口。
 */
import type { ToolDefinition } from './diagnostics.js';

const RAW: ToolDefinition[] = [
  {
    name: 'wps_get_audit_history',
    description:
      '分页查询本地审计留痕（写操作会留痕）。默认返回最近 5 条摘要。' +
      '留痕分两类：rollbackable=true（patch_cells 的值/公式）可用 wps_rollback 回滚；' +
      'rollbackable=false（格式、条件格式、冻结窗格、行列结构、图表、数据有效性、筛选排序、批注、工作表结构、批量替换、透视表、保存）只登记操作事实、不可回滚。',
    inputSchema: {
      type: 'object',
      properties: {
        limit: { type: 'number', description: '返回条数上限；summary 视图最多 50，detail 视图最多 5。' },
        offset: { type: 'number', description: '跳过前 N 条，用于翻页。' },
        view: {
          type: 'string',
          enum: ['ids', 'summary', 'detail'],
          description: '返回粒度：ids 只给 id/时间/状态；summary 给常用字段（默认，含 rollbackable）；detail 给完整记录与快照。'
        },
        workbookName: { type: 'string', description: '按工作簿名过滤。' },
        sheetName: { type: 'string', description: '按工作表名过滤。' },
        clientName: {
          type: 'string',
          description: '按调用方过滤（审计归属）。HTTP /mcp 通道按客户端 initialize 的 clientInfo 名称记录，stdio 通道固定记为 "stdio MCP"；同名客户端可用记录里的 sessionId 区分。'
        },
        actionType: {
          type: 'string',
          description: '按操作类型过滤。可回滚：update_values / update_formulas；只留痕：format、conditional_format、freeze_panes、insert_row、insert_col、delete_dimension、hide_dimension、sheet_structure、chart、data_validation、find_replace、pivot_table、filter_sort、comment、save、script、rollback。'
        },
        status: {
          type: 'string',
          enum: ['applied', 'rolled_back'],
          description: '按状态过滤：applied = 尚未回滚；rolled_back = 已回滚。'
        },
        fromTimestamp: { type: 'number', description: '起始时间（毫秒时间戳，含）。' },
        toTimestamp: { type: 'number', description: '结束时间（毫秒时间戳，含）。' }
      },
      additionalProperties: false
    }
  },
  {
    name: 'wps_get_audit_record',
    description: '按审计 ID 读取一条完整记录（含修改前后的值/公式快照；只留痕类记录没有快照）。ID 来自写操作返回的 auditId，或 wps_get_audit_history。',
    inputSchema: {
      type: 'object',
      properties: { auditId: { type: 'string', description: '审计记录 ID（写操作返回的 auditId）。' } },
      required: ['auditId'],
      additionalProperties: false
    }
  },
  {
    name: 'wps_clear_audit_history',
    description:
      '清空本地存储的全部审计留痕。**不可恢复**：清空后历史写入将无法再回滚（wps_rollback 依赖这些记录）。' +
      '只清本地留痕，不修改文档内容。',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false }
  }
];

export function auditTools(): ToolDefinition[] {
  return RAW.map(tool => ({ ...tool, inputSchema: structuredClone(tool.inputSchema) }));
}
