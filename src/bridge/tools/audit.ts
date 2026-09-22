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
    description: '分页筛选修改记录；默认 5 条摘要',
    inputSchema: {
      type: 'object',
      properties: {
        limit: { type: 'number' }, offset: { type: 'number' },
        view: { type: 'string', enum: ['ids', 'summary', 'detail'] },
        workbookName: { type: 'string' }, sheetName: { type: 'string' }, clientName: { type: 'string' },
        actionType: { type: 'string' }, status: { type: 'string', enum: ['applied', 'rolled_back'] },
        fromTimestamp: { type: 'number' }, toTimestamp: { type: 'number' }
      },
      additionalProperties: false
    }
  },
  {
    name: 'wps_get_audit_record',
    description: '按审计 ID 读取完整记录和快照',
    inputSchema: { type: 'object', properties: { auditId: { type: 'string' } }, required: ['auditId'], additionalProperties: false }
  },
  {
    name: 'wps_clear_audit_history',
    description: '清空本地存储的全部修改记录留痕',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false }
  }
];

export function auditTools(): ToolDefinition[] {
  return RAW.map(tool => ({ ...tool, inputSchema: structuredClone(tool.inputSchema) }));
}
