/**
 * 诊断类工具定义（P2.3 逐类迁移）。
 *
 * 契约：只读。`bridge_diagnose` 额外返回逐组件检查结果，但不修改配置或文档。
 * 顺序：`bridge_get_capabilities` 在 `bridge_diagnose` 之前，与迁移前一致（影响 tools/list 与契约快照）。
 * 入口名 `bridge_get_capabilities`、`bridge_diagnose` 为已发布入口，不得改名。
 */
import { isReadOnlyTool } from '../contracts/host-methods.js';

export interface ToolDefinition {
  name: string;
  description: string;
  inputSchema: unknown;
}

const RAW: ToolDefinition[] = [
  {
    name: 'bridge_get_capabilities',
    description: '接入后首先调用：获取平台、真实连接状态、支持工具、验证程度、回滚限制和运行方式。',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false }
  },
  {
    name: 'bridge_diagnose',
    description: '检查 Bridge、WPS 连接和能力边界，不修改配置或文档。Office 状态可另用 office_get_status 探测。',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false }
  }
];

/** 按既有顺序返回诊断工具（顺序影响 tools/list 与契约快照，不可随意调整）。 */
export function diagnosticTools(): ToolDefinition[] {
  return RAW.map(tool => ({ ...tool }));
}

/** 该工具是否应标记 readOnlyHint；统一走契约层，避免各处各写一份。 */
export function readOnlyOf(name: string): boolean {
  return isReadOnlyTool(name);
}
