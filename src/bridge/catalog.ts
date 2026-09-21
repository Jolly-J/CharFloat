import { UniversalGateway } from './gateway.js';
import { currentHost } from './context.js';
import { bridgeServer } from './ws-server.js';
import { VERSION } from './runtime.js';

export const EXCEL_METHODS = [
  'get_workspace_summary', 'get_sheet_outline', 'read_range', 'get_range_styles', 'search_cells',
  'create_sheet', 'delete_sheet', 'clear_range', 'patch_cells', 'format_cells', 'add_conditional_formatting',
  'freeze_panes', 'modify_rows_columns', 'auto_fit_columns', 'insert_dimension', 'get_charts', 'add_chart',
  'update_chart', 'delete_chart', 'create_pivot_table', 'set_filter_and_sort', 'set_data_validation', 'manage_sheet',
  'manage_rows_and_columns', 'manage_cell_comments', 'find_and_replace', 'duplicate_sheet', 'capture_sheet_preview',
  'rollback_cells', 'save_workbook'
] as const;
export const READ_TOOLS = new Set(['get_workspace_summary', 'get_sheet_outline', 'read_range', 'get_range_styles', 'search_cells', 'get_charts', 'get_audit_history', 'get_audit_record']);
export function capabilities() {
  return {
    version: VERSION, platform: process.platform, preferredWorkflow: '在已打开的 Excel / WPS 表格中操作，先读后写，明确目标并核验结果',
    connection: bridgeServer.getState(),
    hosts: {
      wps: { transport: 'authenticated-addon', excel: { implemented: [...EXCEL_METHODS], validation: process.platform === 'darwin' ? '既有功能曾在 macOS 使用；2.0 回归需实机确认' : 'Windows 实机待验收' }, word: '有限支持：读取、编辑、保存；按实际工具清单使用', ppt: '有限支持：基础页面、文字、形状、表格；复杂图表与母版需实机验证，不保证高保真复刻' },
      microsoft: { transport: 'Office.js Web Add-in / WebSocket', excel: { implemented: [...EXCEL_METHODS], validation: '全量 42 项结构化能力支持 (macOS / Windows 统一)' }, word: 'Office.js 原生通道', ppt: 'Office.js 原生通道' }

    },
    audit: { covered: ['patch_cells 值与公式'], uncovered: ['样式', '图表', '结构修改', 'Word/PPT', '原生脚本'], rollback: '检查当前值和公式与记录的修改后快照一致；有后续修改则拒绝覆盖' },
    lifecycle: '后台进程独立于窗口。stdio 自动启动或复用；HTTP 客户端需先启动 Bridge。停止服务会断开所有客户端。',
    repairs: ['启动后台', '检查版本与端口', '重新部署本项目加载项', '仅合并用户选择的客户端配置'],
    limitations: ['不承诺未实测的平台可用', '写入超时结果未知，先读回，禁止自动重放', 'MCP 已连接不等于办公软件已连接', '持久保存必须有明确保存结果']
  };
}
export function getTools() {
  const existing = UniversalGateway.getOpenAiTools().map(t => ({ name: t.function.name, description: t.function.description.replace(/【[^】]*】/g, '').replace(/实现 100% 任意操作无死角！|支持任意生僻 API 与长尾操作！/g, ''), inputSchema: structuredClone(t.function.parameters) as any }));
  // Legacy wps_* calls always address WPS. Unified excel_* calls explicitly choose the host.
  const excel = existing.filter(t => EXCEL_METHODS.includes(t.name.replace(/^wps_/, '') as any)).map(t => ({
    ...t, name: t.name.replace(/^wps_/, 'excel_'), description: `${t.description} 统一表格入口；host 必填。Windows Microsoft Excel 尚待实机验收。`,
    inputSchema: { ...t.inputSchema, properties: { ...t.inputSchema.properties, host: { type: 'string', enum: ['wps', 'microsoft'], description: '明确选择 WPS 表格或 Microsoft Excel' } }, required: [...(t.inputSchema.required || []), 'host'] }
  }));
  const diagnostic = ['bridge_get_capabilities', 'bridge_diagnose'].map(name => ({ name, description: name === 'bridge_diagnose' ? '检查 Bridge、WPS 连接和能力边界，不修改配置或文档。Office 状态可另用 office_get_status 探测。' : '接入后首先调用：获取平台、真实连接状态、支持工具、验证程度、回滚限制和运行方式。', inputSchema: { type: 'object', properties: {}, additionalProperties: false } }));
  const audits = [
    { name: 'wps_get_audit_history', description: '分页筛选修改记录；默认 5 条摘要', inputSchema: { type: 'object', properties: { limit: { type: 'number' }, offset: { type: 'number' }, view: { type: 'string', enum: ['ids', 'summary', 'detail'] }, workbookName: { type: 'string' }, sheetName: { type: 'string' }, clientName: { type: 'string' }, actionType: { type: 'string' }, status: { type: 'string', enum: ['applied', 'rolled_back'] }, fromTimestamp: { type: 'number' }, toTimestamp: { type: 'number' } }, additionalProperties: false } },
    { name: 'wps_get_audit_record', description: '按审计 ID 读取完整记录和快照', inputSchema: { type: 'object', properties: { auditId: { type: 'string' } }, required: ['auditId'], additionalProperties: false } },
    { name: 'wps_clear_audit_history', description: '清空本地存储的全部修改记录留痕', inputSchema: { type: 'object', properties: {}, additionalProperties: false } }
  ];
  return [...diagnostic, ...excel, ...existing, ...audits].map(t => ({ ...t, annotations: { readOnlyHint: t.name.startsWith('bridge_') || READ_TOOLS.has(t.name.replace(/^(wps|excel)_/, '')) || t.name === 'office_get_status', openWorldHint: false } }));
}
export function validateArgs(schema: any, value: any, location = 'arguments'): void {
  if (!schema) return;
  if (schema.oneOf) {
    const matches = schema.oneOf.filter((part: any) => { try { validateArgs(part, value, location); return true; } catch { return false; } });
    if (matches.length !== 1) throw new Error(`${location}: 类型不匹配`);
  }
  if (schema.enum && !schema.enum.includes(value)) throw new Error(`${location}: 不在允许值中`);
  const types = Array.isArray(schema.type) ? schema.type : [schema.type];
  if (schema.type && !types.some((type: string) => type === 'null' ? value === null : type === 'array' ? Array.isArray(value) : type === 'object' ? value !== null && typeof value === 'object' && !Array.isArray(value) : type === 'integer' ? Number.isInteger(value) : typeof value === type)) throw new Error(`${location}: 类型不正确`);
  if (typeof value === 'number' && !Number.isFinite(value)) throw new Error(`${location}: 数字无效`);
  if (Array.isArray(value)) value.forEach((v, i) => validateArgs(schema.items, v, `${location}[${i}]`));
  else if (value && typeof value === 'object') {
    for (const key of schema.required || []) if (value[key] === undefined) throw new Error(`${location}.${key}: 必填`);
    for (const key of Object.keys(value)) {
      if (schema.additionalProperties === false && !schema.properties?.[key]) throw new Error(`${location}.${key}: 未知参数`);
      validateArgs(schema.properties?.[key], value[key], `${location}.${key}`);
    }
  }
}
const executionQueues = new Map<string, Promise<any>>();
export async function executeCatalogTool(name: string, args: any, clientName: string) {
  const key = 'office';
  const job = (executionQueues.get(key) || Promise.resolve()).catch(() => {}).then(() => executeValidatedTool(name, args, clientName));
  executionQueues.set(key, job);
  try { return await job; } finally { if (executionQueues.get(key) === job) executionQueues.delete(key); }
}
async function executeValidatedTool(name: string, args: any, clientName: string) {
  const tool = getTools().find(t => t.name === name);
  if (!tool) throw new Error(`未知工具 ${name}`);
  validateArgs(tool.inputSchema, args);
  if (name === 'bridge_get_capabilities') return capabilities();
  if (name === 'bridge_diagnose') return { ...capabilities(), checks: [{ id: 'service', status: 'ok' }, ...Object.entries(bridgeServer.getState().components).map(([id, s]) => ({ id: `wps-${id}`, status: s.connected ? 'connected' : 'not-connected', next: s.connected ? '先读取目标文档' : '打开目标办公组件，并检查加载项是否安装、版本是否匹配' }))] };
  const realName = name.replace(/^excel_/, 'wps_');
  args = structuredClone(args);
  if (realName === 'wps_patch_cells') {
    for (const field of ['values', 'formulas']) if (args[field] != null && (!Array.isArray(args[field]) || args[field].length === 0 || args[field].some((row: any) => !Array.isArray(row) || row.length !== args[field][0].length))) throw new Error(`${field} 必须为非空矩形二维数组`);
  }
  const result = await UniversalGateway.executeTool(realName, args, clientName);
  if (result?.success === false) throw new Error(result.error || result.message || '办公软件未完成请求');
  return result;
}
