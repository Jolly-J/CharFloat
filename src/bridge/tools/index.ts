/**
 * 工具清单装配（P2.3 逐类迁移的装配点）。
 *
 * 顺序**影响 tools/list 与契约快照**，必须与迁移前一致：
 *   诊断 → 统一 excel 入口 → 网关原生工具（wps_ 与 office_ 前缀）→ 审计
 *
 * 依赖方向：本模块位于运行时注册层，引用契约层与 `definitions/`（纯数据定义）；
 * **不引用 catalog，也不反向 import 网关门面**——`UniversalGateway.getOpenAiTools()`
 * 与本模块调用的是同一个 `gatewayToolDefinitions()`，网关只是兼容转发。
 */
import { gatewayToolDefinitions } from './definitions/index.js';
import { isReadOnlyTool } from '../contracts/host-methods.js';
import { diagnosticTools, type ToolDefinition } from './diagnostics.js';
import { auditTools } from './audit.js';
import { unifiedExcelTools, type GatewayTool } from './excel.js';

/** 网关原生工具，做一次描述清洗并深拷贝 schema，避免调用方改动污染定义。 */
function gatewayTools(): GatewayTool[] {
  return gatewayToolDefinitions().map(tool => ({
    name: tool.function.name,
    description: tool.function.description
      .replace(/【[^】]*】/g, '')
      .replace(/实现 100% 任意操作无死角！|支持任意生僻 API 与长尾操作！/g, ''),
    inputSchema: structuredClone(tool.function.parameters)
  }));
}

/**
 * 工具分类（P2.3 逐类迁移的显式标签）。
 *
 * 分类驱动后续检查：宿主路由（Excel 类才走 `callOffice`）、只读判定、以及 P2.5 的
 * 宿主分支/参数转换/返回值校验。新增工具必须能落到某一类，否则测试会失败。
 */
export type ToolClass = 'diagnostic' | 'excel' | 'word' | 'ppt' | 'script' | 'lock' | 'audit' | 'microsoft';

export function toolClassOf(name: string): ToolClass {
  if (name.startsWith('bridge_')) return 'diagnostic';
  // 审计类：审计查询/清理，以及由审计记录驱动的回滚
  if (name === 'wps_rollback' || /^wps_(get_|clear_)?audit/.test(name)) return 'audit';
  // 目标锁：WPS 与 Microsoft 各一套，均为本地内存状态
  if (name === 'wps_get_locked_status' || /^(wps|office)_(lock|unlock)_/.test(name)) return 'lock';
  // 原生脚本与 API 反射：可跨组件，不属单表格工具
  if (name === 'wps_execute_script' || name === 'wps_inspect_api') return 'script';
  // 加载项生命周期：让刚部署的构建免重启 WPS 生效，属平台操作而非表格工具
  if (name === 'wps_reload_addon') return 'script';
  if (name.startsWith('wps_word_')) return 'word';
  if (name.startsWith('wps_ppt_')) return 'ppt';
  if (name.startsWith('office_')) return 'microsoft';
  // 统一入口与剩余 wps_ 兼容名均为表格工具
  if (name.startsWith('excel_') || name.startsWith('wps_')) return 'excel';
  throw new Error(`工具 ${name} 无法归类，请在 toolClassOf 中补充分类`);
}

export type RegisteredTool = ToolDefinition & {
  toolClass: ToolClass;
  annotations: { readOnlyHint: boolean; openWorldHint: boolean };
};

/** 组装完整对外工具清单；`readOnlyHint` 统一由契约层判定。 */
export function assembleTools(): RegisteredTool[] {
  const existing = gatewayTools();
  const ordered: ToolDefinition[] = [
    ...diagnosticTools(),
    ...unifiedExcelTools(existing),
    ...existing,
    ...auditTools()
  ];
  return ordered.map(tool => ({
    ...tool,
    toolClass: toolClassOf(tool.name),
    annotations: { readOnlyHint: isReadOnlyTool(tool.name), openWorldHint: false }
  }));
}

export { type ToolDefinition };
