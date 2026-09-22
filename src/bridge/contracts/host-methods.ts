/**
 * 宿主方法契约：方法标识、路由表、宿主支持缺口与只读/重放属性。
 *
 * 位置约束（改造计划 §3、P2.1）：本模块属于**契约层**，只依赖 Node 内置能力，
 * 不引用 gateway、catalog、adapter、ws-server 或任何服务实例。
 * `office/adapter.ts` 从这里取宿主方法标识，从而**不再反向依赖 catalog**，
 * 消除 catalog → gateway → adapter → catalog 的依赖回环。
 */

export type HostName = 'wps' | 'microsoft';

/**
 * 宿主方法路由表，同时是 Microsoft Office.js → Windows 原生 COM 的回退允许表。
 *
 * 注意：这里列的是**宿主方法名**，不等于都能被 AI 直接调用；可调用性由
 * `capabilities(). hosts.*.excel.implemented` 如实反映，声明了但没有可调用工具的项
 * 单列在 `declaredNotCallable`。
 */
export const EXCEL_METHODS = [
  'get_workspace_summary', 'get_sheet_outline', 'read_range', 'get_range_styles', 'search_cells',
  'create_sheet', 'delete_sheet', 'clear_range', 'patch_cells', 'format_cells', 'add_conditional_formatting',
  'freeze_panes', 'modify_rows_columns', 'auto_fit_columns', 'insert_dimension', 'get_charts', 'add_chart',
  'update_chart', 'delete_chart', 'create_pivot_table', 'set_filter_and_sort', 'set_data_validation', 'manage_sheet',
  'manage_rows_and_columns', 'manage_cell_comments', 'find_and_replace', 'duplicate_sheet', 'capture_sheet_preview',
  'rollback_cells', 'save_workbook', 'get_style_token', 'format_text_segment', 'configure_print_layout', 'export_sheet_pdf',
  'set_sheet_view', 'export_chart_image', 'copy_range', 'manage_hyperlink', 'manage_named_range',
  'manage_document_properties', 'manage_table', 'manage_pictures',
  'add_shape', 'list_shapes', 'update_shape', 'group_shapes', 'ungroup_shapes', 'set_shape_zorder', 'export_shape_image'
] as const;

/**
 * 宿主侧实现缺口：声明了路由、但对应加载项里没有 RPC 分支的方法。
 *
 * 依据（静态核对，未实机复现）：
 * - `wps-addon/addon-core.js` 的 RPC 分发 switch 只有 add_chart / get_charts / delete_chart，没有 update_chart；
 *   未识别方法会走到 `default:` 抛出"未知的 RPC 方法"。因此 `excel_update_chart`（host=wps）必然失败。
 * - `office-addon/public/taskpane.js` 的 dispatchExcelTool 有 update_chart，故 Microsoft 侧无此缺口。
 */
export const HOST_IMPLEMENTATION_GAPS: Record<HostName, readonly string[]> = {
  // CAP-07 后 WPS 表格侧已实现 add_shape/list_shapes/update_shape/group_shapes/
  // ungroup_shapes/set_shape_zorder/update_chart（CAP-21 已补 WPS 侧实现）；
  wps: ['export_shape_image'],
  // get_style_token 是 WPS 宿主独有的"读原表设计语言"能力，Office.js 的 dispatchExcelTool 无此分支；
  // 声明在此，使 excel_get_style_token(host=microsoft) 的失败被正确归类为"宿主未实现"而非未知工具。
  // format_text_segment 同为 WPS 独有（Office.js 无字符级富文本入口）
  // set_sheet_view（网格线/行列标题/缩放）是 WPS 表格专有的视图能力，
  // Office.js 没有等价的视图 API，因此对 Microsoft 声明为未实现。
  // CAP-15~20 第二类能力（区域复制/超链接/命名区域/文档属性/结构化表格/图片）目前只实现了 WPS 侧，
  // Office.js 侧尚未实现，故对 Microsoft 声明为未实现。
  microsoft: ['get_style_token', 'format_text_segment', 'configure_print_layout', 'export_sheet_pdf', 'set_sheet_view',
    'copy_range', 'manage_hyperlink', 'manage_named_range', 'manage_document_properties', 'manage_table', 'manage_pictures',
    'export_chart_image']
};

/**
 * Windows 原生 COM 通道（`resources/office/excel.ps1`）**实际实现**的方法缺口。
 *
 * ISS-97：`office/adapter.ts` 曾把整张 `EXCEL_METHODS`(30) 当作 COM 回退白名单，
 * 而 COM 只覆盖 28/30 —— `update_chart` 与 `save_workbook` 在 `excel.ps1` 的
 * `switch($method)` 里没有分支（`save_workbook` 连 case 都不存在），
 * 于是这两个方法会被判为"可回退"，回退后却抛 `Unsupported Excel method`，
 * 白名单成了过度声明。这里按实际覆盖生成，回退判定只认这张表。
 */
export const COM_IMPLEMENTATION_GAPS: readonly string[] = ['update_chart', 'save_workbook'];

/** COM 通道真实可执行的方法集（按实际实现生成，不是整张路由表）。 */
export const COM_EXCEL_METHODS: readonly string[] = EXCEL_METHODS.filter(
  method => !COM_IMPLEMENTATION_GAPS.includes(method)
);

/**
 * 不修改宿主文档内容的宿主方法。
 *
 * 只有这些方法允许在前一通道结果不可判定时改走另一通道重放；未列出的方法一律按"可能修改文档"处理。
 * 注意：本集合服务于「重放安全性」，与 MCP 的 readOnlyHint 不是同一概念 ——
 * 例如截图类方法会写本地文件（因此不是 readOnly），但不会改动文档内容（因此可以安全重放）。
 *
 * `inspect_api` **不在此列**：它接收的是任意表达式而非受限读取，表达式本身可以带副作用
 * （例如取属性时触发宿主求值、调用会改文档的方法、创建或删除对象）。因此既不能算只读，也不能重放。
 */
export const REPLAY_SAFE_METHODS: ReadonlySet<string> = new Set([
  'get_workspace_summary', 'get_sheet_outline', 'get_style_token', 'read_range', 'get_range_styles', 'search_cells',
  'get_charts', 'word_read_document', 'ppt_read_presentation', 'ppt_get_slide_shapes'
]);

/**
 * MCP readOnlyHint 使用的只读工具集合（不含 wps_/excel_ 前缀）。
 * 只列不会改动文档内容、也不会产生持久副作用的工具；写入与截图导出不在其中。
 *
 * `inspect_api` **不在此列**：它执行调用方给出的任意表达式，无法静态保证只读（见上）。
 */
export const MCP_READ_ONLY_TOOLS: ReadonlySet<string> = new Set([
  'get_workspace_summary', 'get_sheet_outline', 'get_style_token', 'read_range', 'get_range_styles', 'search_cells',
  'get_charts', 'word_read_document', 'ppt_read_presentation', 'ppt_get_slide_shapes',
  'get_locked_status', 'get_audit_history', 'get_audit_record'
]);

/** 去掉 wps_/excel_ 兼容前缀，得到宿主机层方法名或裸工具名。 */
export function bareName(name: string): string {
  return String(name).replace(/^(wps|excel)[._]/, '');
}

export function isReplaySafeMethod(method: string): boolean {
  return REPLAY_SAFE_METHODS.has(bareName(method));
}

/** 判断一个对外工具是否应标记 readOnlyHint。 */
export function isReadOnlyTool(toolName: string): boolean {
  if (toolName.startsWith('bridge_')) return true;
  if (toolName === 'office_get_status') return true;
  return MCP_READ_ONLY_TOOLS.has(bareName(toolName));
}

/** 该宿主方法是否存在已知的宿主实现缺口。 */
export function isUnimplementedOn(host: HostName, method: string): boolean {
  return HOST_IMPLEMENTATION_GAPS[host].includes(bareName(method));
}
