import type { WorkspaceSummary, SheetOutline, PatchResult } from '../types.js';
import type { ChannelParams, ToolArgs } from '../contracts/boundary.js';

export interface NormalizedRequest {
  method: string;
  params: ChannelParams;
}

/**
 * 图表类型映射表 (从通用下划线/小写映射至 Office.js PascalCase 规范)
 */
const CHART_TYPE_MAP: Record<string, string> = {
  column: 'ColumnClustered',
  column_clustered: 'ColumnClustered',
  columnclustered: 'ColumnClustered',
  bar: 'BarClustered',
  bar_clustered: 'BarClustered',
  barclustered: 'BarClustered',
  line: 'Line',
  pie: 'Pie',
  doughnut: 'Doughnut',
  area: 'Area',
  scatter: 'Scatter'
};

/**
 * Microsoft 通道明确无法实现的字段：**显式报错**，不再原样透传后被加载项静默忽略。
 *
 * 背景（ISS-93 / M1–M8）：schema、gateway、normalizer 三层按 WPS 语义书写，Office.js
 * 加载项读的是另一套字段名；映射不了的参数原样透传时，加载项既不报错也不生效，
 * 调用方拿到 `success:true` 却什么都没发生。这里对"决定操作效果"的字段直接抛错，
 * 让失败变得可判定（抛错属执行前拒绝，Windows 上仍可回退原生 COM）。
 */
function msUnsupported(feature: string, alternative: string): Error {
  return new Error(
    `Microsoft Excel 通道未实现 ${feature}：为避免"返回成功但实际未生效"，已在调用宿主前拒绝本次请求。` +
    `替代路径：${alternative}`
  );
}

/** 从 schema 参数里挑出 Microsoft 通道无法实现、但不阻断主效果的字段，用于在响应里如实暴露。 */
function reportUnsupportedFields(method: string, params: ToolArgs = {}): string[] {
  const p: any = params || {};
  const out: string[] = [];
  const push = (...fields: string[]) => { for (const f of fields) if (!out.includes(f)) out.push(f); };
  switch (method) {
    case 'set_filter_and_sort':
      // Office.js 侧没有"只开筛选按钮"的路径：enableAutoFilter 无法下发，如实回报。
      if (p.enableAutoFilter === true) push('enableAutoFilter');
      break;
    case 'duplicate_sheet':
      // Office.js 的 copy 分支固定复制到源表之后（positionType 未被读取），如实回报。
      if (p.position && String(p.position).toLowerCase() !== 'after') push('position');
      break;
    default:
      break;
  }
  return out;
}

/**
 * 将 Gateway 标准工具调用请求正规化为 Office.js 加载项期望的契约 (统一参数适配管道)
 */
export function normalizeOfficeRequest(method: string, params: ToolArgs = {}): NormalizedRequest {
  // 边界签名已收敛为 ToolArgs；函数内部的规范化工作副本按 P2.4 的范围要求保留宽松类型，
  // 避免为了消除 any 而重写全部字段访问（那属于大规模类型重写）。
  const p: any = { ...params };

  switch (method) {
    case 'wps_get_workspace_summary':
    case 'excel_get_workspace_summary':
    case 'get_workspace_summary':
      return { method: 'get_workspace_summary', params: p };

    case 'wps_get_sheet_outline':
    case 'excel_get_sheet_outline':
    case 'get_sheet_outline':
      return { method: 'get_sheet_outline', params: p };

    case 'wps_read_range':
    case 'excel_read_range':
    case 'read_range':
      return { method: 'read_range', params: p };

    case 'wps_patch_cells':
    case 'excel_patch_cells':
    case 'patch_cells':
      return {
        method: 'patch_cells',
        params: {
          ...p,
          sheetName: p.sheetName,
          address: p.address,
          values: p.values,
          formulas: p.formulas,
          workbookName: p.workbookName
        }
      };

    case 'wps_format_cells':
    case 'excel_format_cells':
    case 'format_cells':
    case 'format_range': {
      let normalizedBorders = p.borders;
      if (p.borders === true) {
        normalizedBorders = '#CBD5E1';
      } else if (p.borders === false || p.borders === 'none') {
        normalizedBorders = null;
      }

      return {
        method: 'format_cells',
        params: {
          ...p,
          sheetName: p.sheetName,
          address: p.address,
          workbookName: p.workbookName,
          fontName: p.fontName,
          fontSize: p.fontSize,
          bold: p.bold,
          fontColor: p.fontColor,
          backgroundColor: p.backgroundColor,
          horizontalAlignment: p.horizontalAlignment,
          verticalAlignment: p.verticalAlignment,
          numberFormat: p.numberFormat,
          rowHeight: p.rowHeight,
          columnWidth: p.columnWidth,
          wrapText: p.wrapText,
          borders: normalizedBorders,
          merge: p.merge,
          unmerge: p.unmerge,
          font: {
            name: p.fontName || p.font?.name,
            size: p.fontSize || p.font?.size,
            bold: p.bold !== undefined ? p.bold : p.font?.bold,
            color: p.fontColor || p.font?.color
          },
          fill: {
            color: p.backgroundColor || p.fill?.color
          },
          alignment: {
            horizontal: p.horizontalAlignment || p.alignment?.horizontal,
            vertical: p.verticalAlignment || p.alignment?.vertical,
            wrapText: p.wrapText !== undefined ? p.wrapText : p.alignment?.wrapText
          }
        }
      };
    }

    case 'wps_auto_fit_columns':
    case 'excel_auto_fit_columns':
    case 'auto_fit_columns':
      return {
        method: 'auto_fit_columns',
        params: {
          ...p,
          sheetName: p.sheetName,
          address: p.address || p.range,
          columnRules: p.columnRules,
          workbookName: p.workbookName
        }
      };

    case 'wps_add_conditional_formatting':
    case 'excel_add_conditional_formatting':
    case 'add_conditional_formatting':
      return { method: 'add_conditional_formatting', params: p };

    case 'wps_freeze_panes':
    case 'excel_freeze_panes':
    case 'freeze_panes':
      return { method: 'freeze_panes', params: p };

    case 'wps_modify_rows_columns':
    case 'excel_modify_rows_columns':
    case 'modify_rows_columns':
    case 'wps_manage_rows_and_columns':
    case 'excel_manage_rows_and_columns':
    case 'manage_rows_and_columns': {
      // ISS-93：schema 用 targetType('row'|'column')，早期加载项读的是 dimension('rows'|'columns')
      // 且默认 rows → 想插列却插行。bridge 侧把两者都给出（Office.js 侧现在也两套都认，且识别不了时显式报错）。
      // hide/unhide/set_size 在 Microsoft 侧已实现（office-addon/src/excel/range.js），因此**不在这里阻断**。
      if (p.targetType === undefined && p.dimension === undefined) {
        return { method: 'modify_rows_columns', params: p };
      }
      const dimension = p.targetType === 'column' ? 'columns' : p.targetType === 'row' ? 'rows' : p.dimension;
      const kind = dimension === 'columns' ? 'column' : 'row';
      return {
        method: 'modify_rows_columns',
        params: { ...p, targetType: p.targetType ?? kind, dimension, count: p.count ?? 1 }
      };
    }

    case 'wps_get_charts':
    case 'excel_get_charts':
    case 'get_charts':
      // 选择器（shapeName/chartIndex/chartTitle）由响应侧按 schema 语义过滤，见 normalizeOfficeResponse。
      return { method: 'get_charts', params: p };

    case 'wps_add_chart':
    case 'excel_add_chart':
    case 'add_chart': {
      // ISS-93（M8）：早期加载项只读 seriesBy / seriesColors / hasLegend / title / position.*，
      // yAxis / seriesSettings / smoothLine / dataRanges / hasDataLabels 被静默忽略。
      // Office.js 侧现已逐项实现并**读回校验**（未生效写进返回体 warnings），因此桥接侧只做数据源与
      // 定位字段归一，不再阻断这些参数。
      const dataRanges = Array.isArray(p.dataRanges) ? p.dataRanges.filter(Boolean) : [];
      const dataRange = p.dataRange || p.sourceAddress || p.sourceRange || p.range || (dataRanges.length === 1 ? dataRanges[0] : undefined);
      const rawChartType = String(p.chartType || 'column_clustered').toLowerCase().replace(/-/g, '_');
      const chartType = CHART_TYPE_MAP[rawChartType] || p.chartType || 'ColumnClustered';

      return {
        method: 'add_chart',
        params: {
          ...p,
          sourceAddress: dataRange,
          dataRange,
          chartType,
          left: p.left ?? p.position?.left,
          top: p.top ?? p.position?.top,
          width: p.width ?? p.position?.width,
          height: p.height ?? p.position?.height
        }
      };
    }

    case 'wps_delete_chart':
    case 'excel_delete_chart':
    case 'delete_chart':
      return {
        method: 'delete_chart',
        params: {
          ...p,
          clearAll: Boolean(p.clearAll),
          chartName: p.chartName || p.name || p.shapeName || p.id
        }
      };

    case 'wps_search_cells':
    case 'excel_search_cells':
    case 'search_cells':
    case 'find_replace':
      return {
        method: 'find_replace',
        params: {
          ...p,
          text: p.text || p.query || p.keyword || p.findText || '',
          matchCase: !!p.matchCase,
          matchEntireCell: !!p.matchEntireCell
        }
      };

    case 'wps_save_workbook':
    case 'excel_save_workbook':
    case 'save_workbook':
    case 'save':
      return { method: 'save', params: p };

    // ---- ISS-93：以下方法原来走 default 原样透传，字段名与加载项不一致 ----

    case 'wps_create_sheet':
    case 'excel_create_sheet':
    case 'create_sheet':
      // 加载项读 name || sheetName。
      return { method: 'create_sheet', params: { ...p, name: p.name || p.sheetName } };

    case 'wps_delete_sheet':
    case 'excel_delete_sheet':
    case 'delete_sheet':
      // 加载项读 sheetName || name。
      return { method: 'delete_sheet', params: { ...p, sheetName: p.sheetName || p.name } };

    case 'wps_duplicate_sheet':
    case 'excel_duplicate_sheet':
    case 'duplicate_sheet': {
      // 加载项读 sourceSheet / newName / positionType，schema 给的是 sourceSheetName / newSheetName / position。
      const source = p.sourceSheetName || p.sheetName || p.sourceSheet;
      const hasIntent = Boolean(source || p.newSheetName || p.newName);
      if (!hasIntent) return { method: 'duplicate_sheet', params: p };
      if (!source) throw msUnsupported('缺少 sourceSheetName 的工作表克隆', '请传入要克隆的源工作表名 sourceSheetName。');
      if (!p.newSheetName && !p.newName) throw msUnsupported('缺少 newSheetName 的工作表克隆', '请传入新工作表名 newSheetName。');
      const position = String(p.position || 'after').toLowerCase();
      const positionType = position === 'before' ? 'Before' : position === 'end' ? 'End' : 'After';
      return {
        method: 'duplicate_sheet',
        params: { ...p, sheetName: source, sourceSheet: source, newName: p.newSheetName || p.newName, positionType }
      };
    }

    case 'wps_set_data_validation':
    case 'excel_set_data_validation':
    case 'set_data_validation':
      // ISS-93（M1）：早期加载项只认 params.rule → 先 clear() 再什么都不设（静默清空既有校验）。
      // Office.js 侧现在直接用 schema 字段（validationType / listItems / operator / minVal / maxVal /
      // prompt* / error*）构造规则，并在构造失败时**在 clear() 之前**报错；桥接侧不再自建 rule，
      // 避免两处各维护一份映射（遵循"不重复维护能力集合"）。
      return { method: 'set_data_validation', params: p };

    case 'wps_set_filter_and_sort':
    case 'excel_set_filter_and_sort':
    case 'set_filter_and_sort': {
      // 加载项读 address / sortColumn / sortOrder / hasHeader；schema 给的是 range / sortRules。
      const address = p.range || p.address;
      if (!address) return { method: 'set_filter_and_sort', params: p };
      const rules = Array.isArray(p.sortRules) ? p.sortRules.filter((r: any) => r && r.colIndex !== undefined) : [];
      if (rules.length > 1) {
        throw msUnsupported(
          '多级排序（sortRules 多于 1 项）',
          'Microsoft 通道的加载项只应用单列排序；请只保留主排序列，或改用 host=wps。'
        );
      }
      const params: any = { ...p, address, hasHeader: true };
      // enableAutoFilter 无法只开筛选按钮而不施加筛选条件（加载项的 filter 分支必须给 criteria），
      // 这里不下发该字段，改由响应体的 unsupportedFields 如实说明。
      delete params.enableAutoFilter;
      if (!rules.length) {
        // 没有排序规则时，加载项会退回"按第 1 列升序排序"——那是调用方没要求的写操作，必须拒绝（ISS-93）。
        throw msUnsupported(
          '仅切换自动筛选（未提供 sortRules）',
          'Microsoft 通道的加载项只有"按列排序"与"按条件筛选"两条路径，无法只开/关筛选按钮；请改用 host=wps。'
        );
      }
      {
        const colIndex = Number(rules[0].colIndex);
        if (!Number.isFinite(colIndex) || colIndex < 1) throw msUnsupported(`sortRules[0].colIndex=${rules[0].colIndex}`, 'colIndex 必须是 ≥1 的数字（1=区域第 1 列）。');
        params.action = 'sort';
        // Office.js 的 SortField.key 是区域内 0 基列号，schema 的 colIndex 是 1 基。
        params.sortColumn = colIndex - 1;
        params.sortOrder = rules[0].order === 'desc' ? 'desc' : 'asc';
      }
      return { method: 'set_filter_and_sort', params };
    }

    case 'wps_manage_sheet':
    case 'excel_manage_sheet':
    case 'manage_sheet': {
      // ISS-93（M6）：早期加载项只识别 copy/duplicate/rename/hide/show/color，且读 params.tabColor，
      // 导致 move/tab_color/protect/unprotect 静默 no-op。Office.js 侧现在同时识别两套语义，
      // 并对 protect/unprotect（MS 桌面版不可靠）显式报错；这里只补齐字段别名，让新旧实现都能正确落位。
      const params: any = { ...p };
      if (!params.newName && params.newSheetName) params.newName = params.newSheetName;
      if (params.action === 'tab_color' && !params.tabColor && params.color) params.tabColor = params.color;
      return { method: 'manage_sheet', params };
    }

    case 'wps_create_pivot_table':
    case 'excel_create_pivot_table':
    case 'create_pivot_table': {
      // 加载项读 sourceAddress / destinationAddress / name；schema 给的是 sourceRange / destCell。
      const sourceAddress = p.sourceRange || p.sourceAddress;
      const destinationAddress = p.destCell || p.destinationAddress;
      const hasIntent = Boolean(sourceAddress || destinationAddress || p.rowFields || p.columnFields || p.dataFields);
      if (!hasIntent) return { method: 'create_pivot_table', params: p };
      if (!sourceAddress) throw msUnsupported('缺少 sourceRange 的数据透视表', '请传入源区域 sourceRange。');
      if (!destinationAddress) throw msUnsupported('缺少 destCell 的数据透视表', '请传入目标单元格 destCell。');
      const arranged = ['rowFields', 'columnFields', 'dataFields'].some(field => Array.isArray(p[field]) && p[field].length > 0);
      if (arranged) {
        throw msUnsupported(
          '透视表字段编排（rowFields / columnFields / dataFields）',
          'Microsoft 通道只能创建空透视表骨架，不会写入行列与聚合字段；请改用 host=wps，或在 MS 侧创建后用 Office.js 脚本自行编排。'
        );
      }
      const sourceSheetName = p.sourceSheetName;
      const destSheetName = p.destSheetName;
      if (sourceSheetName && destSheetName && sourceSheetName !== destSheetName) {
        throw msUnsupported('跨工作表的透视表来源', 'Microsoft 通道按目标工作表解析源区域；请让源表与目标表一致，或改用 host=wps。');
      }
      const sheetName = destSheetName || sourceSheetName || p.sheetName;
      return {
        method: 'create_pivot_table',
        params: { ...p, sheetName, sourceAddress, destinationAddress, name: p.name || 'PivotTable1' }
      };
    }

    case 'wps_manage_cell_comments':
    case 'excel_manage_cell_comments':
    case 'manage_cell_comments': {
      // 加载项 action 只识别 add / list / read / delete；schema 的 read 可用，但 delete 缺 id 入口。
      const action = String(p.action || '');
      if (action === 'clear_all') {
        throw msUnsupported('清空全表批注（clear_all）', 'Microsoft 通道只能逐条 delete；请改用 host=wps，或先 list 取 id 再逐条删除。');
      }
      if (action === 'delete') {
        throw msUnsupported('按单元格删除批注（delete 无 id 入口）', 'Microsoft 通道删除批注需要批注 id，当前 schema 未提供；请改用 host=wps 或 host=wps 的 clear_all。');
      }
      if (p.author !== undefined && p.author !== null && p.author !== '') {
        throw msUnsupported('设置批注作者（author）', 'Office.js 的批注作者由宿主账户决定、不可写；请去掉 author，或改用 host=wps（WPS 侧作者同样恒为当前用户）。');
      }
      if (action === 'read') {
        return { method: 'manage_cell_comments', params: { ...p, action: 'list' } };
      }
      return { method: 'manage_cell_comments', params: p };
    }

    case 'wps_find_and_replace':
    case 'excel_find_and_replace':
    case 'find_and_replace':
    case 'find_replace_cells': {
      // 加载项读 text / searchQuery / address；schema 给的是 searchQuery / searchRange。
      const supplied = ['searchQuery', 'text', 'query', 'findText'].filter(k => (p as any)[k] !== undefined);
      const text = p.searchQuery ?? p.text ?? p.query ?? p.findText;
      if (supplied.length && (text === undefined || text === null || String(text) === '')) {
        throw msUnsupported('空搜索串的查找替换', '空串会命中整个区域并可能破坏内容，已阻断；请显式提供 searchQuery。');
      }
      return {
        method: 'find_and_replace',
        params: {
          ...p,
          text: text ?? '',
          address: p.searchRange || p.address,
          matchCase: !!p.matchCase,
          matchEntireCell: !!p.matchEntireCell
        }
      };
    }

    case 'wps_capture_sheet_preview':
    case 'excel_capture_sheet_preview':
    case 'capture_sheet_preview':
      // 加载项读 address / sheetName / chartName（其余定位参数在 MS 侧不可用）。
      return {
        method: 'capture_sheet_preview',
        params: { ...p, address: p.address || p.range, chartName: p.chartName || p.name }
      };

    default:
      return { method: method.replace(/^(wps_|excel_)/, ''), params: p };
  }
}

/**
 * 将 Office.js 加载项的原生响应标准化为 Gateway 与 MCP 契约
 */
/**
 * @param raw 宿主原始响应。形状由各宿主决定，保持 `any` 是刻意的：
 *            这里正是"宿主差异显式留在适配器"的落点，收窄会掩盖不同宿主的返回语义。
 *            返回类型同样保持宽松：各分支产出的是具名业务类型（WorkspaceSummary / PatchResult 等），
 *            强行套统一的 ToolResult 会要求它们都有索引签名，属大规模类型重写。
 */
export function normalizeOfficeResponse(method: string, raw: any, originalParams: ToolArgs = {}): any {
  if (!raw) return raw;

  const baseMethod = method.replace(/^(wps_|excel_)/, '');

  switch (baseMethod) {
    case 'get_workspace_summary':
      return {
        // 加载项只在工作簿已打开时存活，因此默认 true；显式返回 false 时以宿主的真实报告为准（ISS-02）。
        hasOpenWorkbook: raw.hasOpenWorkbook ?? true,
        workbookName: raw.workbookName || originalParams?.workbookName || '工作簿1.xlsx',
        fullName: raw.fullName || raw.workbookName || originalParams?.workbookName || '工作簿1.xlsx',
        activeSheetName: raw.activeSheetName || raw.activeSheet || (raw.sheets && raw.sheets[0]?.name) || 'Sheet1',
        sheetCount: raw.sheetCount || (Array.isArray(raw.sheets) ? raw.sheets.length : 1),
        sheets: (raw.sheets || []).map((s: any, idx: number) => ({
          name: typeof s === 'string' ? s : s.name,
          position: typeof s === 'object' && s.position ? s.position : idx + 1,
          visibility: typeof s === 'object' && s.visibility ? s.visibility : 'Visible'
        })),
        openWorkbooks: Array.isArray(raw.openWorkbooks) ? raw.openWorkbooks : [],
        selection: raw.selection || null,
        message: raw.message || '获取工作区摘要成功'
      } as WorkspaceSummary;

    case 'get_sheet_outline':
      return {
        sheetName: raw.sheetName || originalParams?.sheetName || 'Sheet1',
        isEmpty: raw.isEmpty ?? false,
        usedRangeAddress: raw.usedRangeAddress || raw.address || 'A1',
        startRow: raw.startRow || 1,
        startColumn: raw.startColumn || 1,
        rowCount: raw.rowCount || 0,
        columnCount: raw.columnCount || 0,
        headerPreview: raw.headerPreview || []
      } as SheetOutline;

    case 'patch_cells':
    case 'write_range': {
      // 边界签名是 ToolArgs（值形状由 schema 约束），此处按二维数组读取是已知契约。
      const values = originalParams.values as unknown[][] | undefined;
      const formulas = originalParams.formulas as unknown[][] | undefined;
      const rows = raw.rowCount || values?.length || formulas?.length || 1;
      const cols = raw.columnCount || values?.[0]?.length || formulas?.[0]?.length || 1;
      const count = raw.modifiedCount || (rows * cols);
      return {
        // ISS-99：先铺开宿主原始响应（warnings / 读回字段等），再覆盖桥接侧统一字段，
        // 避免白名单式重建把宿主新增的"如实提示"丢掉。
        ...raw,
        success: true,
        sheetName: raw.sheetName || originalParams?.sheetName || 'Sheet1',
        address: raw.address || originalParams?.address,
        workbookName: raw.workbookName || originalParams?.workbookName,
        modifiedCount: count,
        rowCount: rows,
        columnCount: cols,
        diff: raw.diff || [],
        message: raw.message || `已成功修改 ${count} 个单元格`
      };
    }

    case 'format_cells':
    case 'format_range':
      return {
        ...raw,
        success: true,
        sheetName: raw.sheetName || originalParams?.sheetName,
        address: raw.address || originalParams?.address,
        message: raw.message || `已成功设置区域 ${raw.address || originalParams?.address} 格式`
      };

    case 'auto_fit_columns':
      return {
        ...raw,
        success: true,
        message: raw.message || '已成功自适应调整列宽'
      };

    case 'add_chart': {
      const chartId = raw.id || raw.name;
      return {
        // ISS-99：warnings / dataRangesApplied / seriesCount / yAxisApplied 等如实提示必须原样传给调用方。
        ...raw,
        success: true,
        id: chartId,
        name: raw.name || chartId,
        shapeName: raw.shapeName || raw.name || chartId,
        chartIndex: raw.chartIndex || 1,
        title: raw.title || originalParams?.title || '',
        chartType: raw.chartType || originalParams?.chartType || 'column_clustered',
        dataRange: raw.dataRange || originalParams?.dataRange || originalParams?.sourceAddress,
        sheetName: raw.sheetName || originalParams?.sheetName || '',
        workbookName: raw.workbookName || originalParams?.workbookName || '工作簿1.xlsx',
        replaceExisting: raw.replaceExisting ?? (originalParams?.replaceExisting !== false),
        left: raw.left,
        top: raw.top,
        width: raw.width,
        height: raw.height,
        message: raw.message || `图表已成功创建 (ID: ${chartId})`
      };
    }

    case 'get_charts': {
      const all: any[] = Array.isArray(raw.charts) ? raw.charts : [];
      // ISS-93（M7）：早期加载项忽略 shapeName/chartIndex/chartTitle，会把整表图表全返回。
      // Office.js 侧现已自行过滤（返回 filtered/selector/totalCount）；只有宿主没做时才在桥接侧补，
      // 避免重复过滤把宿主如实回报的 totalCount 语义搞乱（ISS-99）。
      if (raw.filtered !== undefined || raw.selector !== undefined) {
        return { ...raw, success: true, count: raw.count ?? all.length, charts: all, detail: !!originalParams?.detail };
      }
      let charts = all;
      const wantIndex = (originalParams as any)?.chartIndex;
      const wantName = (originalParams as any)?.shapeName;
      const wantTitle = (originalParams as any)?.chartTitle;
      if (wantIndex !== undefined && wantIndex !== null && wantIndex !== '') {
        charts = charts.filter(c => Number(c?.chartIndex) === Number(wantIndex));
      }
      if (wantName) charts = charts.filter(c => c?.name === wantName || c?.shapeName === wantName);
      if (wantTitle) charts = charts.filter(c => String(c?.title ?? '').includes(String(wantTitle)));
      const filtered = charts.length !== all.length;
      return {
        success: true,
        count: charts.length,
        charts,
        detail: !!originalParams?.detail,
        ...(filtered ? { filteredFrom: all.length, filterApplied: { shapeName: wantName ?? null, chartIndex: wantIndex ?? null, chartTitle: wantTitle ?? null } } : {})
      };
    }

    case 'delete_chart':
      return {
        ...raw,
        success: true,
        message: raw.message || '图表已成功删除'
      };

    case 'find_replace':
    case 'find_and_replace':
    case 'search_cells': {
      const allMatches = (raw.matches || []).map((m: any) => ({
        address: m.address,
        row: m.row,
        column: m.column,
        text: m.text ?? m.value,
        value: m.value ?? m.text
      }));
      const total = raw.count ?? allMatches.length;
      // ISS-93：早期加载项不截断结果，maxResults 由桥接侧补齐；宿主已截断时保留其 truncated/maxResults。
      const requestedMax = Number((originalParams as any)?.maxResults ?? 50);
      const max = Number.isFinite(requestedMax) && requestedMax > 0 ? Math.trunc(requestedMax) : 50;
      const matches = allMatches.slice(0, max);
      return {
        ...raw,
        success: true,
        sheetName: raw.sheetName || originalParams?.sheetName || '',
        query: originalParams?.query || originalParams?.text || (originalParams as any)?.searchQuery || '',
        totalFound: raw.totalFound ?? total,
        count: total,
        returnedCount: matches.length,
        truncated: Boolean(raw.truncated) || matches.length < allMatches.length,
        ...(raw.replacedWith !== undefined ? { replacedWith: raw.replacedWith } : {}),
        matches
      };
    }

    case 'get_range_styles': {
      // ISS-93：Microsoft 侧返回 {font,fill,alignment,numberFormat}，WPS 侧返回 {styles,mixedOrUnavailableFields}。
      // 同名工具两宿主形状不同会让调用方写两份解析；这里把 MS 形状折算成 WPS 契约。
      const font = raw.font || {};
      const fill = raw.fill || {};
      const alignment = raw.alignment || {};
      const styles: Record<string, any> = {};
      const put = (key: string, value: any) => { if (value !== undefined && value !== null && value !== '') styles[key] = value; };
      put('fontName', font.name);
      put('fontSize', font.size);
      put('bold', font.bold);
      put('fontColor', font.color);
      put('backgroundColor', fill.color);
      put('numberFormat', raw.numberFormat);
      put('horizontalAlignment', alignment.horizontal);
      put('verticalAlignment', alignment.vertical);
      // MS 加载项未读取这些字段：如实列出为不可用，而不是留空让调用方误判"未设置"。
      const unavailable = ['wrapText', 'rowHeight', 'columnWidth', 'merged', 'mergeArea', 'borders'];
      const notes = ['Microsoft 通道只读取区域首个单元格的字体/填充/对齐，不检查区域一致性'];
      if (unavailable.length) notes.push(`Microsoft 通道未读取字段：${unavailable.join(', ')}`);
      return {
        ...raw,
        success: true,
        workbookName: raw.workbookName || originalParams?.workbookName || '',
        sheetName: raw.sheetName || originalParams?.sheetName || '',
        address: raw.address || originalParams?.address || '',
        mode: (originalParams as any)?.mode || 'summary',
        styles,
        mixedOrUnavailableFields: unavailable,
        notes
      };
    }

    case 'duplicate_sheet':
    case 'set_filter_and_sort': {
      const unsupportedFields = reportUnsupportedFields(baseMethod, originalParams);
      return {
        ...raw,
        success: raw.success ?? true,
        ...(unsupportedFields.length
          ? {
              unsupportedFields,
              warnings: [`Microsoft 通道未应用以下参数：${unsupportedFields.join(', ')}（加载项无对应实现，已在响应中如实暴露）`]
            }
          : {})
      };
    }

    case 'save_workbook':
    case 'save': {
      const wbName = raw.workbookName || originalParams?.workbookName || '工作簿1.xlsx';
      return {
        ...raw,
        success: true,
        saved: true,
        workbookName: wbName,
        fullName: raw.fullName || wbName,
        message: raw.message || `工作簿 [${wbName}] 已成功保存到磁盘`
      };
    }

    case 'capture_sheet_preview':
      return {
        ...raw,
        success: true,
        workbookName: raw.workbookName || originalParams?.workbookName || '工作簿1.xlsx',
        sheetName: raw.sheetName || originalParams?.sheetName || 'Sheet1',
        address: raw.address || originalParams?.address || 'A1',
        imageBase64: raw.imageBase64,
        imageMimeType: raw.imageMimeType || 'image/png',
        hasChart: raw.hasChart ?? false,
        // 宿主要是有 renderMode/synthetic 之类的"是否为真实渲染"标记，原样透传，不做美化（ISS-48 的同源问题：如实暴露）。
        ...(raw.renderMode !== undefined ? { renderMode: raw.renderMode } : {}),
        ...(raw.synthetic !== undefined ? { synthetic: raw.synthetic } : {}),
        message: raw.message || `已成功生成高保真渲染图`
      };

    default:
      return raw;
  }
}
