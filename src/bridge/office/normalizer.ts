import type { WorkspaceSummary, SheetOutline, PatchResult } from '../types.js';

export interface NormalizedRequest {
  method: string;
  params: any;
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
 * 将 Gateway 标准工具调用请求正规化为 Office.js 加载项期望的契约 (统一参数适配管道)
 */
export function normalizeOfficeRequest(method: string, params: any = {}): NormalizedRequest {
  const p = { ...params };

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
    case 'manage_rows_and_columns':
      return { method: 'modify_rows_columns', params: p };

    case 'wps_get_charts':
    case 'excel_get_charts':
    case 'get_charts':
      return { method: 'get_charts', params: p };

    case 'wps_add_chart':
    case 'excel_add_chart':
    case 'add_chart': {
      const dataRange = p.dataRange || p.sourceAddress || p.sourceRange || p.range;
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

    default:
      return { method: method.replace(/^(wps_|excel_)/, ''), params: p };
  }
}

/**
 * 将 Office.js 加载项的原生响应标准化为 Gateway 与 MCP 契约
 */
export function normalizeOfficeResponse(method: string, raw: any, originalParams: any = {}): any {
  if (!raw) return raw;

  const baseMethod = method.replace(/^(wps_|excel_)/, '');

  switch (baseMethod) {
    case 'get_workspace_summary':
      return {
        hasOpenWorkbook: true,
        workbookName: raw.workbookName || originalParams?.workbookName || '工作簿1.xlsx',
        fullName: raw.fullName || raw.workbookName || originalParams?.workbookName || '工作簿1.xlsx',
        activeSheetName: raw.activeSheetName || raw.activeSheet || (raw.sheets && raw.sheets[0]?.name) || 'Sheet1',
        sheetCount: raw.sheetCount || (Array.isArray(raw.sheets) ? raw.sheets.length : 1),
        sheets: (raw.sheets || []).map((s: any, idx: number) => ({
          name: typeof s === 'string' ? s : s.name,
          position: typeof s === 'object' && s.position ? s.position : idx + 1,
          visibility: typeof s === 'object' && s.visibility ? s.visibility : 'Visible'
        })),
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
      const rows = raw.rowCount || (originalParams.values?.length) || (originalParams.formulas?.length) || 1;
      const cols = raw.columnCount || (originalParams.values?.[0]?.length) || (originalParams.formulas?.[0]?.length) || 1;
      const count = raw.modifiedCount || (rows * cols);
      return {
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
        success: true,
        sheetName: raw.sheetName || originalParams?.sheetName,
        address: raw.address || originalParams?.address,
        message: raw.message || `已成功设置区域 ${raw.address || originalParams?.address} 格式`
      };

    case 'auto_fit_columns':
      return {
        success: true,
        message: raw.message || '已成功自适应调整列宽'
      };

    case 'add_chart': {
      const chartId = raw.id || raw.name;
      return {
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

    case 'get_charts':
      return {
        success: true,
        count: raw.count ?? (Array.isArray(raw.charts) ? raw.charts.length : 0),
        charts: raw.charts || [],
        detail: !!originalParams?.detail
      };

    case 'delete_chart':
      return {
        success: true,
        message: raw.message || '图表已成功删除'
      };

    case 'find_replace':
    case 'search_cells': {
      const matches = (raw.matches || []).map((m: any) => ({
        address: m.address,
        row: m.row,
        column: m.column,
        text: m.text ?? m.value,
        value: m.value ?? m.text
      }));
      const total = raw.count ?? matches.length;
      return {
        success: true,
        sheetName: raw.sheetName || originalParams?.sheetName || '',
        query: originalParams?.query || originalParams?.text || '',
        totalFound: total,
        count: total,
        matches
      };
    }

    case 'save_workbook':
    case 'save': {
      const wbName = raw.workbookName || originalParams?.workbookName || '工作簿1.xlsx';
      return {
        success: true,
        saved: true,
        workbookName: wbName,
        fullName: raw.fullName || wbName,
        message: raw.message || `工作簿 [${wbName}] 已成功保存到磁盘`
      };
    }

    case 'capture_sheet_preview':
      return {
        success: true,
        workbookName: raw.workbookName || originalParams?.workbookName || '工作簿1.xlsx',
        sheetName: raw.sheetName || originalParams?.sheetName || 'Sheet1',
        address: raw.address || originalParams?.address || 'A1',
        imageBase64: raw.imageBase64,
        imageMimeType: raw.imageMimeType || 'image/png',
        hasChart: raw.hasChart ?? false,
        message: raw.message || `已成功生成高保真渲染图`
      };

    default:
      return raw;
  }
}
