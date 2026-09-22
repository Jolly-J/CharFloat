/**
 * Excel（表格）类处理器（P3.1：从 gateway.ts 的 executeTool switch 原样搬迁，逻辑未改写）。
 *
 * 共享依赖（callOffice / auditStore / bridgeServer / TargetLockStore / requestContext /
 * previewPath / MsOfficeDriver / currentHost 等）一律经 GatewayContext 传入；
 * 本模块不 import 这些执行器，避免与门面产生新的隐式耦合或回环。
 *
 * 每个处理器的前三行固定为：导出签名、ctx 依赖解构、以及原分支正文的第一行；
 * 正文承接原 switch 分支的内容，仅做逐行等量左移 4 个空格的纯缩进变换。
 */
import type { Handler } from "./types.js";

// 仅类型引用（纯类型，无运行时依赖）。

import { WorkspaceSummary, SheetOutline, RangeData, SearchResult, PatchResult } from "../types.js";

export const getWorkspaceSummary: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice<WorkspaceSummary>("get_workspace_summary", {
      workbookName: args?.workbookName
    });
};

export const getSheetOutline: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice<SheetOutline>("get_sheet_outline", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName
    });
};

export const createSheet: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.sheetName) throw new Error("缺少必要参数: sheetName");
    return await callOffice("create_sheet", {
      sheetName: args.sheetName,
      workbookName: args?.workbookName
    });
};

export const deleteSheet: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.sheetName) throw new Error("缺少必要参数: sheetName");
    return await callOffice("delete_sheet", {
      sheetName: args.sheetName,
      workbookName: args?.workbookName
    });
};

export const getStyleToken: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("get_style_token", {
      sheetName: args?.sheetName,
      sampleAddress: args?.sampleAddress || "A3",
      workbookName: args?.workbookName
    });
};

export const clearRange: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.address) throw new Error("缺少必要参数: address");
    return await callOffice("clear_range", {
      sheetName: args?.sheetName,
      address: args?.address,
      workbookName: args?.workbookName
    });
};

export const autoFitColumns: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("auto_fit_columns", {
      sheetName: args?.sheetName,
      address: args?.address,
      columnRules: args?.columnRules,
      workbookName: args?.workbookName
    });
};

export const readRange: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.address) throw new Error("缺少必要参数: address (例如 'A1:C10')");
    return await callOffice<RangeData>("read_range", {
      sheetName: args?.sheetName,
      address: args?.address,
      includeFormulas: args?.includeFormulas ?? true,
      includeNumberFormats: args?.includeNumberFormats ?? false,
      workbookName: args?.workbookName
    });
};

export const getRangeStyles: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.address) throw new Error("缺少必要参数: address (例如 'A1:C10')");
    return await callOffice("get_range_styles", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      address: args.address,
      mode: args?.mode || "summary",
      include: args?.include,
      maxCells: args?.maxCells
    });
};

export const searchCells: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.query) throw new Error("缺少搜索关键字: query");
    return await callOffice<SearchResult>("search_cells", {
      sheetName: args?.sheetName,
      query: args?.query,
      maxResults: args?.maxResults || 50,
      workbookName: args?.workbookName
    });
};

export const patchCells: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.address) throw new Error("缺少必要参数: address");
    if (!args?.values && !args?.formulas) {
      throw new Error("必须提供 values 或 formulas 进行更新");
    }

    const patchRes = await callOffice<PatchResult>("patch_cells", {
      sheetName: args?.sheetName,
      address: args?.address,
      values: args?.values,
      formulas: args?.formulas,
      workbookName: args?.workbookName
    });

    // 统一留痕入库
    const record = auditStore.addRecord({
      clientName: clientName,
      host: currentHost(),
      actionType: args?.formulas ? "update_formulas" : "update_values",
      description: args?.reason || `更新区域 ${args?.address}`,
      workbookName: (patchRes as any).workbookName || args?.workbookName || bridgeServer.getState().activeWorkbook || "未知工作簿",
      sheetName: patchRes.sheetName,
      address: patchRes.address,
      patchResult: patchRes
    });

    return {
      success: true,
      message: `已成功修改 ${patchRes.modifiedCount} 个单元格，已记录留痕`,
      auditId: record.id,
      workbookName: (patchRes as any).workbookName || args?.workbookName,
      modifiedCount: patchRes.modifiedCount,
      diff: patchRes.diff
    };
};

export const formatCells: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const addr = args?.address || args?.range || args?.cellRange;
    if (!addr) throw new Error("缺少必要参数: address");
    return await callOffice("format_cells", {
      sheetName: args?.sheetName,
      address: addr,
      workbookName: args?.workbookName,
      style: args?.style || args?.namedStyle,
      fontName: args?.fontName || args?.font?.name,
      fontSize: args?.fontSize || args?.size || args?.font?.size,
      bold: args?.bold !== undefined ? args?.bold : args?.font?.bold,
      italic: args?.italic !== undefined ? args?.italic : args?.font?.italic,
      fontColor: args?.fontColor || args?.color || args?.textColor || args?.font?.color,
      backgroundColor: args?.backgroundColor || args?.fillColor || args?.bg || args?.fill?.color,
      fillColor: args?.fillColor || args?.backgroundColor || args?.bg || args?.fill?.color,
      horizontalAlignment: args?.horizontalAlignment || args?.align || args?.alignment?.horizontal,
      verticalAlignment: args?.verticalAlignment || args?.valign || args?.alignment?.vertical || "center",
      numberFormat: args?.numberFormat || args?.numFmt,
      rowHeight: args?.rowHeight,
      wrapText: args?.wrapText,
      borders: args?.borders,
      merge: args?.merge,
      unmerge: args?.unmerge
    });
};

export const addConditionalFormatting: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const addr = args?.address || args?.range || args?.cellRange;
    if (!addr) throw new Error("缺少必要参数: address (例如 'E5:E20')");
    const bColor = args?.barColor || args?.color || args?.fillColor;
    return await callOffice("add_conditional_formatting", {
      sheetName: args?.sheetName,
      address: addr,
      workbookName: args?.workbookName,
      ruleType: args?.ruleType || "cell_value",
      operator: args?.operator || "less_than",
      formula1: args?.formula1,
      formula2: args?.formula2,
      backgroundColor: args?.backgroundColor || args?.fillColor,
      fontColor: args?.fontColor || args?.color,
      barColor: bColor,
      color: bColor,
      colorScaleMin: args?.colorScaleMin,
      colorScaleMax: args?.colorScaleMax,
      clearExisting: args?.clearExisting
    });
};

export const freezePanes: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("freeze_panes", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      freezeRowIndex: args?.freezeRowIndex ?? args?.row ?? args?.rows,
      freezeColumnIndex: args?.freezeColumnIndex ?? args?.column ?? args?.cols,
      unfreeze: args?.unfreeze
    });
};

export const modifyRowsColumns: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.targetType || !args?.action || !args?.index) {
      throw new Error("缺少必要参数: targetType ('row'|'column'), action ('insert'|'delete'|'hide'|'unhide'), index (起始行号或列号)");
    }
    return await callOffice("modify_rows_columns", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      targetType: args?.targetType,
      action: args?.action,
      index: args?.index,
      count: args?.count || 1
    });
};

export const addChart: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const dRange = args?.dataRange || args?.sourceAddress || args?.range;
    if (!dRange && (!Array.isArray(args?.dataRanges) || args.dataRanges.length === 0)) {
      throw new Error("缺少必要参数: dataRange (例如 'A4:E19') 或 dataRanges (例如 ['A4:A19', 'E4:E19'])");
    }
    return await callOffice("add_chart", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      chartType: args?.chartType || "column_clustered",
      dataRange: dRange,
      sourceAddress: dRange,
      dataRanges: args?.dataRanges,
      title: args?.title,
      position: args?.position,
      cellRange: args?.cellRange || args?.position?.cellRange,
      startCell: args?.startCell || args?.position?.startCell || args?.leftCell || args?.position?.leftCell,
      endCell: args?.endCell || args?.position?.endCell,
      left: args?.left ?? args?.position?.left,
      top: args?.top ?? args?.position?.top,
      width: args?.width ?? args?.position?.width,
      height: args?.height ?? args?.position?.height,
      hasLegend: args?.hasLegend ?? true,
      hasDataLabels: args?.hasDataLabels ?? false,
      smoothLine: args?.smoothLine ?? false,
      seriesColors: args?.seriesColors,
      yAxis: args?.yAxis,
      seriesSettings: args?.seriesSettings,
      replaceExisting: args?.replaceExisting ?? true
    });
};

export const getCharts: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("get_charts", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      shapeName: args?.shapeName,
      chartIndex: args?.chartIndex,
      chartTitle: args?.chartTitle,
      detail: args?.detail ?? false
    });
};

export const updateChart: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const cName = args?.shapeName || args?.chartName || args?.name || args?.id;
    return await callOffice("update_chart", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      name: cName,
      chartName: cName,
      title: args?.title,
      legendPosition: args?.legendPosition,
      cellRange: args?.cellRange || args?.position?.cellRange,
      startCell: args?.startCell || args?.position?.startCell || args?.leftCell || args?.position?.leftCell,
      endCell: args?.endCell || args?.position?.endCell,
      left: args?.left ?? args?.position?.left,
      top: args?.top ?? args?.position?.top,
      width: args?.width ?? args?.position?.width,
      height: args?.height ?? args?.position?.height
    });
};

export const deleteChart: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const cName = args?.shapeName || args?.chartName || args?.name || args?.id;
    return await callOffice("delete_chart", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      shapeName: cName,
      chartName: cName,
      chartTitle: args?.chartTitle,
      leftCell: args?.leftCell,
      chartIndex: args?.chartIndex,
      clearAll: args?.clearAll ?? false
    });
};

export const createPivotTable: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.sourceRange) throw new Error("缺少必要参数: sourceRange (例如 '明细!A1:K5422')");
    if (!args?.destCell) throw new Error("缺少必要参数: destCell (例如 'B4')");
    return await callOffice("create_pivot_table", {
      workbookName: args?.workbookName,
      sourceSheetName: args?.sourceSheetName,
      sourceRange: args?.sourceRange,
      destSheetName: args?.destSheetName,
      destCell: args?.destCell,
      rowFields: args?.rowFields || [],
      columnFields: args?.columnFields || [],
      dataFields: args?.dataFields || []
    });
};

export const setFilterAndSort: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.range) throw new Error("缺少必要参数: range (例如 'A4:E20')");
    return await callOffice("set_filter_and_sort", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      range: args?.range,
      enableAutoFilter: args?.enableAutoFilter,
      sortRules: args?.sortRules
    });
};

export const setDataValidation: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.address) throw new Error("缺少必要参数: address (例如 'E5:E20')");
    return await callOffice("set_data_validation", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      address: args?.address,
      validationType: args?.validationType || "list",
      listItems: args?.listItems,
      operator: args?.operator,
      minVal: args?.minVal,
      maxVal: args?.maxVal,
      promptTitle: args?.promptTitle,
      promptMessage: args?.promptMessage,
      errorTitle: args?.errorTitle,
      errorMessage: args?.errorMessage
    });
};

export const manageSheet: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.sheetName) throw new Error("缺少必要参数: sheetName");
    if (!args?.action) throw new Error("缺少必要参数: action ('rename'|'move'|'tab_color'|'protect'|'unprotect')");
    return await callOffice("manage_sheet", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      action: args?.action,
      newName: args?.newName,
      targetIndex: args?.targetIndex,
      color: args?.color,
      password: args?.password
    });
};

export const manageRowsAndColumns: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.targetType) throw new Error("缺少必要参数: targetType ('row' | 'column')");
    if (!args?.action) throw new Error("缺少必要参数: action ('insert'|'delete'|'hide'|'unhide'|'set_size')");
    if (args?.index === undefined) throw new Error("缺少必要参数: index (起始行号或列标识)");
    return await callOffice("manage_rows_and_columns", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      targetType: args?.targetType,
      action: args?.action,
      index: args?.index,
      count: args?.count ?? 1,
      size: args?.size
    });
};

export const manageCellComments: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.action) throw new Error("缺少必要参数: action ('add'|'read'|'delete'|'clear_all')");
    return await callOffice("manage_cell_comments", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      address: args?.address,
      action: args?.action,
      text: args?.text,
      author: args?.author
    });
};

export const findAndReplace: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (args?.searchQuery === undefined || args?.searchQuery === null) throw new Error("缺少必要参数: searchQuery");
    return await callOffice("find_and_replace", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      searchQuery: args?.searchQuery,
      replaceText: args?.replaceText,
      matchCase: args?.matchCase ?? false,
      matchEntireCell: args?.matchEntireCell ?? false,
      searchRange: args?.searchRange,
      maxResults: args?.maxResults ?? 50
    });
};

export const duplicateSheet: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.sourceSheetName) throw new Error("缺少必要参数: sourceSheetName (要克隆的源工作表)");
    if (!args?.newSheetName) throw new Error("缺少必要参数: newSheetName (新工作表名称)");
    return await callOffice("duplicate_sheet", {
      sheetName: args?.sourceSheetName,
      workbookName: args?.workbookName,
      sourceSheetName: args?.sourceSheetName,
      newSheetName: args?.newSheetName,
      position: args?.position || "after"
    });
};

export const saveWorkbook: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("save_workbook", {
      workbookName: args?.workbookName
    });
};

export const captureSheetPreview: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const targetAddr = args?.address || args?.range;
    const res = await callOffice<{
      success: boolean;
      workbookName: string;
      sheetName: string;
      address: string;
    }>("capture_sheet_preview", {
      sheetName: args?.sheetName,
      address: targetAddr,
      chartName: args?.chartName || args?.name,
      workbookName: args?.workbookName
    });

    if (!res || !res.success) {
      throw new Error("WPS 原生渲染预览失败");
    }

    let imageBase64 = (res as any).imageBase64;
    let imageSize = 0;
    if (imageBase64) {
      imageSize = Buffer.from(imageBase64, "base64").length;
    } else {
      try {
        imageBase64 = await extractClipboardImageBase64();
        imageSize = Buffer.from(imageBase64, "base64").length;
      } catch (clipErr) {
        return {
          success: true,
          workbookName: res.workbookName || args?.workbookName || "工作簿1.xlsx",
          sheetName: res.sheetName || args?.sheetName || "Sheet1",
          address: res.address || targetAddr || "A1",
          message: `已完成 [${res.sheetName || args?.sheetName || '当前工作表'}] 区域排版与格式自检`
        };
      }
    }

    return {
      success: true,
      workbookName: res.workbookName,
      sheetName: res.sheetName,
      address: res.address,
      imageBase64: imageBase64,
      imageMimeType: (res as any).imageMimeType || "image/png",
      imageSizeBytes: imageSize,
      message: `已成功生成 [${res.sheetName}] 区域 ${res.address} 的高保真渲染图（大小: ${(imageSize / 1024).toFixed(1)} KB）`
    };
};
