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

import { WorkspaceSummary, SheetOutline, RangeData, SearchResult, PatchResult, AuditRecord } from "../types.js";
import type { GatewayContext } from "./types.js";

/**
 * 登记一次"只留痕、不可回滚"的写操作（ISS-48）。
 *
 * 背景：格式、条件格式、冻结窗格、行列结构、图表、数据校验等写操作原来完全不进审计，
 * 调用方既拿不到 auditId、也无法回答"谁在什么时候改了什么"。
 * 这里为它们补上操作事实记录：`rollbackable: false`、没有值快照，
 * `wps_rollback` 对这类记录会明确拒绝（没有可恢复的快照），不会假装能回滚。
 */
function recordWrite(
  ctx: GatewayContext,
  input: {
    actionType: AuditRecord["actionType"];
    description: string;
    sheetName?: string;
    address?: string;
    modifiedCount?: number;
    workbookName?: string;
  }
): AuditRecord {
  const { auditStore, clientName, currentHost, currentSession, args, bridgeServer } = ctx;
  return auditStore.addRecord({
    clientName,
    sessionId: currentSession(),
    host: currentHost(),
    actionType: input.actionType,
    description: input.description,
    workbookName: input.workbookName || args?.workbookName || bridgeServer.getState().activeWorkbook || "未知工作簿",
    sheetName: input.sheetName || args?.sheetName || "当前工作表",
    address: input.address || args?.address || args?.range || "",
    rollbackable: false,
    modifiedCount: input.modifiedCount ?? 0
  });
}

/** 把一次"不可回滚"写操作的结果补上审计标识与范围说明，避免调用方以为能回滚。 */
function withAuditScope(record: AuditRecord, result: unknown, note: string) {
  const base = result && typeof result === "object" && !Array.isArray(result) ? (result as Record<string, unknown>) : {};
  return {
    ...base,
    auditId: record.id,
    rollbackable: false,
    rollbackScope: { values: false, formulas: false, styles: false, charts: false, structure: false },
    auditNote: note
  };
}

export const getWorkspaceSummary: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    // ISS-02 / ISS-77：`hasOpenWorkbook` 原来反映的是"目标锁指向的文稿是否存在"，不是"当前打开了什么"。
    // 锁是加载项进程级全局、且可能来自更早的会话；不传目标时会回落到陈旧锁目标并报"未找到[另一个文件]"，
    // 而唯一能列出已打开文稿的工具恰好就是失败的那个（自举死锁）。
    //
    // 处理方式：
    //   1. 先按原语义解析目标（显式 workbookName > 本会话锁）；
    //   2. 目标不可用时**再查一次不带目标**的摘要，拿到真实的 openWorkbooks 列表；
    //   3. 返回体把"已打开列表"与"锁目标是否存在"拆成两个字段，并给出回退顺序说明。
    const requested = args?.workbookName;
    const lockTarget = locks?.excel;
    const targetSource: "request" | "session-lock" | "none" = ctx.targetSource
      || (lockTarget ? "session-lock" : "none");

    let primary: any = null;
    let primaryError: string | null = null;
    try {
      primary = await callOffice<WorkspaceSummary>("get_workspace_summary", { workbookName: requested });
    } catch (error: any) {
      primaryError = error?.message || String(error);
    }

    const targetResolved = Boolean(primary && primary.hasOpenWorkbook !== false);
    let fallback: any = null;
    if (!targetResolved) {
      // 不带目标再查一次：这是唯一能列出真实已打开文稿的路径（不传目标时宿主用活动文稿/唯一文稿）。
      try {
        fallback = await callOffice<WorkspaceSummary>("get_workspace_summary", {});
      } catch { /* 宿主确实没有可用文稿时保持原样 */ }
    }

    const openWorkbooks = ((fallback?.openWorkbooks || primary?.openWorkbooks || []) as any[]);
    const openCount = Array.isArray(openWorkbooks) ? openWorkbooks.length : 0;
    const anyOpen = Boolean(fallback?.hasOpenWorkbook || primary?.hasOpenWorkbook);
    const resolved = targetResolved ? primary : (fallback && fallback.hasOpenWorkbook ? fallback : null);
    const lockMissing = Boolean(lockTarget) && !targetResolved;

    const resolutionOrder = [
      "1) 本次请求的 workbookName",
      "2) 本会话的目标锁 wps_lock_target_document（可能来自更早的会话，不代表文稿还开着）",
      "3) 不传目标：宿主使用活动文稿；只打开了一个文稿时使用该文稿"
    ];
    const message = targetResolved
      ? "获取工作区摘要成功"
      : anyOpen
        ? `已打开 ${openCount || "若干"} 个工作簿，但未命中目标${lockTarget ? ` [${lockTarget}]` : ""}${primaryError ? `（宿主返回：${primaryError}）` : ""}。` +
          (lockTarget ? "目标可能来自陈旧的目标锁：请显式传 workbookName，或用 wps_unlock_target_document 清除锁后重试。" : "")
        : primaryError || "当前没有打开的工作簿";

    return {
      ...(resolved || {}),
      hasOpenWorkbook: anyOpen,
      openWorkbooks,
      openWorkbookCount: openCount || undefined,
      lockTarget: lockTarget
        ? { name: lockTarget, source: targetSource, exists: targetResolved, missing: lockMissing }
        : { name: null, source: "none", exists: false, missing: false },
      targetSource,
      targetResolution: resolutionOrder,
      ...(primaryError ? { targetError: primaryError } : {}),
      message
    };
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
    const result = await callOffice("create_sheet", {
      sheetName: args.sheetName,
      workbookName: args?.workbookName
    });
    const record = recordWrite(ctx, {
      actionType: "sheet_structure",
      description: `新建工作表 ${args.sheetName}`,
      sheetName: args.sheetName
    });
    return withAuditScope(record, result, "工作表结构操作只留痕、不可回滚（回滚不会删除已创建的工作表）。");
};

export const deleteSheet: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.sheetName) throw new Error("缺少必要参数: sheetName");
    const result = await callOffice("delete_sheet", {
      sheetName: args.sheetName,
      workbookName: args?.workbookName
    });
    const record = recordWrite(ctx, {
      actionType: "sheet_structure",
      description: `删除工作表 ${args.sheetName}（不可回滚）`,
      sheetName: args.sheetName
    });
    return withAuditScope(record, result, "删除工作表只留痕、不可回滚（内容不在快照范围内）。");
};

export const getStyleToken: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("get_style_token", {
      sheetName: args?.sheetName,
      sampleAddress: args?.sampleAddress || "A3",
      workbookName: args?.workbookName
    });
};

/** CAP-03 单元格内局部格式（富文本）。字段必须逐个转发——漏一个就等于该参数静默无效。 */
export const formatTextSegment: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.address) throw new Error("缺少必要参数: address");
    if (args?.find === undefined && args?.start === undefined) {
      throw new Error("需要提供 find（按文本定位，可配 occurrence）或 start（1 基起始位置，可配 length）");
    }
    if (args?.bold === undefined && args?.italic === undefined && args?.underline === undefined &&
        args?.fontColor === undefined && args?.fontSize === undefined && args?.fontName === undefined) {
      throw new Error("至少要给出一项格式：bold / italic / underline / fontColor / fontSize / fontName");
    }
    return await callOffice("format_text_segment", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      address: args?.address,
      find: args?.find,
      occurrence: args?.occurrence,
      start: args?.start,
      length: args?.length,
      bold: args?.bold,
      italic: args?.italic,
      underline: args?.underline,
      fontColor: args?.fontColor,
      fontSize: args?.fontSize,
      fontName: args?.fontName
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
    const result = await callOffice("auto_fit_columns", {
      sheetName: args?.sheetName,
      address: args?.address,
      columnRules: args?.columnRules,
      workbookName: args?.workbookName
    });
    const record = recordWrite(ctx, {
      actionType: "format",
      description: `自适应列宽${args?.address ? `（${args.address}）` : ""}`,
      address: args?.address
    });
    return withAuditScope(record, result, "列宽调整只留痕、不可回滚。");
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
      sessionId: currentSession(),
      host: currentHost(),
      actionType: args?.formulas ? "update_formulas" : "update_values",
      description: args?.reason || `更新区域 ${args?.address}`,
      workbookName: (patchRes as any).workbookName || args?.workbookName || bridgeServer.getState().activeWorkbook || "未知工作簿",
      sheetName: patchRes.sheetName,
      address: patchRes.address,
      patchResult: patchRes,
      rollbackable: true
    });

    return {
      success: true,
      message: `已成功修改 ${patchRes.modifiedCount} 个单元格，已记录留痕`,
      auditId: record.id,
      rollbackable: true,
      workbookName: (patchRes as any).workbookName || args?.workbookName,
      modifiedCount: patchRes.modifiedCount,
      diff: patchRes.diff
    };
};

export const formatCells: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const addr = args?.address || args?.range || args?.cellRange;
    if (!addr) throw new Error("缺少必要参数: address");
    const result = await callOffice("format_cells", {
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
    // ISS-48：格式类操作没有值快照，不回滚但必须留痕，否则"谁改的样式"无从回答。
    const record = recordWrite(ctx, {
      actionType: "format",
      description: args?.reason || `设置区域 ${addr} 格式`,
      address: addr
    });
    return withAuditScope(record, result, "格式修改只留痕、不可回滚（本记录没有值/样式快照）。");
};

export const addConditionalFormatting: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const addr = args?.address || args?.range || args?.cellRange;
    if (!addr) throw new Error("缺少必要参数: address (例如 'E5:E20')");
    const bColor = args?.barColor || args?.color || args?.fillColor;
    const result = await callOffice("add_conditional_formatting", {
      action: args?.action,
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
      clearExisting: args?.clearExisting,
      // CAP-06：图标集 / Top-N / 重复值 / 公式 / 文字包含的参数必须逐个转发
      iconSet: args?.iconSet,
      iconThresholds: args?.iconThresholds,
      topBottom: args?.topBottom,
      topRank: args?.topRank,
      topPercent: args?.topPercent,
      containsText: args?.containsText
    });
    const record = recordWrite(ctx, {
      actionType: "conditional_format",
      description: args?.reason || `为 ${addr} 添加条件格式（规则 ${args?.ruleType || "cell_value"}）`,
      address: addr
    });
    return withAuditScope(record, result, "条件格式只留痕、不可回滚；清除需用宿主脚本或后续工具。");
};

export const freezePanes: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const result = await callOffice("freeze_panes", {
      action: args?.action,
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      freezeRowIndex: args?.freezeRowIndex ?? args?.row ?? args?.rows,
      freezeColumnIndex: args?.freezeColumnIndex ?? args?.column ?? args?.cols,
      unfreeze: args?.unfreeze
    });
    const record = recordWrite(ctx, {
      actionType: "freeze_panes",
      description: args?.unfreeze
        ? "取消冻结窗格"
        : `冻结窗格（行 ${args?.freezeRowIndex ?? args?.row ?? args?.rows ?? 0} / 列 ${args?.freezeColumnIndex ?? args?.column ?? args?.cols ?? 0}）`,
      address: args?.address || ""
    });
    return withAuditScope(record, result, "冻结窗格只留痕、不可回滚。");
};

export const modifyRowsColumns: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.targetType || !args?.action || !args?.index) {
      throw new Error("缺少必要参数: targetType ('row'|'column'), action ('insert'|'delete'|'hide'|'unhide'), index (起始行号或列号)");
    }
    const result = await callOffice("modify_rows_columns", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      targetType: args?.targetType,
      action: args?.action,
      index: args?.index,
      count: args?.count || 1
    });
    const dim = args?.targetType === "column" ? "列" : "行";
    const actionType = args?.action === "delete" ? "delete_dimension"
      : args?.action === "insert" ? (args?.targetType === "column" ? "insert_col" : "insert_row")
      : "hide_dimension";
    const record = recordWrite(ctx, {
      actionType,
      description: `${args?.action} ${dim}（起始 ${args?.index}，共 ${args?.count || 1}）`,
      address: String(args?.index ?? "")
    });
    return withAuditScope(record, result, "行列结构修改只留痕、不可回滚：值快照无法表达行列位移带来的整体变化。");
};

export const addChart: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const dRange = args?.dataRange || args?.sourceAddress || args?.range;
    if (!dRange && (!Array.isArray(args?.dataRanges) || args.dataRanges.length === 0)) {
      throw new Error("缺少必要参数: dataRange (例如 'A4:E19') 或 dataRanges (例如 ['A4:A19', 'E4:E19'])");
    }
    const result = await callOffice("add_chart", {
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
    const record = recordWrite(ctx, {
      actionType: "chart",
      description: args?.reason || `创建图表（类型 ${args?.chartType || "column_clustered"}，数据源 ${dRange || (args?.dataRanges || []).join("+")}）`,
      address: dRange || ""
    });
    return withAuditScope(record, result, "新增图表只留痕、不可回滚：回滚不会删除已创建的图表。");
};

export const getCharts: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("get_charts", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      shapeName: args?.shapeName,
      chartIndex: args?.chartIndex,
      chartTitle: args?.chartTitle,
      detail: args?.detail ?? false,
      // CAP-08 诊断入口：让 Office.js 任务窗格跑一遍形状能力自检，用于确认窗格是否已换到最新代码
      shapeSelfTest: args?.shapeSelfTest ?? false,
      selfTestKeep: args?.selfTestKeep ?? false
    });
};

export const updateChart: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const cName = args?.shapeName || args?.chartName || args?.name || args?.id;
    const result = await callOffice("update_chart", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      name: cName,
      chartName: cName,
      chartIndex: args?.chartIndex,
      title: args?.title,
      chartType: args?.chartType,
      hasLegend: args?.hasLegend,
      showDataLabels: args?.showDataLabels,
      fontName: args?.fontName,
      legendPosition: args?.legendPosition,
      dataLabelColorMatchesSeries: args?.dataLabelColorMatchesSeries,
      position: args?.position,
      cellRange: args?.cellRange || args?.position?.cellRange,
      startCell: args?.startCell || args?.position?.startCell || args?.leftCell || args?.position?.leftCell,
      endCell: args?.endCell || args?.position?.endCell,
      left: args?.left ?? args?.position?.left,
      top: args?.top ?? args?.position?.top,
      width: args?.width ?? args?.position?.width,
      height: args?.height ?? args?.position?.height
    });
    const record = recordWrite(ctx, {
      actionType: "chart",
      description: `更新图表 ${cName || "(未指定)"}`,
      address: args?.cellRange || args?.startCell || ""
    });
    return withAuditScope(record, result, "图表更新只留痕、不可回滚。");
};

export const deleteChart: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const cName = args?.shapeName || args?.chartName || args?.name || args?.id;
    const result = await callOffice("delete_chart", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      shapeName: cName,
      chartName: cName,
      chartTitle: args?.chartTitle,
      leftCell: args?.leftCell,
      chartIndex: args?.chartIndex,
      clearAll: args?.clearAll ?? false
    });
    const record = recordWrite(ctx, {
      actionType: "chart",
      description: args?.clearAll ? "删除工作表中的全部图表" : `删除图表 ${cName || args?.chartTitle || args?.chartIndex || "(未指定)"}`,
      address: args?.leftCell || ""
    });
    return withAuditScope(record, result, "删除图表只留痕、不可回滚（回滚不会重建图表）。");
};

export const createPivotTable: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (args?.action !== "read" && !args?.sourceRange) throw new Error("缺少必要参数: sourceRange (例如 '明细!A1:K5422')");
    if (args?.action !== "read" && !args?.destCell) throw new Error("缺少必要参数: destCell (例如 'B4')");
    const result = await callOffice("create_pivot_table", {
      action: args?.action,
      workbookName: args?.workbookName,
      sourceSheetName: args?.sourceSheetName,
      sourceRange: args?.sourceRange,
      destSheetName: args?.destSheetName,
      destCell: args?.destCell,
      rowFields: args?.rowFields || [],
      columnFields: args?.columnFields || [],
      dataFields: args?.dataFields || []
    });
    const record = recordWrite(ctx, {
      actionType: "pivot_table",
      description: `创建数据透视表（源 ${args?.sourceRange} → ${args?.destSheetName || args?.sourceSheetName || "当前表"}!${args?.destCell}）`,
      sheetName: args?.destSheetName || args?.sourceSheetName,
      address: args?.destCell
    });
    return withAuditScope(record, result, "创建透视表只留痕、不可回滚。");
};

export const setFilterAndSort: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (args?.action !== "read" && !args?.range) throw new Error("缺少必要参数: range (例如 'A4:E20')");
    const result = await callOffice("set_filter_and_sort", {
      action: args?.action,
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      range: args?.range,
      enableAutoFilter: args?.enableAutoFilter,
      sortRules: args?.sortRules
    });
    const record = recordWrite(ctx, {
      actionType: "filter_sort",
      description: `筛选/排序 ${args?.range}（筛选 ${args?.enableAutoFilter === undefined ? "不变" : args?.enableAutoFilter ? "开启" : "关闭"}，排序规则 ${(args?.sortRules || []).length} 条）`,
      address: args?.range
    });
    return withAuditScope(record, result, "筛选与排序只留痕、不可回滚（排序不可逆地改变了行顺序，值快照无法安全恢复）。");
};

export const setDataValidation: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.address) throw new Error("缺少必要参数: address (例如 'E5:E20')");
    const result = await callOffice("set_data_validation", {
      action: args?.action,
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
    const record = recordWrite(ctx, {
      actionType: "data_validation",
      description: `设置数据有效性 ${args?.address}（类型 ${args?.validationType || "list"}）`,
      address: args?.address
    });
    return withAuditScope(record, result, "数据有效性只留痕、不可回滚；如需清除请用宿主脚本或覆盖设置。");
};

export const manageSheet: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.sheetName) throw new Error("缺少必要参数: sheetName");
    if (!args?.action) throw new Error("缺少必要参数: action ('rename'|'move'|'tab_color'|'protect'|'unprotect')");
    const result = await callOffice("manage_sheet", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      action: args?.action,
      newName: args?.newName,
      targetIndex: args?.targetIndex,
      color: args?.color,
      password: args?.password
    });
    const record = recordWrite(ctx, {
      actionType: "sheet_structure",
      description: `工作表管理 ${args?.action}（${args?.sheetName}${args?.newName ? ` → ${args.newName}` : ""}）`,
      sheetName: args?.sheetName
    });
    return withAuditScope(record, result, "工作表管理只留痕、不可回滚（含密码保护，回滚不会还原保护状态）。");
};

export const manageRowsAndColumns: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.targetType) throw new Error("缺少必要参数: targetType ('row' | 'column')");
    if (!args?.action) throw new Error("缺少必要参数: action ('insert'|'delete'|'hide'|'unhide'|'set_size')");
    if (args?.index === undefined) throw new Error("缺少必要参数: index (起始行号或列标识)");
    const result = await callOffice("manage_rows_and_columns", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      targetType: args?.targetType,
      action: args?.action,
      index: args?.index,
      count: args?.count ?? 1,
      size: args?.size
    });
    const dim = args?.targetType === "column" ? "列" : "行";
    const actionType = args?.action === "delete" ? "delete_dimension"
      : args?.action === "insert" ? (args?.targetType === "column" ? "insert_col" : "insert_row")
      : "hide_dimension";
    const record = recordWrite(ctx, {
      actionType,
      description: `manage_rows_and_columns ${args?.action} ${dim}（起始 ${args?.index}，共 ${args?.count ?? 1}）`,
      address: String(args?.index ?? "")
    });
    return withAuditScope(record, result, "行列结构修改只留痕、不可回滚（行列位移无法用值快照表达）。");
};

export const manageCellComments: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.action) throw new Error("缺少必要参数: action ('add'|'read'|'delete'|'clear_all')");
    const result = await callOffice("manage_cell_comments", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      address: args?.address,
      action: args?.action,
      text: args?.text,
      author: args?.author
    });
    // 读批注不是写操作，不登记留痕。
    if (args.action === "read") return result;
    const record = recordWrite(ctx, {
      actionType: "comment",
      description: `批注操作 ${args.action}${args?.address ? `（${args.address}）` : ""}`,
      address: args?.address
    });
    return withAuditScope(record, result, "批注操作只留痕、不可回滚（回滚不会还原或删除批注）。");
};

export const findAndReplace: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (args?.searchQuery === undefined || args?.searchQuery === null) throw new Error("缺少必要参数: searchQuery");
    const result = await callOffice("find_and_replace", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      searchQuery: args?.searchQuery,
      replaceText: args?.replaceText,
      matchCase: args?.matchCase ?? false,
      matchEntireCell: args?.matchEntireCell ?? false,
      searchRange: args?.searchRange,
      maxResults: args?.maxResults ?? 50
    });
    // 只查找不替换时没有改动文档，不登记留痕。
    if (args?.replaceText === undefined) return result;
    const record = recordWrite(ctx, {
      actionType: "find_replace",
      description: `查找替换「${args.searchQuery}」→「${args.replaceText}」`,
      address: args?.searchRange || ""
    });
    return withAuditScope(record, result, "批量替换只留痕、不可回滚（整片区域被改写，值快照不适用）。");
};

export const duplicateSheet: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.sourceSheetName) throw new Error("缺少必要参数: sourceSheetName (要克隆的源工作表)");
    if (!args?.newSheetName) throw new Error("缺少必要参数: newSheetName (新工作表名称)");
    const result = await callOffice("duplicate_sheet", {
      sheetName: args?.sourceSheetName,
      workbookName: args?.workbookName,
      sourceSheetName: args?.sourceSheetName,
      newSheetName: args?.newSheetName,
      position: args?.position || "after"
    });
    const record = recordWrite(ctx, {
      actionType: "sheet_structure",
      description: `克隆工作表 ${args.sourceSheetName} → ${args.newSheetName}`,
      sheetName: args?.newSheetName
    });
    return withAuditScope(record, result, "克隆工作表只留痕、不可回滚（回滚不会删除新表）。");
};

export const saveWorkbook: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const result = await callOffice("save_workbook", {
      workbookName: args?.workbookName
    });
    // ISS-16：保存是**工作簿级**的，会把整个工作簿的内存状态写盘，包含其他会话/任务尚未完成、
    // 也不打算保留的中间结果（4 路并行子代理场景已实际发生）。这里在返回体里给出针对性警告，
    // 而不是让调用方以为"只保存了我的那部分"。
    const warning =
      "保存是工作簿级操作：会把该工作簿内存中**所有**未保存改动一并落盘，" +
      "包括其他会话/任务正在进行、尚未完成的中间结果。并发编辑同一工作簿时，" +
      "请先确认他人已保存或已完成，或让每个任务使用独立副本。";
    const record = recordWrite(ctx, {
      actionType: "save",
      description: `保存工作簿${args?.workbookName ? ` ${args.workbookName}` : ""}（工作簿级，含他人在途改动）`
    });
    const base = result && typeof result === "object" && !Array.isArray(result) ? (result as Record<string, unknown>) : {};
    return {
      ...base,
      auditId: record.id,
      rollbackable: false,
      scope: "workbook",
      warning,
      ...(Array.isArray(ctx.ignoredParams) && ctx.ignoredParams.length
        ? { ignoredParams: ctx.ignoredParams, ignoredParamsNote: `以下参数对保存操作无意义，已忽略：${ctx.ignoredParams.join(", ")}` }
        : {})
    };
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

/** CAP-01 打印与分页设置。字段逐个转发（漏一个就等于该参数静默无效）。 */
export const configurePrintLayout: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("configure_print_layout", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      action: args?.action || "apply",
      printArea: args?.printArea,
      printTitleRows: args?.printTitleRows,
      printTitleColumns: args?.printTitleColumns,
      orientation: args?.orientation,
      paperSize: args?.paperSize,
      zoom: args?.zoom,
      fitToPagesWide: args?.fitToPagesWide,
      fitToPagesTall: args?.fitToPagesTall,
      centerHorizontally: args?.centerHorizontally,
      centerVertically: args?.centerVertically,
      printGridlines: args?.printGridlines,
      leftMargin: args?.leftMargin,
      rightMargin: args?.rightMargin,
      topMargin: args?.topMargin,
      bottomMargin: args?.bottomMargin,
      headerMargin: args?.headerMargin,
      footerMargin: args?.footerMargin,
      centerHeader: args?.centerHeader,
      leftHeader: args?.leftHeader,
      rightHeader: args?.rightHeader,
      centerFooter: args?.centerFooter,
      leftFooter: args?.leftFooter,
      rightFooter: args?.rightFooter,
      clearPrintArea: args?.clearPrintArea,
      addHorizontalPageBreak: args?.addHorizontalPageBreak,
      addVerticalPageBreak: args?.addVerticalPageBreak,
      clearPageBreaks: args?.clearPageBreaks
    });
};

/** CAP-02 导出 PDF。输出路径由桥接生成（必须是**宿主可写**目录），落盘由桥接侧校验。 */
export const exportSheetPdf: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const outputPath = previewPath("pdf");
    const res: any = await callOffice("export_sheet_pdf", {
      sheetName: args?.sheetName,
      workbookName: args?.workbookName,
      outputPath,
      scope: args?.scope || "workbook",
      quality: args?.quality || "standard"
    });
    if (!res?.success) throw new Error(`导出 PDF 失败：${res?.hostError || res?.error || "宿主未给出原因"}`);
    // 加载项没有文件系统访问：**落盘校验只能在桥接侧做**。
    // Word 预览的教训（ISS-111）：宿主返回成功但文件可能根本没写出。
    const { existsSync, statSync } = await import("node:fs");
    if (!existsSync(outputPath)) {
      throw new Error(
        `导出 PDF 未落盘：宿主接受了请求，但 ${outputPath} 不存在。` +
        `请确认输出目录是宿主可写位置（WPS 为沙箱应用）。宿主返回：${JSON.stringify(res).slice(0, 200)}`
      );
    }
    const size = statSync(outputPath).size;
    if (size === 0) throw new Error(`导出 PDF 落盘但为空文件：${outputPath}`);
    return { ...res, outputPath, verifiedOnDisk: true, fileSizeBytes: size, message: `已导出 PDF 并确认落盘（${size} 字节）：${outputPath}` };
};

/** CAP-08 MS 侧矢量形状。宿主差异（本机实测）：矩形/文本框/分组/层级/导图可用；
 *  直线报「当前对象不允许此操作」，SVG 与 getActiveShape 在本机 Excel 无对应 API。 */
export const addShape: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("add_shape", {
      workbookName: args?.workbookName, sheetName: args?.sheetName, kind: args?.kind, shapeType: args?.shapeType, text: args?.text,
      left: args?.left, top: args?.top, width: args?.width, height: args?.height, rotation: args?.rotation,
      x1: args?.x1, y1: args?.y1, x2: args?.x2, y2: args?.y2,
      fillColor: args?.fillColor, fill: args?.fill, lineColor: args?.lineColor, lineWeight: args?.lineWeight,
      fontName: args?.fontName, fontSize: args?.fontSize, bold: args?.bold, italic: args?.italic, wordArtPreset: args?.wordArtPreset, fontColor: args?.fontColor,
      textAlign: args?.textAlign, textVAlign: args?.textVAlign, marginLeft: args?.marginLeft, marginRight: args?.marginRight,
      name: args?.name
    });
};

export const groupShapes: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    // 宿主读的是 `names`，schema 对外是 `shapeNames`：两个都传，避免"schema 有但宿主读不到"的静默失效。
    return await callOffice("group_shapes", { workbookName: args?.workbookName, sheetName: args?.sheetName, names: args?.shapeNames || args?.names, shapeNames: args?.shapeNames, groupName: args?.groupName });
};

export const ungroupShapes: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("ungroup_shapes", { workbookName: args?.workbookName, sheetName: args?.sheetName, name: args?.shapeName || args?.name, shapeName: args?.shapeName, shapeId: args?.shapeId });
};

export const setShapeZOrder: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.zOrder) throw new Error("缺少必要参数: zOrder (bringToFront | sendToBack | bringForward | sendBackward)");
    return await callOffice("set_shape_zorder", { workbookName: args?.workbookName, sheetName: args?.sheetName, name: args?.shapeName || args?.name, shapeName: args?.shapeName, shapeId: args?.shapeId, zOrder: args?.zOrder });
};

export const exportShapeImage: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("export_shape_image", { workbookName: args?.workbookName, sheetName: args?.sheetName, shapeName: args?.shapeName, shapeId: args?.shapeId, format: args?.format, scale: args?.scale });
};

/** CAP-07 WPS 表格矢量绘图：读回全部形状（绘图能力的验收入口，双宿主通用）。 */
export const listShapes: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("list_shapes", {
      sheetName: args?.sheetName, workbookName: args?.workbookName,
      detail: args?.detail ?? true, filterName: args?.filterName
    });
};

/** 修改形状：位置/尺寸/旋转/填充/线条/文字/可见性，或删除。 */
export const updateShape: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("update_shape", {
      sheetName: args?.sheetName, workbookName: args?.workbookName,
      name: args?.name, shapeName: args?.shapeName, shapeIndex: args?.shapeIndex, shapeId: args?.shapeId,
      action: args?.action || "update",
      left: args?.left, top: args?.top, width: args?.width, height: args?.height, rotation: args?.rotation,
      fillColor: args?.fillColor, lineColor: args?.lineColor, lineWeight: args?.lineWeight,
      text: args?.text, visible: args?.visible, newName: args?.newName
    });
};

/** 工作表视图设置：网格线/行列标题/缩放。画布页（矢量形状拼版）必须先隐藏网格线。 */
export const setSheetView: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("set_sheet_view", {
      sheetName: args?.sheetName, workbookName: args?.workbookName,
      showGridlines: args?.showGridlines, showHeadings: args?.showHeadings, zoom: args?.zoom
    });
};

// ── CAP-15~20 第二类常用能力
export const copyRange: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
  return await callOffice("copy_range", { sheetName: args?.sheetName, workbookName: args?.workbookName,
    sourceRange: args?.sourceRange, destRange: args?.destRange, destSheetName: args?.destSheetName,
    copyType: args?.copyType, transpose: args?.transpose });
};

export const manageHyperlink: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
  return await callOffice("manage_hyperlink", { sheetName: args?.sheetName, workbookName: args?.workbookName,
    action: args?.action, address: args?.address, url: args?.url, displayText: args?.displayText,
    tooltip: args?.tooltip, targetAddress: args?.targetAddress });
};

export const manageNamedRange: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
  return await callOffice("manage_named_range", { workbookName: args?.workbookName, action: args?.action,
    name: args?.name, refersTo: args?.refersTo, comment: args?.comment });
};

export const manageDocumentProperties: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
  return await callOffice("manage_document_properties", { workbookName: args?.workbookName, action: args?.action,
    properties: args?.properties, propertyNames: args?.propertyNames });
};

export const manageTable: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
  return await callOffice("manage_table", { sheetName: args?.sheetName, workbookName: args?.workbookName,
    action: args?.action, tableName: args?.tableName, address: args?.address, styleName: args?.styleName,
    newName: args?.newName, hasHeaders: args?.hasHeaders, totalsRow: args?.totalsRow });
};

export const managePictures: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
  return await callOffice("manage_pictures", { sheetName: args?.sheetName, workbookName: args?.workbookName,
    action: args?.action, filePath: args?.filePath, pictureName: args?.pictureName, pictureIndex: args?.pictureIndex,
    left: args?.left, top: args?.top, width: args?.width, height: args?.height });
};

/** CAP-22 图表导图。宿主只负责调用 Export；**落盘与否由桥接侧用文件系统核对**——
 *  写到 WPS 沙箱不可达的路径时宿主不报错但文件不存在，这里如实回填 fileWritten。 */
export const exportChartImage: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
  const result: any = await callOffice("export_chart_image", {
    sheetName: args?.sheetName, workbookName: args?.workbookName,
    chartName: args?.chartName, chartIndex: args?.chartIndex,
    outputPath: args?.outputPath, format: args?.format
  });
  const p = typeof result?.outputPath === "string" ? result.outputPath : null;
  if (p) {
    try {
      const st = await import("node:fs").then(fs => fs.statSync(p));
      result.fileWritten = true;
      result.fileSizeBytes = st.size;
      result.message = `已导出图表 ${result.chartName ?? ""} → ${p}（${result.fileSizeBytes} 字节）`;
    } catch (e) {
      result.fileWritten = false;
      result.warnings = [...(result.warnings ?? []), `宿主已执行导出，但文件不存在：${p}。请确认 outputPath 落在 WPS 可写目录内。`];
      result.message = `图表导出未落盘：${p}`;
    }
  }
  return result;
};

/** CAP-53 新建工作簿。宿主 `Workbooks.Add()` + `SaveAs` 一直可用，只是此前没有工具入口。 */
export const createWorkbook: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
  return await callOffice("create_workbook", { savePath: args?.savePath, sheetName: args?.sheetName });
};
