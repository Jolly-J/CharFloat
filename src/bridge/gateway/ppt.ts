/**
 * PowerPoint（演示）类处理器（P3.1：从 gateway.ts 的 executeTool switch 原样搬迁，逻辑未改写）。
 *
 * 共享依赖（callOffice / auditStore / bridgeServer / TargetLockStore / requestContext /
 * previewPath / MsOfficeDriver / currentHost 等）一律经 GatewayContext 传入；
 * 本模块不 import 这些执行器，避免与门面产生新的隐式耦合或回环。
 *
 * 每个处理器的前三行固定为：导出签名、ctx 依赖解构、以及原分支正文的第一行；
 * 正文承接原 switch 分支的内容，仅做逐行等量左移 4 个空格的纯缩进变换。
 */
import type { Handler } from "./types.js";

import fs from "fs";

export const readPresentation: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("ppt_read_presentation", {
      presentationName: args?.presentationName,
      includeNotes: args?.includeNotes ?? true,
      maxSlides: args?.maxSlides ?? 50
    });
};

export const getSlideShapes: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("ppt_get_slide_shapes", {
      presentationName: args?.presentationName,
      slideIndex: args?.slideIndex
    });
};

export const generateDeck: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.slides || !Array.isArray(args?.slides)) throw new Error("缺少必要参数: slides (幻灯片结构化大纲数组)");
    return await callOffice("ppt_generate_deck", {
      presentationName: args?.presentationName,
      themeColor: args?.themeColor,
      themePreset: args?.themePreset || "business_blue",
      slides: args?.slides
    });
};

export const manageSlides: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.action) throw new Error("缺少必要参数: action ('add'|'delete'|'move'|'duplicate'|'set_background')");
    return await callOffice("ppt_manage_slides", {
      presentationName: args?.presentationName,
      action: args?.action,
      slideIndex: args?.slideIndex,
      targetIndex: args?.targetIndex,
      layoutIndex: args?.layoutIndex,
      backgroundColor: args?.backgroundColor
    });
};

export const manageTable: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.action) throw new Error("缺少必要参数: action ('create_table'|'read_table'|'set_table_data'|'set_cell_text'|'style_table')");
    return await callOffice("ppt_manage_table", {
      presentationName: args?.presentationName,
      slideIndex: args?.slideIndex,
      action: args?.action,
      shapeId: args?.shapeId,
      tableIndex: args?.tableIndex,
      rows: args?.rows,
      columns: args?.columns,
      left: args?.left,
      top: args?.top,
      width: args?.width,
      height: args?.height,
      data: args?.data,
      row: args?.row,
      column: args?.column,
      text: args?.text,
      fontSize: args?.fontSize,
      fontColor: args?.fontColor,
      fontBold: args?.fontBold,
      fillColor: args?.fillColor,
      headerFillColor: args?.headerFillColor,
      headerFontSize: args?.headerFontSize,
      bodyFontSize: args?.bodyFontSize,
      borderColor: args?.borderColor,
      columnWidths: args?.columnWidths,
      rowHeights: args?.rowHeights,
      zebra: args?.zebra
    });
};

export const addBusinessCards: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.cards || !Array.isArray(args?.cards)) throw new Error("缺少必要参数: cards (商业信息卡片数组)");
    return await callOffice("ppt_add_business_cards", {
      presentationName: args?.presentationName,
      slideIndex: args?.slideIndex,
      columnCount: args?.columnCount || 3,
      cards: args?.cards,
      topY: args?.topY,
      cardHeight: args?.cardHeight
    });
};

export const insertNativeChart: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.categories || !args?.series) throw new Error("缺少必要参数: categories 和 series 数据");
    return await callOffice("ppt_insert_native_chart", {
      presentationName: args?.presentationName,
      slideIndex: args?.slideIndex,
      chartType: args?.chartType || "column",
      title: args?.title,
      hasLegend: args?.hasLegend,
      showDataLabels: args?.showDataLabels,
      categories: args?.categories,
      series: args?.series,
      left: args?.left,
      top: args?.top,
      width: args?.width,
      height: args?.height
    });
};

export const manageShapesAndMedia: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.action) throw new Error("缺少必要参数: action ('add_textbox'|'add_shape'|'update_shape'|'swap_shapes'|'set_z_order'|'align_shapes'|'delete_shape')");
    return await callOffice("ppt_manage_shapes_and_media", {
      presentationName: args?.presentationName,
      slideIndex: args?.slideIndex,
      action: args?.action,
      shapeId: args?.shapeId,
      shapeId1: args?.shapeId1,
      shapeId2: args?.shapeId2,
      shapeType: args?.shapeType || "rectangle",
      text: args?.text,
      left: args?.left,
      top: args?.top,
      width: args?.width,
      height: args?.height,
      fontSize: args?.fontSize,
      fontColor: args?.fontColor,
      fontBold: args?.fontBold,
      alignment: args?.alignment,
      fillColor: args?.fillColor,
      lineColor: args?.lineColor,
      rotation: args?.rotation,
      zOrderAction: args?.zOrderAction,
      alignType: args?.alignType,
      shapeIds: args?.shapeIds
    });
};

export const captureSlidePreview: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const outputPath = previewPath('png');
    const res: any = await callOffice("ppt_capture_slide_preview", { presentationName: args?.presentationName, slideIndex: args?.slideIndex, outputPath });
    if (!res?.success || !fs.existsSync(outputPath)) throw new Error(res?.error || 'PPT 未生成预览');
    return { ...res, imageBase64: fs.readFileSync(outputPath).toString('base64'), imageMimeType: 'image/png' };
};
