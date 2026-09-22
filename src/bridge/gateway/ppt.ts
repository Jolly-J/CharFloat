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
    if (!args?.action) throw new Error("缺少必要参数: action ('add'|'delete'|'move'|'duplicate'|'set_background'|'save'|'save_as'|'new_presentation')");
    // ISS-106：`filePath` / `format` 原来没转发——schema 里声明了却到不了宿主，
    // save_as 只能落到宿主默认路径、format 完全不生效。
    return await callOffice("ppt_manage_slides", {
      presentationName: args?.presentationName,
      action: args?.action,
      slideIndex: args?.slideIndex,
      targetIndex: args?.targetIndex,
      layoutIndex: args?.layoutIndex,
      backgroundColor: args?.backgroundColor,
      filePath: args?.filePath,
      format: args?.format
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
    // schema 声明了 outputPath，就该真的用它——否则调用方传了也不生效（check:params 会抓这种）
    const outputPath = args?.outputPath || previewPath('png');
    const res: any = await callOffice("ppt_capture_slide_preview", { presentationName: args?.presentationName, slideIndex: args?.slideIndex, outputPath });
    // ISS-105/76/86：原来这里用 `res?.error || 'PPT 未生成预览'` 兜底，把宿主回传的真实错误
    // （`hostError` / 两次导出尝试的失败原因）盖成一句无信息量的话，调用方无从排查。
    // 现在优先透传宿主原始信息，并带上桥接侧观察到的事实（返回体、目标路径、文件是否存在）。
    if (!res?.success || !fs.existsSync(outputPath)) {
      const hostDetail = res?.hostError || res?.error;
      // 宿主的 attempts 是对象数组（{path, scale, ok, hostError}），按字符串 join 会打成 [object Object]
      const attempts = Array.isArray(res?.attempts) && res.attempts.length
        ? `；宿主尝试：${(res.attempts as any[])
            .map((a: any) => typeof a === "string"
              ? a
              : `${a.path ?? "?"}（${a.scale ?? "-"}）${a.ok ? "成功" : "失败"}${a.hostError ? `：${a.hostError}` : ""}`)
            .join(" | ")}`
        : "";
      const exported = Array.isArray(res?.exportedPaths) && res.exportedPaths.length
        ? `；宿主实际导出到：${res.exportedPaths.join(" | ")}`
        : "";
      throw new Error(
        `PPT 未生成预览：${hostDetail || "宿主返回 success 但目标文件不存在（未给出原因）"}` +
        `${attempts}${exported}。桥接期望的路径：${outputPath}（${fs.existsSync(outputPath) ? "已存在" : "不存在"}）。`
      );
    }
    return { ...res, imageBase64: fs.readFileSync(outputPath).toString('base64'), imageMimeType: 'image/png' };
};

/** CAP-09 页面尺寸与母版版式。宿主 `PageSetup.SlideWidth/Height` 可读写；
 *  写后逐项读回核对，未生效写 warnings。 */
export const configureSlideLayout: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
  return await callOffice("configure_slide_layout", {
    presentationName: args?.presentationName,
    action: args?.action,
    preset: args?.preset,
    slideWidth: args?.slideWidth,
    slideHeight: args?.slideHeight,
    orientation: args?.orientation,
    templatePath: args?.templatePath,
    layoutName: args?.layoutName,
    layoutIndex: args?.layoutIndex
  });
};
