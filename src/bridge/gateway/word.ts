/**
 * Word（文字）类处理器（P3.1：从 gateway.ts 的 executeTool switch 原样搬迁，逻辑未改写）。
 *
 * 共享依赖（callOffice / auditStore / bridgeServer / TargetLockStore / requestContext /
 * previewPath / MsOfficeDriver / currentHost 等）一律经 GatewayContext 传入；
 * 本模块不 import 这些执行器，避免与门面产生新的隐式耦合或回环。
 *
 * 每个处理器的前三行固定为：导出签名、ctx 依赖解构、以及原分支正文的第一行；
 * 正文承接原 switch 分支的内容，仅做逐行等量左移 4 个空格的纯缩进变换。
 */
import type { Handler } from "./types.js";

export const createDocument: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("word_create_document", {
      templatePath: args?.templatePath,
      isVisible: args?.isVisible ?? true
    });
};

export const saveDocument: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("word_save_document", {
      documentName: args?.documentName,
      filePath: args?.filePath,
      format: args?.format || "docx"
    });
};

export const closeDocument: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("word_close_document", {
      documentName: args?.documentName,
      saveChanges: args?.saveChanges ?? false
    });
};

export const manageContent: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.action) throw new Error("缺少必要参数: action ('delete_paragraph'|'delete_table'|'clear_all')");
    return await callOffice("word_manage_content", {
      documentName: args?.documentName,
      action: args?.action,
      paragraphIndex: args?.paragraphIndex,
      paragraphRange: args?.paragraphRange,
      tableIndex: args?.tableIndex
    });
};

export const readDocument: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("word_read_document", {
      documentName: args?.documentName,
      scope: args?.scope || "full",
      maxParagraphs: args?.maxParagraphs ?? 200,
      includeFormatting: args?.includeFormatting ?? true,
      includeTables: args?.includeTables ?? true
    });
};

export const writeContent: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.content) throw new Error("缺少必要参数: content");
    return await callOffice("word_write_content", {
      documentName: args?.documentName,
      location: args?.location || "end",
      targetBookmark: args?.targetBookmark,
      paragraphIndex: args?.paragraphIndex,
      type: args?.type || "paragraph",
      content: args?.content,
      formatting: args?.formatting
    });
};

export const formatDocument: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("word_format_document", {
      documentName: args?.documentName,
      target: args?.target,
      paragraphIndex: args?.paragraphIndex,
      paragraphRange: args?.paragraphRange,
      searchQuery: args?.searchQuery,
      searchQueries: args?.searchQueries,
      preset: args?.preset,
      fontName: args?.fontName,
      fontSizePt: args?.fontSizePt,
      bold: args?.bold,
      italic: args?.italic,
      lineSpacingPt: args?.lineSpacingPt,
      firstLineIndentChars: args?.firstLineIndentChars,
      spaceBeforePt: args?.spaceBeforePt,
      spaceAfterPt: args?.spaceAfterPt,
      margins: args?.margins,
      alignment: args?.alignment
    });
};

export const insertTableOfContents: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("word_insert_table_of_contents", {
      documentName: args?.documentName,
      upperHeadingLevel: args?.upperHeadingLevel ?? 1,
      lowerHeadingLevel: args?.lowerHeadingLevel ?? 3,
      insertLocation: args?.insertLocation || "start",
      includePageNumbers: args?.includePageNumbers ?? true
    });
};

export const manageTable: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.action) throw new Error("缺少必要参数: action ('inspect'|'insert'|'update_data'|'write_matrix'|'format_cell'|'add_row'|'delete_row'|'merge_cells')");
    return await callOffice("word_manage_table", {
      documentName: args?.documentName,
      action: args?.action,
      tableIndex: args?.tableIndex,
      rows: args?.rows,
      columns: args?.columns,
      data: args?.data,
      stylePreset: args?.stylePreset || "mckinsey_three_line",
      repeatHeader: args?.repeatHeader ?? true,
      mergeRange: args?.mergeRange,
      cellRow: args?.cellRow,
      cellColumn: args?.cellColumn,
      cellFormat: args?.cellFormat,
      rowIndex: args?.rowIndex,
      columnIndex: args?.columnIndex
    });
};

export const reviewAndComments: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.action) throw new Error("缺少必要参数: action ('enable_track_changes'|'disable_track_changes'|'accept_all_revisions'|'reject_all_revisions'|'add_comment'|'list_comments')");
    return await callOffice("word_review_and_comments", {
      documentName: args?.documentName,
      action: args?.action,
      commentText: args?.commentText,
      author: args?.author || clientName
    });
};

export const pageLayoutAndWatermark: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await callOffice("word_page_layout_and_watermark", {
      documentName: args?.documentName,
      headerText: args?.headerText,
      footerText: args?.footerText,
      pageNumberFormat: args?.pageNumberFormat,
      differentFirstPage: args?.differentFirstPage,
      differentOddEvenPages: args?.differentOddEvenPages,
      watermarkText: args?.watermarkText,
      watermarkColor: args?.watermarkColor
    });
};

export const findAndReplace: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.searchQuery) throw new Error("缺少必要参数: searchQuery");
    return await callOffice("word_find_and_replace", {
      documentName: args?.documentName,
      searchQuery: args?.searchQuery,
      replaceText: args?.replaceText,
      matchCase: args?.matchCase ?? false,
      matchWholeWord: args?.matchWholeWord ?? false,
      useWildcards: args?.useWildcards ?? false,
      scope: args?.scope || "full",
      replaceFormatting: args?.replaceFormatting
    });
};

export const capturePreview: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const outputPath = previewPath('pdf');
    const res: any = await callOffice("word_capture_preview", { documentName: args?.documentName, outputPath });
    if (!res?.success) throw new Error(res?.error || 'Word 预览导出失败');
    return { ...res, message: '已导出 PDF；此结果为文件路径，请用 PDF 查看器检查。' };
};
