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
    const result = await callOffice<any>("word_save_document", {
      documentName: args?.documentName,
      filePath: args?.filePath,
      format: args?.format || "docx"
    });

    // 落盘校验：加载项没有文件系统访问，宿主 `ExportAsFixedFormat` 实测**返回成功但不落盘**
    // （问题台账 ISS-69）。这里在桥接侧（Node，有 fs）确认文件真的存在，否则报错而不是返回假成功。
    const requestedPath = typeof args?.filePath === "string" ? args.filePath : null;
    const savedPath = typeof result?.savedPath === "string" ? result.savedPath : null;
    const target = requestedPath || savedPath;
    if (target && target.includes(".")) {
      const { existsSync, statSync } = await import("node:fs");
      const exists = existsSync(target);
      const size = exists ? statSync(target).size : 0;
      if (!exists || size === 0) {
        throw new Error(
          `保存未落盘：${target} ${exists ? "存在但为空文件" : "不存在"}。` +
          `宿主返回成功不代表文件已写出；请检查目标目录权限与宿主导出实现。`
        );
      }
      return { ...result, savedPath: target, verifiedOnDisk: true, fileSizeBytes: size };
    }
    return result;
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
    // ISS-111：宿主回 `hasPdf: true` 但**磁盘上并没有文件**（真机实测：~/.wps-bridge/previews 下
    // 一个 pdf 都没有）。加载项没有文件系统访问，所以落盘校验只能在桥接侧做——
    // 与 ISS-69（word_save_document）同一类问题，当时漏了这条预览路径。
    const produced = typeof res.pdfPath === 'string' && res.pdfPath ? res.pdfPath : outputPath;
    const { existsSync, statSync } = await import("node:fs");
    if (!existsSync(produced)) {
      throw new Error(
        `Word 预览未落盘：宿主返回成功并声称已导出，但 ${produced} 不存在。` +
        `宿主返回：${JSON.stringify({ pdfPath: res.pdfPath, hasPdf: res.hasPdf, hostError: res.hostError }).slice(0, 220)}`
      );
    }
    const size = statSync(produced).size;
    if (size === 0) throw new Error(`Word 预览落盘但为空文件：${produced}`);
    return { ...res, pdfPath: produced, verifiedOnDisk: true, fileSizeBytes: size, message: '已导出 PDF 并确认落盘；此结果为文件路径，请用 PDF 查看器检查。' };
};
