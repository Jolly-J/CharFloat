  // ---------------------------------------------------------------------------
  // word.js — WPS 文字（Word/WPS）全部 RPC 实现
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
  // ==========================================
  // Word (文字) 模块 工业级全套核心操作实现
  // ==========================================

  function wordCreateDocument(app, params) {
    const { templatePath, isVisible } = params || {};
    const wordApp = getWordApp() || app || (typeof Application !== "undefined" ? Application : null);
    if (!wordApp) throw new Error("WPS 文字 (Word) 未就绪");
    const doc = templatePath ? wordApp.Documents.Add(templatePath) : wordApp.Documents.Add();
    try {
      if (doc.Activate) doc.Activate();
    } catch (e) {}
    return {
      success: true,
      documentName: doc.Name,
      fullName: doc.FullName || doc.Name,
      message: `已成功创建新文档: ${doc.Name}`
    };
  }

  function wordSaveDocument(app, params) {
    const { documentName, filePath, format } = params || {};
    const doc = getWordDocument(app, documentName);
    const fmt = (format || "").toLowerCase();

    if (fmt === "pdf" || (filePath && filePath.toLowerCase().endsWith(".pdf"))) {
      if (!filePath) throw new Error("导出 PDF 必须指定 filePath 完整保存路径");
      // 17 = wdExportFormatPDF
      doc.ExportAsFixedFormat(filePath, 17);
      return {
        success: true,
        documentName: doc.Name,
        savedPath: filePath,
        format: "pdf",
        message: `文档 [${doc.Name}] 已成功导出为 PDF: ${filePath}`
      };
    }

    if (filePath) {
      try {
        doc.SaveAs2(filePath);
      } catch (e) {
        doc.SaveAs(filePath);
      }
      return {
        success: true,
        documentName: doc.Name,
        savedPath: filePath,
        format: "docx",
        message: `文档 [${doc.Name}] 已成功另存为: ${filePath}`
      };
    }

    doc.Save();
    return {
      success: true,
      documentName: doc.Name,
      savedPath: doc.FullName || doc.Name,
      format: "docx",
      message: `文档 [${doc.Name}] 已成功原地保存`
    };
  }

  function wordCloseDocument(app, params) {
    const { documentName, saveChanges } = params || {};
    const doc = getWordDocument(app, documentName);
    const docName = doc.Name;
    let saveFlag = 0; // 0 = wdDoNotSaveChanges
    if (saveChanges === true) saveFlag = -1; // -1 = wdSaveChanges
    doc.Close(saveFlag);
    return {
      success: true,
      documentName: docName,
      saveChanges: Boolean(saveChanges),
      message: `文档 [${docName}] 已成功关闭 (saveChanges: ${Boolean(saveChanges)})`
    };
  }

  function wordManageContent(app, params) {
    const { documentName, action, paragraphIndex, paragraphRange, tableIndex } = params || {};
    const doc = getWordDocument(app, documentName);

    switch (action) {
      case "delete_paragraph": {
        if (paragraphRange && Array.isArray(paragraphRange) && paragraphRange.length === 2) {
          const startP = Number(paragraphRange[0]);
          const endP = Number(paragraphRange[1]);
          const startPos = doc.Paragraphs.Item(startP).Range.Start;
          const endPos = doc.Paragraphs.Item(endP).Range.End;
          const r = doc.Range(startPos, endPos);
          r.Delete();
          return {
            success: true,
            documentName: doc.Name,
            action: "delete_paragraph",
            deletedRange: [startP, endP],
            message: `已成功删除段落 P${startP} ~ P${endP}`
          };
        }
        if (!paragraphIndex) throw new Error("delete_paragraph 操作必须提供 paragraphIndex 或 paragraphRange");
        const idx = Number(paragraphIndex);
        const p = doc.Paragraphs.Item(idx);
        p.Range.Delete();
        return {
          success: true,
          documentName: doc.Name,
          action: "delete_paragraph",
          deletedIndex: idx,
          message: `已成功删除第 ${idx} 个段落`
        };
      }
      case "delete_table": {
        const tIdx = Number(tableIndex) || 1;
        if (tIdx > doc.Tables.Count) throw new Error(`表格索引超出范围: 当前仅有 ${doc.Tables.Count} 个表格`);
        const table = doc.Tables.Item(tIdx);
        table.Delete();
        return {
          success: true,
          documentName: doc.Name,
          action: "delete_table",
          deletedTableIndex: tIdx,
          message: `已成功删除第 ${tIdx} 个表格`
        };
      }
      case "clear_all": {
        doc.Content.Delete();
        return {
          success: true,
          documentName: doc.Name,
          action: "clear_all",
          message: `已成功清空文档 [${doc.Name}] 全部内容`
        };
      }
      default:
        throw new Error(`未知的 Word 内容管理操作: ${action} (支持 delete_paragraph, delete_table, clear_all)`);
    }
  }

  // ---------------------------------------------------------------------------
  // Word 只读读回（ISS-44 / ISS-73）：节与页面版式、水印形状、页码域、书签、内容控件、
  // 文档属性、样式清单。全部按"逐项 try/catch + 结果里带 error 字段"实现：
  // 单项读不到不影响整次读回，也不会抛宿主内部错误给调用方。
  // ---------------------------------------------------------------------------
  function wordReadPageSetup(doc) {
    const readSetup = (ps) => ({
      pageWidth: Number(ps.PageWidth),
      pageHeight: Number(ps.PageHeight),
      orientation: Number(ps.Orientation),
      topMargin: Number(ps.TopMargin),
      bottomMargin: Number(ps.BottomMargin),
      leftMargin: Number(ps.LeftMargin),
      rightMargin: Number(ps.RightMargin),
      headerDistance: Number(ps.HeaderDistance),
      footerDistance: Number(ps.FooterDistance),
      differentFirstPage: Number(ps.DifferentFirstPageHeaderFooter) !== 0,
      differentOddEvenPages: Number(ps.OddAndEvenPagesHeaderFooter) !== 0
    });
    const sections = [];
    const count = doc.Sections.Count;
    for (let i = 1; i <= count; i++) {
      const entry = { index: i };
      try {
        const s = doc.Sections.Item(i);
        entry.pageSetup = readSetup(s.PageSetup);
        entry.range = { start: s.Range.Start, end: s.Range.End };
        entry.header = (() => {
          try {
            const h = s.Headers.Item(1);
            return {
              text: (h.Range.Text || "").replace(/[\r\n\x07]/g, ""),
              fieldCount: (() => { try { return h.Range.Fields.Count; } catch (e) { return null; } })(),
              shapeCount: (() => { try { return h.Shapes.Count; } catch (e) { return null; } })(),
              linkToPrevious: (() => { try { return h.LinkToPrevious; } catch (e) { return null; } })()
            };
          } catch (e) { return { error: e.message }; }
        })();
        entry.footer = (() => {
          try {
            const f = s.Footers.Item(1);
            const fields = [];
            try {
              for (let k = 1; k <= f.Range.Fields.Count; k++) {
                const fld = f.Range.Fields.Item(k);
                fields.push({ type: Number(fld.Type), code: (fld.Code ? fld.Code.Text : "").trim() });
              }
            } catch (e) {}
            return {
              text: (f.Range.Text || "").replace(/[\r\n\x07]/g, ""),
              fieldCount: (() => { try { return f.Range.Fields.Count; } catch (e) { return null; } })(),
              fields: fields,
              pageNumberCount: (() => { try { return f.PageNumbers.Count; } catch (e) { return null; } })(),
              linkToPrevious: (() => { try { return f.LinkToPrevious; } catch (e) { return null; } })()
            };
          } catch (e) { return { error: e.message }; }
        })();
        // 该节内的水印形状（正文层 WordArt）；按 Name 以 WordArt 前缀识别
        entry.watermarkShapes = (() => {
          try {
            const list = [];
            const s0 = s.Range.Start, e0 = s.Range.End;
            for (let k = 1; k <= doc.Shapes.Count; k++) {
              const shp = doc.Shapes.Item(k);
              let anchorStart = null;
              try { anchorStart = shp.Anchor.Start; } catch (e) { anchorStart = null; }
              if (anchorStart === null || (anchorStart >= s0 && anchorStart <= e0)) {
                list.push({
                  name: shp.Name,
                  type: Number(shp.Type),
                  text: (() => { try { return shp.TextEffect ? shp.TextEffect.Text : undefined; } catch (e) { return undefined; } })(),
                  left: Number(shp.Left),
                  top: Number(shp.Top),
                  anchoredInSection: anchorStart !== null && anchorStart >= s0 && anchorStart <= e0
                });
              }
            }
            return list;
          } catch (e) { return { error: e.message }; }
        })();
      } catch (e) {
        entry.error = e.message;
      }
      sections.push(entry);
    }
    // 页脚页码格式：从域码推断（PAGE / NUMPAGES 组合）
    for (const entry of sections) {
      if (!entry.footer || entry.footer.error) continue;
      const codes = (entry.footer.fields || []).map(f => f.code);
      const hasPage = codes.some(c => /^PAGE\b/.test(c));
      const hasNumPages = codes.some(c => /^NUMPAGES\b/.test(c));
      entry.pageNumberFormat = hasPage && hasNumPages ? "page_of_pages" : (hasPage ? (codes.some(c => /PAGE/.test(c)) && /-\s*$|^-\s/.test(entry.footer.text) ? "dash" : "simple") : null);
    }
    return { document: readSetup(doc.PageSetup), sections: sections };
  }

  function wordReadBookmarks(doc) {
    const list = [];
    try {
      for (let i = 1; i <= doc.Bookmarks.Count; i++) {
        const b = doc.Bookmarks.Item(i);
        list.push({
          name: b.Name,
          start: b.Range.Start,
          end: b.Range.End,
          empty: (() => { try { return Boolean(b.Empty); } catch (e) { return undefined; } })(),
          text: (b.Range.Text || "").replace(/[\r\n\x07]/g, "")
        });
      }
    } catch (e) {
      return { count: 0, list: list, error: e.message };
    }
    return { count: list.length, list: list };
  }

  function wordReadFields(doc) {
    const list = [];
    try {
      const limit = Math.min(doc.Fields.Count, 100);
      for (let i = 1; i <= limit; i++) {
        const f = doc.Fields.Item(i);
        list.push({
          index: i,
          type: Number(f.Type),
          code: (f.Code ? f.Code.Text : "").trim(),
          result: (f.Result ? f.Result.Text : "").replace(/[\r\n\x07]/g, "")
        });
      }
    } catch (e) {
      return { count: 0, list: list, error: e.message };
    }
    return { count: doc.Fields.Count, list: list };
  }

  function wordReadContentControls(doc) {
    const list = [];
    try {
      const count = doc.ContentControls.Count;
      for (let i = 1; i <= count; i++) {
        const c = doc.ContentControls.Item(i);
        list.push({
          index: i,
          title: (() => { try { return c.Title; } catch (e) { return undefined; } })(),
          tag: (() => { try { return c.Tag; } catch (e) { return undefined; } })(),
          type: (() => { try { return Number(c.Type); } catch (e) { return undefined; } })(),
          text: (() => { try { return (c.Range.Text || "").replace(/[\r\n\x07]/g, ""); } catch (e) { return undefined; } })()
        });
      }
      return { count: count, list: list };
    } catch (e) {
      return { count: 0, list: list, error: e.message };
    }
  }

  function wordReadStyles(doc, nameFilter) {
    try {
      const total = doc.Styles.Count;
      const kw = nameFilter ? String(nameFilter).toLowerCase() : null;
      const inUse = [];
      const matched = [];
      for (let i = 1; i <= total; i++) {
        const st = doc.Styles.Item(i);
        let name = "";
        let iu = false;
        try { name = st.NameLocal || st.Name || ""; } catch (e) {}
        try { iu = Boolean(st.InUse); } catch (e) {}
        if (iu) inUse.push({ name: name, type: (() => { try { return Number(st.Type); } catch (e) { return undefined; } })(), builtIn: (() => { try { return Boolean(st.BuiltIn); } catch (e) { return undefined; } })() });
        if (kw && name.toLowerCase().indexOf(kw) >= 0 && matched.length < 100) {
          matched.push({ name: name, inUse: iu });
        }
      }
      return { count: total, inUseCount: inUse.length, inUse: inUse.slice(0, 100), matched: kw ? matched : undefined };
    } catch (e) {
      return { count: 0, error: e.message };
    }
  }

  function wordReadDocumentProperties(doc) {
    const readBuiltIn = (name) => {
      try {
        return doc.BuiltInDocumentProperties.Item(name).Value;
      } catch (e) {
        return undefined;
      }
    };
    const custom = [];
    try {
      for (let i = 1; i <= doc.CustomDocumentProperties.Count; i++) {
        const p = doc.CustomDocumentProperties.Item(i);
        custom.push({
          name: (() => { try { return p.Name; } catch (e) { return undefined; } })(),
          value: (() => { try { return p.Value; } catch (e) { return undefined; } })()
        });
      }
    } catch (e) {}
    return {
      title: readBuiltIn("Title"),
      subject: readBuiltIn("Subject"),
      author: readBuiltIn("Author"),
      keywords: readBuiltIn("Keywords"),
      category: readBuiltIn("Category"),
      comments: readBuiltIn("Comments"),
      lastAuthor: readBuiltIn("Last Author"),
      revision: readBuiltIn("Revision Number"),
      custom: custom
    };
  }

  function wordReadDocument(app, params) {
    const { documentName, scope, maxParagraphs, includeFormatting, includeTables } = params || {};
    const doc = getWordDocument(app, documentName);
    const outline = [];
    const maxP = Number(maxParagraphs) || 200;

    // 新增的读回 scope（ISS-44 / ISS-73）：与旧 scope 并存，均为**增量**返回，
    // 不改变 full / outline / paragraphs / tables / selection 的既有语义。
    const EXTRA_SCOPES = { layout: 1, styles: 1, properties: 1, bookmarks: 1, content_controls: 1, fields: 1 };
    const extra = Object.prototype.hasOwnProperty.call(EXTRA_SCOPES, scope);

    try {
      const paraCount = doc.Paragraphs.Count;
      for (let i = 1; i <= Math.min(paraCount, 200); i++) {
        const para = doc.Paragraphs.Item(i);
        const text = (para.Range.Text || "").trim();
        if (!text) continue;
        const level = para.OutlineLevel;
        if (level >= 1 && level <= 3) {
          outline.push({ level, text: text.replace(/[\r\n\x07]/g, ""), paragraphIndex: i });
        }
      }
    } catch (e) {}

    let previewText = "";
    const paragraphDetails = [];
    if (scope === "full" || scope === "paragraphs" || !scope) {
      try {
        const count = Math.min(doc.Paragraphs.Count, maxP);
        const snippets = [];
        for (let i = 1; i <= count; i++) {
          const para = doc.Paragraphs.Item(i);
          const t = (para.Range.Text || "").trim();
          const cleanText = t.replace(/[\r\n\x07]/g, "");
          snippets.push(`[P${i}] ${cleanText}`);

          if (includeFormatting) {
            let isBold = false;
            let isItalic = false;
            let fSize = 12;
            let fName = "";
            let align = 0;
            try {
              isBold = para.Range.Font.Bold === true || para.Range.Font.Bold === -1;
              isItalic = para.Range.Font.Italic === true || para.Range.Font.Italic === -1;
              fSize = para.Range.Font.Size;
              fName = para.Range.Font.NameFarEast || para.Range.Font.Name;
              align = para.Format.Alignment;
            } catch (fe) {}

            paragraphDetails.push({
              index: i,
              text: cleanText,
              bold: isBold,
              italic: isItalic,
              fontSize: fSize,
              fontName: fName,
              alignment: align,
              outlineLevel: para.OutlineLevel
            });
          }
        }
        previewText = snippets.join("\n");
      } catch (e) {}
    }

    const tablesSummary = [];
    if (includeTables !== false && doc.Tables) {
      try {
        const tCount = doc.Tables.Count;
        for (let t = 1; t <= tCount; t++) {
          const tbl = doc.Tables.Item(t);
          const rCount = tbl.Rows.Count;
          const cCount = tbl.Columns.Count;
          const previewRows = [];
          for (let r = 1; r <= Math.min(rCount, 3); r++) {
            const rowData = [];
            for (let c = 1; c <= Math.min(cCount, 10); c++) {
              try {
                rowData.push((tbl.Cell(r, c).Range.Text || "").trim().replace(/[\r\n\x07]/g, ""));
              } catch (ce) {
                rowData.push("");
              }
            }
            previewRows.push(rowData);
          }
          tablesSummary.push({
            tableIndex: t,
            rowCount: rCount,
            columnCount: cCount,
            preview: previewRows
          });
        }
      } catch (te) {}
    }

    let selectionText = "";
    try {
      const wordApp = getWordApp() || app;
      if (wordApp.Selection && wordApp.Selection.Range) {
        selectionText = (wordApp.Selection.Range.Text || "").trim().replace(/[\r\n\x07]/g, "");
      }
    } catch (e) {}

    // 读回域（ISS-44 / ISS-73）：只在显式请求时读，避免每次读文档都枚举几百个样式
    const readBack = {};
    if (extra) {
      if (scope === "layout") {
        readBack.layout = wordReadPageSetup(doc);
      } else if (scope === "styles") {
        readBack.styles = wordReadStyles(doc, params.styleNameFilter);
      } else if (scope === "properties") {
        readBack.properties = wordReadDocumentProperties(doc);
      } else if (scope === "bookmarks") {
        readBack.bookmarks = wordReadBookmarks(doc);
      } else if (scope === "content_controls") {
        readBack.contentControls = wordReadContentControls(doc);
      } else if (scope === "fields") {
        readBack.fields = wordReadFields(doc);
      }
    }

    return {
      documentName: doc.Name,
      fullName: doc.FullName || doc.Name,
      paragraphCount: doc.Paragraphs ? doc.Paragraphs.Count : 0,
      tableCount: doc.Tables ? doc.Tables.Count : 0,
      wordCount: doc.Words ? doc.Words.Count : 0,
      outline,
      tables: tablesSummary,
      paragraphDetails: includeFormatting ? paragraphDetails : undefined,
      previewText: scope === "outline" || extra ? undefined : previewText,
      selectionText: selectionText || undefined,
      ...readBack
    };
  }

  // 找到包含指定字符位置的段落 Range（用于 bookmark / after_paragraph 的精确定位）。
  // 说明：宿主 `Paragraphs.Add(targetRange)` 在 WPS for Mac 上**不可靠**——实测传书签范围时
  // 会把内容写到文首、替换掉书签所在文字并销毁该书签（问题台账 ISS-70）。因此这里改成
  // "定位段落 → InsertParagraphAfter 建段 → InsertAfter 写文本"，全部经真实宿主探针验证。
  function wordFindParagraphRangeAt(doc, pos) {
    const count = doc.Paragraphs.Count;
    for (let i = 1; i <= count; i++) {
      const r = doc.Paragraphs.Item(i).Range;
      if (pos >= r.Start && pos <= r.End) return r;
    }
    return null;
  }

  // 在 anchorRange 所指段落之后写入一个新段落，返回新段落 Range 与命中的定位信息。
  function wordInsertParagraphAfterRange(doc, anchorRange, text) {
    const endPos = anchorRange.End - 1;
    const anchor = doc.Range(endPos, endPos);
    anchor.InsertParagraphAfter();
    anchor.InsertAfter(text);
    return { range: anchor, anchorEnd: endPos };
  }

  // 写入单行文本；位置由 location 决定。返回写入后的 Range 供排版与读回使用。
  function wordWriteOneLine(doc, wordApp, location, targetBookmark, paragraphIndex, text, meta) {
    if (location === "bookmark") {
      const bm = doc.Bookmarks.Item(targetBookmark);
      const bmRange = bm.Range;
      meta.bookmarkRange = { start: bmRange.Start, end: bmRange.End };
      const para = wordFindParagraphRangeAt(doc, bmRange.Start);
      if (!para) throw new Error(`书签 [${targetBookmark}] 所在段落无法定位，未写入；请检查文档结构`);
      // 记录书签锚点所属段落，写完后核对书签仍在（宿主可能在插入过程中丢弃书签）
      const anchorParagraphStart = para.Start;
      const res = wordInsertParagraphAfterRange(doc, para, text);
      meta.bookmarkPreserved = doc.Bookmarks.Exists(targetBookmark);
      meta.insertedAtParagraph = (() => {
        try {
          const after = wordFindParagraphRangeAt(doc, res.anchorEnd);
          return after ? { start: after.Start, end: after.End } : null;
        } catch (e) { return null; }
      })();
      meta.anchorParagraphStartBefore = anchorParagraphStart;
      return res.range;
    }
    if (location === "after_paragraph") {
      const idx = Number(paragraphIndex);
      if (!idx || idx < 1 || idx > doc.Paragraphs.Count) {
        throw new Error(`段落索引越界：第 ${paragraphIndex} 段不存在（当前共 ${doc.Paragraphs.Count} 段）`);
      }
      const para = doc.Paragraphs.Item(idx).Range;
      return wordInsertParagraphAfterRange(doc, para, text).range;
    }
    if (location === "selection" && wordApp && wordApp.Selection && wordApp.Selection.Range) {
      const sel = wordApp.Selection.Range;
      const collapsed = sel.Start === sel.End;
      if (collapsed) {
        const r = doc.Range(sel.Start, sel.Start);
        r.InsertAfter(text);
        return r;
      }
      // 非折叠选区：原位替换所选内容（宿主 Range.Text 赋值同样不可靠，用 InsertAfter + Delete 组合）
      const start = sel.Start;
      sel.InsertAfter(text);
      doc.Range(start, start + text.length).InsertAfter("");
      const removed = doc.Range(start + text.length, sel.End + text.length);
      try { removed.Delete(); } catch (e) {}
      const r2 = doc.Range(start, start + text.length);
      return r2;
    }
    if (location === "start") {
      const r = doc.Range(0, 0);
      r.InsertAfter(text);
      return doc.Range(0, Math.min(text.length, Math.max(0, doc.Content.End - 1)));
    }
    if (location === "selection") {
      // 宿主无选区时退回文档开头，与旧实现一致
      const r = doc.Range(0, 0);
      r.InsertAfter(text);
      return doc.Range(0, Math.min(text.length, Math.max(0, doc.Content.End - 1)));
    }
    const endPos = doc.Content.End > 1 ? doc.Content.End - 1 : 0;
    const newPara = doc.Paragraphs.Add(doc.Range(endPos, endPos));
    newPara.Range.Text = text;
    return newPara.Range;
  }

  function wordApplyLineFormat(doc, range, type, formatting) {
    if (type === "heading1") {
      try { range.Style = doc.Styles.Item(-2); } catch (e) { range.Font.Bold = true; range.Font.Size = 22; }
    } else if (type === "heading2") {
      try { range.Style = doc.Styles.Item(-3); } catch (e) { range.Font.Bold = true; range.Font.Size = 16; }
    } else if (type === "heading3") {
      try { range.Style = doc.Styles.Item(-4); } catch (e) { range.Font.Bold = true; range.Font.Size = 14; }
    } else if (type === "bullet_list") {
      try { range.ListFormat.ApplyBulletDefault(); } catch (e) {}
    } else if (type === "quote") {
      try {
        range.Font.Italic = true;
        range.ParagraphFormat.LeftIndent = 28;
      } catch (e) {}
    } else if (type === "code_block") {
      try {
        range.Font.NameFarEast = "Consolas";
        range.Font.NameAscii = "Consolas";
        range.Font.Size = 10.5;
      } catch (e) {}
    }

    if (formatting && typeof formatting === "object") {
      if (formatting.bold !== undefined) range.Font.Bold = Boolean(formatting.bold);
      if (formatting.italic !== undefined) range.Font.Italic = Boolean(formatting.italic);
      if (formatting.fontSizePt !== undefined) range.Font.Size = Number(formatting.fontSizePt);
      if (formatting.fontName) {
        range.Font.NameFarEast = formatting.fontName;
        range.Font.NameAscii = formatting.fontName;
      }
      if (formatting.alignment !== undefined) range.ParagraphFormat.Alignment = Number(formatting.alignment);
      if (formatting.firstLineIndentChars !== undefined) range.ParagraphFormat.CharacterUnitFirstLineIndent = Number(formatting.firstLineIndentChars);
      if (formatting.lineSpacingPt !== undefined) {
        range.ParagraphFormat.LineSpacingRule = 4;
        range.ParagraphFormat.LineSpacing = Number(formatting.lineSpacingPt);
      }
      if (formatting.spaceBeforePt !== undefined) range.ParagraphFormat.SpaceBefore = Number(formatting.spaceBeforePt);
      if (formatting.spaceAfterPt !== undefined) range.ParagraphFormat.SpaceAfter = Number(formatting.spaceAfterPt);
    } else if (!type || type === "paragraph") {
      // 普通正文默认强制消除前文加粗继承
      try { range.Font.Bold = false; } catch (e) {}
    }
  }

  function wordWriteContent(app, params) {
    const { documentName, location, targetBookmark, paragraphIndex, type, content, formatting } = params || {};
    const doc = getWordDocument(app, documentName);
    const wordApp = getWordApp() || app;
    const loc = location || "end";

    // 书签不存在时**显式拒绝**，不再静默落到文末（旧实现会给出 success，调用方无从发现）
    if (loc === "bookmark") {
      if (!targetBookmark) throw new Error("location=bookmark 必须提供 targetBookmark（书签名称）");
      if (!doc.Bookmarks.Exists(targetBookmark)) {
        const names = [];
        try {
          for (let i = 1; i <= doc.Bookmarks.Count; i++) names.push(doc.Bookmarks.Item(i).Name);
        } catch (e) {}
        throw new Error(
          `书签 [${targetBookmark}] 在此文档中不存在，未写入任何内容（拒绝静默落到文末）。` +
          (names.length ? `现有书签：${names.join("、")}` : "当前文档没有任何书签，请先用 wps_execute_script 的 doc.Bookmarks.Add 创建。")
        );
      }
    }
    if (loc === "after_paragraph" && !paragraphIndex) {
      throw new Error("location=after_paragraph 必须提供 paragraphIndex（1-based 段落序号）");
    }

    const lines = Array.isArray(content) ? content : [String(content || "")];
    const nonEmpty = lines.filter(t => t !== undefined && t !== null && String(t) !== "");
    if (nonEmpty.length === 0) {
      throw new Error("content 为空，未写入任何内容");
    }

    const meta = { insertedParagraphs: [], bookmarkPreserved: undefined };
    const writtenRanges = [];
    for (const raw of nonEmpty) {
      const text = String(raw);
      // 先查宿主是否吞字符：旧构建曾在 Paragraphs.Add(targetRange) + Range.Text 路径上吞掉小写字母，
      // 所以写入后按长度读回一次，长度不符就明确报错，而不是返回 success 让调用方踩坑（ISS-67）。
      const range = wordWriteOneLine(doc, wordApp, loc, targetBookmark, paragraphIndex, text, meta);
      wordApplyLineFormat(doc, range, type, formatting);
      writtenRanges.push(range);
      try {
        const readBack = (range.Text || "").replace(/[\r\n\x07]/g, "");
        if (readBack.length !== text.length) {
          throw new Error(
            `写入后读回长度不一致：写入 ${text.length} 个字符，读回 ${readBack.length} 个。` +
            `疑似宿主在 Range 文本赋值时吞字符；请勿重试覆盖，先读回核对，或改用 wps_execute_script 的 Range.InsertAfter。`
          );
        }
      } catch (e) {
        if (/疑似宿主在 Range 文本赋值时吞字符/.test(e.message)) throw e;
      }
      meta.insertedParagraphs.push({
        textLength: text.length,
        start: range.Start,
        end: range.End,
        style: (() => { try { return range.Style ? range.Style.NameLocal : undefined; } catch (e) { return undefined; } })()
      });
    }

    return {
      success: true,
      documentName: doc.Name,
      insertedLines: nonEmpty.length,
      location: loc,
      type: type || "paragraph",
      insertedParagraphs: meta.insertedParagraphs,
      bookmarkRange: loc === "bookmark" ? meta.bookmarkRange : undefined,
      bookmarkPreserved: loc === "bookmark" ? meta.bookmarkPreserved : undefined,
      message: loc === "bookmark"
        ? `已在书签 [${targetBookmark}] 所在段落之后写入 ${nonEmpty.length} 行内容` +
          (meta.bookmarkPreserved === false ? `；注意：宿主在插入过程中丢弃了该书签，需重建` : "")
        : `已成功向 [${doc.Name}] 写入 ${nonEmpty.length} 行内容`
    };
  }

  function wordFormatDocument(app, params) {
    const { documentName, target, paragraphIndex, paragraphRange, searchQuery, searchQueries, preset, fontName, fontSizePt, bold, italic, lineSpacingPt, firstLineIndentChars, spaceBeforePt, spaceAfterPt, margins, alignment } = params || {};
    const doc = getWordDocument(app, documentName);
    const wordApp = getWordApp() || app;

    try {
      if (preset === "gov_standard") {
        doc.PageSetup.TopMargin = 37 * 2.83465;
        doc.PageSetup.BottomMargin = 35 * 2.83465;
        doc.PageSetup.LeftMargin = 28 * 2.83465;
        doc.PageSetup.RightMargin = 26 * 2.83465;
      } else if (margins) {
        if (margins.topMm) doc.PageSetup.TopMargin = margins.topMm * 2.83465;
        if (margins.bottomMm) doc.PageSetup.BottomMargin = margins.bottomMm * 2.83465;
        if (margins.leftMm) doc.PageSetup.LeftMargin = margins.leftMm * 2.83465;
        if (margins.rightMm) doc.PageSetup.RightMargin = margins.rightMm * 2.83465;
      }
    } catch (e) {}

    function applyFormatToPara(para) {
      if (bold !== undefined) para.Range.Font.Bold = Boolean(bold);
      if (italic !== undefined) para.Range.Font.Italic = Boolean(italic);
      if (fontName) {
        para.Range.Font.NameFarEast = fontName;
        para.Range.Font.NameAscii = fontName;
      }
      if (fontSizePt !== undefined) para.Range.Font.Size = Number(fontSizePt);
      if (alignment !== undefined) para.Format.Alignment = Number(alignment);
      if (lineSpacingPt !== undefined) {
        para.Format.LineSpacingRule = 4;
        para.Format.LineSpacing = Number(lineSpacingPt);
      }
      if (firstLineIndentChars !== undefined) {
        para.Format.CharacterUnitFirstLineIndent = Number(firstLineIndentChars);
      }
      if (spaceBeforePt !== undefined) para.Format.SpaceBefore = Number(spaceBeforePt);
      if (spaceAfterPt !== undefined) para.Format.SpaceAfter = Number(spaceAfterPt);
    }

    // 1. 基于搜索词强力加粗/排版 (支持正文与全量表格穿透)
    const queries = Array.isArray(searchQueries) ? searchQueries : (searchQuery ? [searchQuery] : []);
    if (queries.length > 0) {
      const ranges = [doc.Content];
      if (doc.Tables) {
        for (let t = 1; t <= doc.Tables.Count; t++) {
          try { ranges.push(doc.Tables.Item(t).Range); } catch (te) {}
        }
      }
      let formattedCount = 0;
      for (const q of queries) {
        if (!q) continue;
        for (const rng of ranges) {
          try {
            const f = rng.Find;
            f.ClearFormatting();
            f.Text = q;
            f.MatchCase = false;
            f.MatchWholeWord = false;
            f.MatchWildcards = false;
            f.Forward = true;
            f.Wrap = 0; // wdFindStop
            while (f.Execute()) {
              formattedCount++;
              if (bold !== undefined) f.Parent.Font.Bold = Boolean(bold);
              if (italic !== undefined) f.Parent.Font.Italic = Boolean(italic);
              if (fontSizePt !== undefined) f.Parent.Font.Size = Number(fontSizePt);
              if (fontName) {
                f.Parent.Font.NameFarEast = fontName;
                f.Parent.Font.NameAscii = fontName;
              }
            }
          } catch (fe) {}
        }
      }
      return {
        success: true,
        documentName: doc.Name,
        target: "search_matches",
        queries,
        formattedMatches: formattedCount,
        message: `已成功为 ${queries.length} 个关键词匹配项 (${formattedCount} 处) 应用排版`
      };
    }

    if (target === "paragraph" || paragraphIndex) {
      const pIdx = Number(paragraphIndex || 1);
      const p = doc.Paragraphs.Item(pIdx);
      applyFormatToPara(p);
      return {
        success: true,
        documentName: doc.Name,
        target: `paragraph_${pIdx}`,
        message: `已成功格式化第 ${pIdx} 个段落`
      };
    }

    if (target === "range" || (paragraphRange && Array.isArray(paragraphRange))) {
      const startP = Number(paragraphRange[0]);
      const endP = Number(paragraphRange[1]);
      for (let i = startP; i <= endP; i++) {
        try {
          applyFormatToPara(doc.Paragraphs.Item(i));
        } catch (e) {}
      }
      return {
        success: true,
        documentName: doc.Name,
        target: `range_${startP}_${endP}`,
        message: `已成功格式化段落 P${startP} ~ P${endP}`
      };
    }

    if (target === "selection" && wordApp.Selection && wordApp.Selection.Range) {
      const rng = wordApp.Selection.Range;
      if (bold !== undefined) rng.Font.Bold = Boolean(bold);
      if (italic !== undefined) rng.Font.Italic = Boolean(italic);
      if (fontName) {
        rng.Font.NameFarEast = fontName;
        rng.Font.NameAscii = fontName;
      }
      if (fontSizePt !== undefined) rng.Font.Size = Number(fontSizePt);
      return {
        success: true,
        documentName: doc.Name,
        target: "selection",
        message: `已成功格式化当前选区文字`
      };
    }

    const paraCount = doc.Paragraphs.Count;
    for (let i = 1; i <= paraCount; i++) {
      try {
        const para = doc.Paragraphs.Item(i);
        const level = para.OutlineLevel;

        if (preset === "gov_standard") {
          if (level >= 10 || level === 0) {
            para.Range.Font.NameFarEast = "仿宋_GB2312";
            para.Range.Font.NameAscii = "仿宋_GB2312";
            para.Range.Font.Size = 16;
            para.Format.LineSpacingRule = 4;
            para.Format.LineSpacing = 28;
            para.Format.CharacterUnitFirstLineIndent = 2;
            para.Format.SpaceBefore = 0;
            para.Format.SpaceAfter = 0;
          } else if (level === 1) {
            para.Range.Font.NameFarEast = "黑体";
            para.Range.Font.Size = 16;
            para.Range.Font.Bold = false;
            para.Format.LineSpacingRule = 4;
            para.Format.LineSpacing = 28;
            para.Format.CharacterUnitFirstLineIndent = 2;
          } else if (level === 2) {
            para.Range.Font.NameFarEast = "楷体_GB2312";
            para.Range.Font.Size = 16;
            para.Range.Font.Bold = false;
            para.Format.LineSpacingRule = 4;
            para.Format.LineSpacing = 28;
            para.Format.CharacterUnitFirstLineIndent = 2;
          }
        } else if (preset === "business_modern") {
          para.Range.Font.NameFarEast = "微软雅黑";
          para.Range.Font.NameAscii = "Segoe UI";
          if (level >= 10 || level === 0) {
            para.Range.Font.Size = 11;
            para.Format.LineSpacingRule = 5;
            para.Format.SpaceAfter = 6;
          }
        } else {
          applyFormatToPara(para);
        }
      } catch (e) {}
    }

    return {
      success: true,
      documentName: doc.Name,
      preset: preset || "custom",
      formattedParagraphs: paraCount,
      message: `已成功对 [${doc.Name}] 应用 [${preset || "自定义"}] 排版规范`
    };
  }

  function wordInsertTableOfContents(app, params) {
    const { documentName, upperHeadingLevel, lowerHeadingLevel, insertLocation, includePageNumbers } = params || {};
    const doc = getWordDocument(app, documentName);
    let targetRange = null;

    if (insertLocation === "selection" && app.Selection) {
      targetRange = app.Selection.Range;
    } else {
      targetRange = doc.Range(0, 0);
    }

    const upper = Number(upperHeadingLevel) || 1;
    const lower = Number(lowerHeadingLevel) || 3;

    try {
      doc.TablesOfContents.Add(
        targetRange,
        true,
        upper,
        lower,
        true,
        undefined,
        true,
        includePageNumbers !== false
      );
    } catch (e) {
      const p = doc.Paragraphs.Add(targetRange);
      p.Range.Text = "【目录】\n";
      p.Range.Font.Bold = true;
    }

    return {
      success: true,
      documentName: doc.Name,
      headingLevels: `${upper}-${lower}`,
      message: `已成功在 [${doc.Name}] 插入标准目录`
    };
  }

  function wordManageTable(app, params) {
    const { documentName, action, tableIndex, rows, columns, data, stylePreset, repeatHeader, mergeRange, cellRow, cellColumn, cellFormat, rowIndex, columnIndex } = params || {};
    const doc = getWordDocument(app, documentName);

    if (action === "inspect") {
      const idx = Number(tableIndex) || 1;
      if (idx > doc.Tables.Count) throw new Error(`表格索引超出范围: 当前仅有 ${doc.Tables.Count} 个表格`);
      const table = doc.Tables.Item(idx);
      const rCount = table.Rows.Count;
      const cCount = table.Columns.Count;
      const matrix = [];
      for (let r = 1; r <= rCount; r++) {
        const rowData = [];
        for (let c = 1; c <= cCount; c++) {
          try {
            rowData.push((table.Cell(r, c).Range.Text || "").trim().replace(/[\r\n\x07]/g, ""));
          } catch (ce) {
            rowData.push("");
          }
        }
        matrix.push(rowData);
      }
      return {
        success: true,
        documentName: doc.Name,
        tableIndex: idx,
        rowCount: rCount,
        columnCount: cCount,
        data: matrix,
        message: `表格 ${idx} 结构透视: ${rCount} 行 × ${cCount} 列`
      };
    }

    if (action === "insert") {
      const dataRows = Array.isArray(data) && data.length > 0 ? data.length : (Number(rows) || 3);
      const dataCols = Array.isArray(data) && data[0] && Array.isArray(data[0]) ? data[0].length : (Number(columns) || 3);

      const targetRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
      const table = doc.Tables.Add(targetRange, dataRows, dataCols);

      if (Array.isArray(data)) {
        for (let r = 0; r < data.length; r++) {
          for (let c = 0; c < data[r].length; c++) {
            try {
              table.Cell(r + 1, c + 1).Range.Text = String(data[r][c] !== undefined ? data[r][c] : "");
            } catch (e) {}
          }
        }
      }

      if (stylePreset === "mckinsey_three_line" || !stylePreset) {
        try {
          table.Borders.Enable = false;
          table.Borders.Item(-1).LineStyle = 1;
          table.Borders.Item(-1).LineWidth = 12;
          table.Borders.Item(-3).LineStyle = 1;
          table.Borders.Item(-3).LineWidth = 12;
          if (table.Rows.Count > 1) {
            table.Rows.Item(1).Borders.Item(-3).LineStyle = 1;
            table.Rows.Item(1).Borders.Item(-3).LineWidth = 6;
            table.Rows.Item(1).Range.Font.Bold = true;
          }
        } catch (e) {}
      }

      if (repeatHeader !== false && table.Rows.Count > 0) {
        try {
          table.Rows.Item(1).HeadingFormat = true;
        } catch (e) {}
      }

      return {
        success: true,
        documentName: doc.Name,
        tableIndex: doc.Tables.Count,
        rows: dataRows,
        columns: dataCols,
        style: stylePreset || "mckinsey_three_line",
        message: `已成功在 [${doc.Name}] 插入 ${dataRows}行 × ${dataCols}列 专业三线表`
      };
    }

    if (action === "update_data" || action === "write_matrix") {
      const idx = Number(tableIndex) || 1;
      if (idx > doc.Tables.Count) throw new Error(`表格索引超出范围: 当前仅有 ${doc.Tables.Count} 个表格`);
      const table = doc.Tables.Item(idx);

      if (Array.isArray(data)) {
        const reqRows = data.length;
        // 动态扩容行数，绝不静默丢弃任何行
        while (table.Rows.Count < reqRows) {
          table.Rows.Add();
        }

        for (let r = 0; r < data.length; r++) {
          const rowArr = data[r];
          if (!Array.isArray(rowArr)) continue;
          for (let c = 0; c < rowArr.length; c++) {
            try {
              const cell = table.Cell(r + 1, c + 1);
              cell.Range.Text = String(rowArr[c] !== undefined && rowArr[c] !== null ? rowArr[c] : "");
              // 消除意外的粗体继承污染
              if (r > 0) {
                cell.Range.Font.Bold = false;
              }
            } catch (e) {}
          }
        }
      }
      return {
        success: true,
        documentName: doc.Name,
        tableIndex: idx,
        rowCount: table.Rows.Count,
        columnCount: table.Columns.Count,
        message: `表格 ${idx} 数据已安全更新 (${data ? data.length : 0} 行)`
      };
    }

    if (action === "format_cell") {
      const idx = Number(tableIndex) || 1;
      const table = doc.Tables.Item(idx);
      const r = Number(cellRow || rowIndex || 1);
      const c = Number(cellColumn || columnIndex || 1);
      const cell = table.Cell(r, c);
      if (cellFormat && typeof cellFormat === "object") {
        if (cellFormat.bold !== undefined) cell.Range.Font.Bold = Boolean(cellFormat.bold);
        if (cellFormat.fontSizePt !== undefined) cell.Range.Font.Size = Number(cellFormat.fontSizePt);
        if (cellFormat.fontName) cell.Range.Font.NameFarEast = cellFormat.fontName;
        if (cellFormat.backgroundColor) {
          const bgr = hexToExcelColor(cellFormat.backgroundColor);
          if (bgr !== null) cell.Shading.BackgroundPatternColor = bgr;
        }
      }
      return {
        success: true,
        documentName: doc.Name,
        tableIndex: idx,
        cell: [r, c],
        message: `表格 ${idx} 单元格 [${r}, ${c}] 格式已更新`
      };
    }

    if (action === "add_row") {
      const idx = Number(tableIndex) || 1;
      const table = doc.Tables.Item(idx);
      table.Rows.Add();
      return { success: true, documentName: doc.Name, tableIndex: idx, rowCount: table.Rows.Count, message: `表格 ${idx} 已追加新行` };
    }

    if (action === "delete_row") {
      const idx = Number(tableIndex) || 1;
      const table = doc.Tables.Item(idx);
      const r = Number(rowIndex || table.Rows.Count);
      table.Rows.Item(r).Delete();
      return { success: true, documentName: doc.Name, tableIndex: idx, rowCount: table.Rows.Count, message: `表格 ${idx} 第 ${r} 行已删除` };
    }

    if (action === "merge_cells" && mergeRange) {
      const idx = Number(tableIndex) || 1;
      const table = doc.Tables.Item(idx);
      const startCell = table.Cell(mergeRange.startRow, mergeRange.startCol);
      const endCell = table.Cell(mergeRange.endRow, mergeRange.endCol);
      startCell.Merge(endCell);
      return { success: true, documentName: doc.Name, tableIndex: idx, message: `表格 ${idx} 单元格已合并` };
    }

    throw new Error(`未知的 Word 表格操作: ${action} (支持 inspect, insert, update_data, write_matrix, format_cell, add_row, delete_row, merge_cells)`);
  }

  function wordReviewAndComments(app, params) {
    const { documentName, action, commentText, author } = params || {};
    const doc = getWordDocument(app, documentName);
    const wordApp = getWordApp() || app;

    switch (action) {
      case "enable_track_changes": {
        doc.TrackRevisions = true;
        return { success: true, documentName: doc.Name, trackRevisions: true, message: `文档 [${doc.Name}] 已开启修订记录模式` };
      }
      case "disable_track_changes": {
        doc.TrackRevisions = false;
        return { success: true, documentName: doc.Name, trackRevisions: false, message: `文档 [${doc.Name}] 已关闭修订记录模式` };
      }
      case "accept_all_revisions": {
        doc.AcceptAllRevisions();
        return { success: true, documentName: doc.Name, message: `已接受文档 [${doc.Name}] 中的全部修订` };
      }
      case "reject_all_revisions": {
        doc.RejectAllRevisions();
        return { success: true, documentName: doc.Name, message: `已拒绝文档 [${doc.Name}] 中的全部修订` };
      }
      case "add_comment": {
        if (!commentText) throw new Error("add_comment 操作必须提供 commentText");
        const range = (wordApp.Selection && wordApp.Selection.Range.Text) ? wordApp.Selection.Range : doc.Range(0, 0);
        const comment = doc.Comments.Add(range, commentText);
        if (author) comment.Author = author;
        return { success: true, documentName: doc.Name, commentId: doc.Comments.Count, message: `已成功添加批注: "${commentText}"` };
      }
      case "list_comments": {
        const list = [];
        const count = doc.Comments ? doc.Comments.Count : 0;
        for (let i = 1; i <= count; i++) {
          const c = doc.Comments.Item(i);
          list.push({
            index: i,
            author: c.Author,
            text: (c.Range.Text || "").trim(),
            scopeText: (c.Scope.Text || "").trim()
          });
        }
        return { success: true, documentName: doc.Name, total: count, comments: list };
      }
      default:
        throw new Error(`未知的 Word 审阅操作: ${action}`);
    }
  }

  // 页码格式 → 页脚域组合。页码用真正的 PAGE / NUMPAGES 域写入，不用纯文本（纯文本不会随页变化）。
  // 实测（WPS for Mac 12.0）：`doc.Fields.Add(range, 33)` → code " PAGE "；`33`=wdFieldPage，`26`=wdFieldNumPages。
  function wordSetPageNumberFormat(doc, section, format, label) {
    const fmt = String(format || "").toLowerCase();
    let footer;
    try {
      footer = section.Footers.Item(1);
      if (footer.LinkToPrevious) footer.LinkToPrevious = false;
    } catch (e) {
      return { ok: false, error: `无法访问页脚: ${e.message}` };
    }
    try {
      footer.Range.Text = "";
      const anchor = footer.Range;
      if (fmt === "dash") {
        anchor.InsertAfter("- ");
        const r = doc.Range(anchor.End - 1, anchor.End - 1);
        doc.Fields.Add(r, 33, "", false);
        const tail = doc.Range(r.End, r.End);
        tail.InsertAfter(" -");
      } else if (fmt === "page_of_pages") {
        const r = doc.Range(anchor.End - 1, anchor.End - 1);
        doc.Fields.Add(r, 33, "", false);
        const mid = doc.Range(r.End, r.End);
        mid.InsertAfter(" / ");
        const r2 = doc.Range(mid.End, mid.End);
        doc.Fields.Add(r2, 26, "", false);
      } else if (fmt === "simple") {
        const r = doc.Range(anchor.End - 1, anchor.End - 1);
        doc.Fields.Add(r, 33, "", false);
      } else {
        return { ok: false, error: `未知的 pageNumberFormat: ${format}（支持 dash | simple | page_of_pages）` };
      }
      const readText = (footer.Range.Text || "").replace(/[\r\n\x07]/g, "");
      const fieldCount = (() => { try { return footer.Range.Fields.Count; } catch (e) { return null; } })();
      if (!fieldCount) {
        return { ok: false, error: `页脚域写入后读回为 0 个域（section ${label}），页码可能未生效`, footerText: readText };
      }
      return { ok: true, format: fmt, footerText: readText, fieldCount: fieldCount };
    } catch (e) {
      return { ok: false, error: `写入页码域失败: ${e.message}` };
    }
  }

  // 水印：优先尝试"页眉层"（跨页可见），失败则回退正文层并给出告警。
  //
  // 已实测的宿主事实（WPS for Mac 12.0 / 12.1.28496）：
  //   - `section.Headers.Item(1).Shapes.AddTextEffect(...)` **会静默把形状加到正文层**：
  //     调用后 `header.Shapes.Count` 恒为 0、`doc.Shapes.Count` +1（同一形状对象）；
  //   - `header.Range.ShapeRange.AddTextEffect` 是 undefined（"is not a function"）；
  //   - `Range.InsertXML`（VML `<w:pict>`）对页眉 story 无效，页眉 XML 长度不变；
  //   - `doc.Shapes.AddTextEffect` 落正文层：正文层浮动图形**只在第 1 页渲染**，不是"每页可见"。
  // 因此这里如实返回 placement，并把"跨页水印需另想办法"写进 warnings。
  function wordAddWatermarkToSection(doc, section, text, colorHex, pageSetup, warnings, sectionIndex) {
    const fontSize = 54;
    const shapeColor = hexToExcelColor(colorHex || "#C0C0C0") || 0xc0c0c0;
    let shape = null;
    let placement = "body";

    // 路径 1：页眉层
    try {
      const header = section.Headers.Item(1);
      header.Shapes.AddTextEffect(0, text, "Microsoft YaHei", fontSize, false, false, 0, 0);
      let headerCount = 0;
      try { headerCount = header.Shapes.Count; } catch (e) {}
      if (headerCount > 0) {
        placement = "header";
        try { shape = header.Shapes.Item(headerCount); } catch (e) {}
      }
    } catch (e) {
      warnings.push(`第 ${sectionIndex} 节页眉层水印写入失败：${e.message}`);
    }

    // 路径 2：正文层（页眉层不可用时的回退；宿主会把页眉 Shapes 的写入落到这里）
    if (!shape) {
      try {
        shape = doc.Shapes.AddTextEffect(0, text, "Microsoft YaHei", fontSize, false, false, 0, 0);
        placement = "body";
      } catch (e) {
        warnings.push(`第 ${sectionIndex} 节水印创建失败：${e.message}`);
        return { ok: false, error: e.message };
      }
    }

    let geometry = null;
    try {
      shape.Rotation = -315;
      try { shape.Fill.Transparency = 0.85; } catch (e) {}
      try { shape.Fill.ForeColor.RGB = shapeColor; } catch (e) {}
      try { shape.Line.Visible = false; } catch (e) {}
      try { shape.WrapFormat.Type = 3; } catch (e) {}
      // 居中：先把版式设为"相对页面"，再按页面尺寸居中（属性名在 WPS 上为 Range.ParagraphFormat 同族对象）
      const pw = pageSetup.pageWidth, ph = pageSetup.pageHeight;
      const w = Number(shape.Width) || 0, h = Number(shape.Height) || 0;
      shape.Left = Math.round(((pw - w) / 2) * 100) / 100;
      shape.Top = Math.round(((ph - h) / 2) * 100) / 100;
      geometry = { left: shape.Left, top: shape.Top, width: shape.Width, height: shape.Height, rotation: shape.Rotation };
    } catch (e) {
      warnings.push(`第 ${sectionIndex} 节水印属性设置部分失败：${e.message}`);
    }

    return {
      ok: true,
      sectionIndex: sectionIndex,
      placement: placement,
      shapeName: (() => { try { return shape.Name; } catch (e) { return undefined; } })(),
      geometry: geometry
    };
  }

  function wordPageLayoutAndWatermark(app, params) {
    const { documentName, headerText, footerText, pageNumberFormat, differentFirstPage, differentOddEvenPages, watermarkText, watermarkColor } = params || {};
    const doc = getWordDocument(app, documentName);
    const warnings = [];
    const sectionCount = doc.Sections.Count;

    if (headerText === undefined && footerText === undefined && pageNumberFormat === undefined && watermarkText === undefined &&
        differentFirstPage === undefined && differentOddEvenPages === undefined) {
      throw new Error("headerText / footerText / pageNumberFormat / watermarkText / differentFirstPage / differentOddEvenPages 至少传一项，否则本调用不产生任何变化");
    }

    // 文档级选项
    if (differentFirstPage !== undefined) doc.PageSetup.DifferentFirstPageHeaderFooter = differentFirstPage;
    if (differentOddEvenPages !== undefined) doc.PageSetup.OddAndEvenPagesHeaderFooter = differentOddEvenPages;

    const appliedSections = [];
    for (let i = 1; i <= sectionCount; i++) {
      const section = doc.Sections.Item(i);
      const entry = { sectionIndex: i, header: false, footer: false, pageNumberFormat: null, watermark: null };

      if (headerText !== undefined) {
        try {
          const header = section.Headers.Item(1);
          header.Range.Text = String(headerText);
          entry.header = true;
          entry.headerText = (header.Range.Text || "").replace(/[\r\n\x07]/g, "");
        } catch (e) {
          warnings.push(`第 ${i} 节页眉写入失败：${e.message}`);
        }
      }
      if (footerText !== undefined) {
        try {
          const footer = section.Footers.Item(1);
          footer.Range.Text = String(footerText);
          entry.footer = true;
          entry.footerText = (footer.Range.Text || "").replace(/[\r\n\x07]/g, "");
        } catch (e) {
          warnings.push(`第 ${i} 节页脚写入失败：${e.message}`);
        }
      }
      if (pageNumberFormat !== undefined) {
        const pn = wordSetPageNumberFormat(doc, section, pageNumberFormat, i);
        if (pn.ok) {
          entry.pageNumberFormat = { format: pn.format, footerText: pn.footerText, fieldCount: pn.fieldCount };
        } else {
          warnings.push(`第 ${i} 节页码写入失败：${pn.error}`);
        }
      }
      if (watermarkText) {
        const ps = (() => {
          try {
            return { pageWidth: Number(doc.PageSetup.PageWidth) || 0, pageHeight: Number(doc.PageSetup.PageHeight) || 0 };
          } catch (e) { return { pageWidth: 0, pageHeight: 0 }; }
        })();
        const wm = wordAddWatermarkToSection(doc, section, String(watermarkText), watermarkColor, ps, warnings, i);
        if (wm.ok) entry.watermark = wm;
      }

      appliedSections.push(entry);
    }

    if (watermarkText && sectionCount > 1) {
      warnings.push(
        `水印已按 ${sectionCount} 个节分别写入，但宿主 WPS for Mac 无法把形状放进页眉层` +
        `（Headers.Shapes 的写入会静默落到正文层），因此水印仍是**正文层浮动图形**：` +
        `实测只在第 1 页渲染，不是"每页可见"。跨页水印需在 Word 内手动插入（插入 → 水印），` +
        `或由调用方在 Windows/COM 通道用 Section.Headers.Shapes 处理。`
      );
    } else if (watermarkText) {
      warnings.push(
        `水印落在正文层（宿主 WPS for Mac 的 Headers.Shapes 写入会静默落到正文层）：` +
        `正文层浮动图形只在第 1 页渲染，不是"每页可见"。跨页水印需在 Word 内手动插入（插入 → 水印），` +
        `或改用 Windows/COM 通道。`
      );
    }

    return {
      success: true,
      documentName: doc.Name,
      sectionCount: sectionCount,
      appliedSections: appliedSections,
      header: headerText,
      footer: footerText,
      pageNumberFormat: pageNumberFormat || undefined,
      watermark: watermarkText || undefined,
      warnings: warnings,
      message:
        `已更新 [${doc.Name}] 页面版式：${sectionCount} 个节` +
        (headerText !== undefined ? "，页眉" : "") +
        (footerText !== undefined ? "，页脚" : "") +
        (pageNumberFormat !== undefined ? `，页码(${pageNumberFormat})` : "") +
        (watermarkText ? "，水印" : "") +
        (warnings.length ? `；有 ${warnings.length} 条告警，请逐条查看 warnings` : "")
    };
  }

  function wordFindAndReplace(app, params) {
    const { documentName, searchQuery, replaceText, matchCase, matchWholeWord, useWildcards, scope, replaceFormatting } = params || {};
    if (!searchQuery) throw new Error("缺少必要参数: searchQuery");
    const doc = getWordDocument(app, documentName);
    const wordApp = getWordApp() || app;

    let targetRanges = [];
    if (scope === "selection" && wordApp && wordApp.Selection && wordApp.Selection.Range) {
      targetRanges = [wordApp.Selection.Range];
    } else {
      targetRanges = [doc.Content];
      if (doc.Tables) {
        for (let t = 1; t <= doc.Tables.Count; t++) {
          try { targetRanges.push(doc.Tables.Item(t).Range); } catch (te) {}
        }
      }
    }

    let matchCount = 0;
    let replacedAny = false;
    const errors = [];
    const hasFmt = Boolean(replaceFormatting && typeof replaceFormatting === "object");

    for (const rng of targetRanges) {
      try {
        const findObj = rng.Find;
        findObj.ClearFormatting();
        findObj.Text = searchQuery;
        findObj.MatchCase = Boolean(matchCase);
        findObj.MatchWholeWord = Boolean(matchWholeWord);
        findObj.MatchWildcards = Boolean(useWildcards);
        findObj.Forward = true;
        findObj.Wrap = 0; // wdFindStop

        if (replaceText !== undefined) {
          findObj.Replacement.ClearFormatting();
          findObj.Replacement.Text = replaceText;

          if (hasFmt) {
            if (replaceFormatting.bold !== undefined) findObj.Replacement.Font.Bold = Boolean(replaceFormatting.bold);
            if (replaceFormatting.italic !== undefined) findObj.Replacement.Font.Italic = Boolean(replaceFormatting.italic);
            if (replaceFormatting.fontSizePt !== undefined) findObj.Replacement.Font.Size = Number(replaceFormatting.fontSizePt);
            if (replaceFormatting.fontName) {
              findObj.Replacement.Font.NameFarEast = replaceFormatting.fontName;
              findObj.Replacement.Font.NameAscii = replaceFormatting.fontName;
            }
          }

          // 第 9 个参数 Format 传入 hasFmt，使 Replacement.Font 格式化必定生效。
          // 返回值必须看：原实现忽略了它并无条件 matchCount++，导致"没命中"也报成功（问题台账 ISS-68）。
          const did = findObj.Execute(
            searchQuery,
            Boolean(matchCase),
            Boolean(matchWholeWord),
            Boolean(useWildcards),
            false,
            false,
            true,
            1,
            hasFmt,
            replaceText,
            2 // wdReplaceAll
          );
          if (did) {
            matchCount++;
            replacedAny = true;
          }
        } else {
          // 纯查找 / 查找并格式化：**原实现缺这一支**，所以只查找时 matchCount 恒为 0、
          // 文案却写"已找到并应用格式化"。这里真正遍历计数（带上限防死循环）。
          let guard = 0;
          while (findObj.Execute() && guard++ < 5000) {
            matchCount++;
            if (hasFmt) {
              if (replaceFormatting.bold !== undefined) findObj.Parent.Font.Bold = Boolean(replaceFormatting.bold);
              if (replaceFormatting.italic !== undefined) findObj.Parent.Font.Italic = Boolean(replaceFormatting.italic);
              if (replaceFormatting.fontSizePt !== undefined) findObj.Parent.Font.Size = Number(replaceFormatting.fontSizePt);
              if (replaceFormatting.fontName) {
                findObj.Parent.Font.NameFarEast = replaceFormatting.fontName;
                findObj.Parent.Font.NameAscii = replaceFormatting.fontName;
              }
            }
          }
        }
      } catch (e) {
        // 不再静默吞异常：把原始错误带回给调用方
        errors.push(e.message);
      }
    }

    if (replaceText !== undefined) {
      return {
        success: true,
        documentName: doc.Name,
        searchQuery,
        replaceText,
        action: replacedAny ? "replaced_all" : "no_match",
        matchCount,
        errors: errors.length ? errors : undefined,
        message: replacedAny
          ? `已将 [${doc.Name}] 中的 "${searchQuery}" 全文穿透替换为 "${replaceText}"（命中 ${matchCount} 段）`
          : `未在 [${doc.Name}] 中找到 "${searchQuery}"，未做任何替换`
      };
    }

    return {
      success: true,
      documentName: doc.Name,
      searchQuery,
      matchCount,
      errors: errors.length ? errors : undefined,
      message: hasFmt
        ? `已为 "${searchQuery}" 找到并应用格式化（${matchCount} 处）`
        : `在 [${doc.Name}] 中找到 "${searchQuery}" ${matchCount} 处`
    };
  }

  function wordCapturePreview(app, params) {
    const { documentName } = params || {};
    const doc = getWordDocument(app, documentName);
    const tempPdfPath = params.outputPath;
    if (!tempPdfPath) throw new Error("缺少 Bridge 指定的预览输出路径");
    try {
      // wdExportFormatPDF = 17
      doc.ExportAsFixedFormat(tempPdfPath, 17);
      return {
        success: true,
        documentName: doc.Name,
        pdfPath: tempPdfPath,
        hasPdf: true,
        message: `已成功导出 Word 文档页面快照 PDF: ${tempPdfPath}`
      };
    } catch (e) {
      log(`Word 页面导出异常: ${e.message}`);
      return {
        success: false,
        documentName: doc.Name,
        error: e.message
      };
    }
  }
