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
        // 该节内（或未报到节归属）的水印形状（正文层 WordArt）；按 Name 以 WordArt 前缀识别。
        // 逐个读 Anchor.Start 在超多形状文档上开销不小，限定最多扫 80 个并如实标注是否截断。
        entry.watermarkShapes = (() => {
          try {
            const list = [];
            const s0 = s.Range.Start, e0 = s.Range.End;
            const totalShapes = doc.Shapes.Count;
            const limit = Math.min(totalShapes, 80);
            for (let k = 1; k <= limit; k++) {
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
            return { totalShapes: totalShapes, scannedShapes: limit, truncated: totalShapes > limit, shapes: list };
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
      // 非折叠选区：原位替换所选内容（宿主 Range.Text 赋值不可靠，用 InsertAfter 把新文本插到选区之后，
      // 再删掉被推到后面的原选区文本；宿主探针实测：选区 "乙 " → "乙 替换后 abc"）
      const start = sel.Start;
      sel.InsertAfter(text);
      try { doc.Range(start + text.length, sel.End + text.length).Delete(); } catch (e) {}
      return doc.Range(start, start + text.length);
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

  // 页码格式 → 页脚页码域。
  //
  // 真实宿主实测（WPS for Mac 12.0 / 12.1.28496）：
  //   - ✅ 唯一可用入口是 `Footers.Item(1).PageNumbers.Add(Alignment, FirstPage)`：
  //     调用后页脚出现真 PAGE 域，可读回 `footer.Range.Fields.Item(1).Code.Text === " PAGE "`；
  //   - ❌ `doc.Fields.Add(range, 33)` 传页脚范围时**静默返回 null、页脚一个域都不加**；
  //   - ❌ `footer.Range.InsertAfter("文字")` 在页脚 story 上被宿主静默丢弃（页脚文本仍为空）；
  //   - ❌ `footer.Range.InsertXML(<w:p>…fldChar/PAGE…)` 同样无效。
  // 因此只有 'simple'（纯页码）能可靠实现；'dash' / 'page_of_pages' 需要额外文字或 NUMPAGES 域，
  // 本机做不到，**显式拒绝**并给出可用替代写法，而不是静默降级成纯页码。
  function wordSetPageNumberFormat(doc, section, format, label) {
    const fmt = String(format || "").toLowerCase();
    if (fmt !== "simple") {
      return {
        ok: false,
        unsupported: true,
        error:
          `pageNumberFormat='${fmt}' 在当前宿主（WPS for Mac）上无法实现。` +
          `宿主只在页脚支持“纯页码”一种写法（可通过 Footers.Item(1).PageNumbers.Add 建真 PAGE 域）；` +
          `页脚的文字拼接与 NUMPAGES 域均被宿主静默丢弃（footer.Range.InsertAfter 与 InsertXML 实测无效）。` +
          `请改用 pageNumberFormat='simple'；需要 "- 1 -" 或 "1 / 5" 这类格式请在 Word 内手动插入页码后自行编辑，` +
          `或由调用方在 Windows/COM 通道处理。`
      };
    }
    let footer;
    try {
      footer = section.Footers.Item(1);
      if (footer.LinkToPrevious) footer.LinkToPrevious = false;
    } catch (e) {
      return { ok: false, error: `无法访问页脚: ${e.message}` };
    }
    try {
      footer.Range.Text = "";
      // wdAlignPageNumberCenter = 1；FirstPage = true
      footer.PageNumbers.Add(1, true);
      const fieldCount = (() => { try { return footer.Range.Fields.Count; } catch (e) { return 0; } })();
      const codes = [];
      try {
        for (let k = 1; k <= fieldCount; k++) codes.push((footer.Range.Fields.Item(k).Code.Text || "").trim());
      } catch (e) {}
      const readText = (footer.Range.Text || "").replace(/[\r\n\x07]/g, "");
      if (!fieldCount || !codes.some(c => /^PAGE\b/.test(c))) {
        return { ok: false, error: `页脚页码域写入后读回为 ${fieldCount} 个域（section ${label}），页码未生效`, footerText: readText };
      }
      return {
        ok: true,
        format: "simple",
        footerText: readText,
        fieldCount: fieldCount,
        fieldCodes: codes,
        pageNumberCount: (() => { try { return footer.PageNumbers.Count; } catch (e) { return null; } })()
      };
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
          entry.pageNumberFormat = { format: pn.format, footerText: pn.footerText, fieldCount: pn.fieldCount, fieldCodes: pn.fieldCodes, pageNumberCount: pn.pageNumberCount };
        } else {
          entry.pageNumberFormat = { format: String(pageNumberFormat).toLowerCase(), applied: false, error: pn.error };
          warnings.push(`第 ${i} 节页码未写入：${pn.error}`);
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

  // ===========================================================================
  // CAP-04：Word 域更新与交叉引用（RPC: word_update_fields）
  //
  // 真机实测（WPS for Mac 12.x，2026-09-22；把 `测试文字文稿.docx` 的内容经
  // `Selection.InsertFile` 导入临时文档后测量，原文件未被改动）：
  //   1. 目录页码**不会**随内容自动刷新——在 TOC 之后插入一个整页分页符后，TOC 内 14 个
  //      PAGEREF 结果仍是 ["3"×6,"4"×3,"5"×5]（陈旧）；`doc.Fields.Update()` 100ms 后
  //      变为 ["4"×6,"5"×3,"6"×5]，逐项 +1。这就是"交付物一改目录页码就全错"的根因。
  //   2. `doc.Fields.Update()` **无弹窗、无交互**（实测 100ms / 63 个域）；
  //      但它的**返回值实测恒为 0**，不能当"更新了几个域"的计数——本工具改为更新前后
  //      逐域结果快照对比，自己数 changedFields。
  //   3. `doc.Fields.Update()` 会**连目录条目一起重建**（新增标题会被收录）；
  //      `TablesOfContents.Item(n).Update()` 同样可重建；`UpdatePageNumbers()` 只刷页码、
  //      不收录新标题（实测新增标题后只调 UpdatePageNumbers，目录里仍然没有该标题）。
  //   4. `doc.Fields` **不含页眉/页脚 story 里的域**，页眉页脚域必须逐节
  //      `section.Headers/Footers.Item(k).Range.Fields.Update()` 单独更新。
  //   5. 交叉引用**只能按书签**：`Application.CrossReference` 与 `Document.CrossReference`
  //      实测均为 `undefined`，没有 Word 那套"引用类型 + 引用内容"选择器。
  //      REF(wdFieldRef=3) / PAGEREF(wdFieldPageRef=37) 用 `Fields.Add(range, type, "书名签名 \\h", false)`
  //      插入后**立即**就有结果，无需等待。书签不存在 → 结果恒为 "错误！未定义书签。"；
  //      书签为空（起止位置相同）→ REF 结果为空字符串（两者都实测过）。
  //   6. **域后继续插入内容必须跳过域结束标记**：`Fields.Add` 之后若按
  //      `field.Result.End` 直接插入，插入内容会落在域内部，下一次 `Fields.Update()`
  //      会把这段内容连同后面的域一起清掉（实测 advance=0 时 "（第 1 页）" 与 PAGEREF 域消失）。
  //      正确推进量是 `field.Result.End + 1`（实测 advance=1 时模板完整存活、两个域都在）。
  // ===========================================================================

  const WORD_FIELD_TYPE_NAMES = {
    3: "REF", 13: "TOC", 26: "NUMPAGES", 33: "PAGE", 37: "PAGEREF", 54: "AutoNum", 88: "HYPERLINK"
  };

  /** 域类型名：优先取域码首个单词（更准），否则退回类型号映射。 */
  function wordFieldKindName(type, code) {
    const m = /^\s*([A-Za-z]+)/.exec(String(code || ""));
    if (m) return m[1].toUpperCase();
    return WORD_FIELD_TYPE_NAMES[Number(type)] || String(type);
  }

  function wordNormalizeFieldText(text) {
    return String(text === undefined || text === null ? "" : text)
      .replace(/[\r\n\x07\xa0]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  /** 域快照：更新前后各取一次，用于自己统计"哪些域真的变了"。 */
  function wordFieldSnapshot(doc, limit) {
    const max = Number(limit) || 200;
    let total = 0;
    try { total = doc.Fields.Count; } catch (e) { return { total: 0, list: [], truncated: false, error: e.message }; }
    const list = [];
    for (let i = 1; i <= Math.min(total, max); i++) {
      try {
        const f = doc.Fields.Item(i);
        const code = (f.Code ? f.Code.Text : "").replace(/[\r\n\x07]/g, " ").trim();
        list.push({
          index: i,
          type: Number(f.Type),
          kind: wordFieldKindName(f.Type, code),
          code: code,
          result: wordNormalizeFieldText(f.Result ? f.Result.Text : "")
        });
      } catch (e) {
        list.push({ index: i, kind: "?", code: "", result: "", error: e.message });
      }
    }
    return { total: total, list: list, truncated: total > max };
  }

  /**
   * 两个域快照的差异。
   * 不能按 index 直接比：目录重建后域下标会整体平移（实测更新后域数 65 → 63）。
   * 按 `kind|code` 分桶、同桶内按出现顺序配对比较，才是稳定口径。
   */
  function wordDiffFieldSnapshots(before, after, maxChanges) {
    const limit = Number(maxChanges) || 60;
    const bucket = new Map();
    for (const f of before.list) {
      if (f.error) continue;
      const k = f.kind + "|" + f.code;
      if (!bucket.has(k)) bucket.set(k, []);
      bucket.get(k).push(f.result);
    }
    const cursor = new Map();
    const changes = [];
    let compared = 0;
    let changed = 0;
    for (const f of after.list) {
      if (f.error) continue;
      const k = f.kind + "|" + f.code;
      const arr = bucket.get(k);
      if (!arr) continue;
      const n = cursor.get(k) || 0;
      cursor.set(k, n + 1);
      if (n >= arr.length) continue;
      compared++;
      if (arr[n] !== f.result) {
        changed++;
        if (changes.length < limit) {
          changes.push({ kind: f.kind, code: f.code, resultBefore: arr[n], resultAfter: f.result });
        }
      }
    }
    return { compared: compared, changed: changed, changes: changes, changesTruncated: changed > changes.length };
  }

  /** 目录快照：条目文本 + 各条目的页码（目录页码本体就是条目里的 PAGEREF 域结果）。 */
  function wordTocSnapshot(doc) {
    let count = 0;
    try { count = doc.TablesOfContents.Count; } catch (e) { return { count: 0, list: [], error: e.message }; }
    const list = [];
    for (let i = 1; i <= count; i++) {
      const entry = { index: i, entryCount: 0, entryTitles: [], pageNumbers: [], range: null };
      try {
        const toc = doc.TablesOfContents.Item(i);
        const fields = toc.Range.Fields;
        for (let k = 1; k <= fields.Count; k++) {
          const f = fields.Item(k);
          const code = (f.Code ? f.Code.Text : "").replace(/[\r\n\x07]/g, " ").trim();
          const result = wordNormalizeFieldText(f.Result ? f.Result.Text : "");
          if (/^PAGEREF\b/i.test(code)) entry.pageNumbers.push(result);
          else if (/^HYPERLINK\b/i.test(code)) {
            entry.entryCount++;
            entry.entryTitles.push(result.replace(/\t[\d\s]*$/, "").trim());
          }
        }
        entry.range = { start: toc.Range.Start, end: toc.Range.End };
        entry.upperHeadingLevel = (() => { try { return Number(toc.UpperHeadingLevel); } catch (e) { return undefined; } })();
        entry.lowerHeadingLevel = (() => { try { return Number(toc.LowerHeadingLevel); } catch (e) { return undefined; } })();
      } catch (e) { entry.error = e.message; }
      list.push(entry);
    }
    return { count: count, list: list };
  }

  /** 目录更新前后对照（页码逐条比对 + 条目增删）。 */
  function wordCompareTocs(before, after) {
    const out = [];
    const count = Math.max(before.count, after.count);
    for (let i = 1; i <= count; i++) {
      const b = before.list[i - 1] || {};
      const a = after.list[i - 1] || {};
      const pagesB = b.pageNumbers || [];
      const pagesA = a.pageNumbers || [];
      const titlesB = b.entryTitles || [];
      const titlesA = a.entryTitles || [];
      const changedPages = [];
      const maxLen = Math.max(pagesB.length, pagesA.length);
      for (let k = 0; k < maxLen; k++) {
        if (pagesB[k] !== pagesA[k]) {
          changedPages.push({
            entry: titlesA[k] || titlesB[k] || `#${k + 1}`,
            pageBefore: pagesB[k] === undefined ? null : pagesB[k],
            pageAfter: pagesA[k] === undefined ? null : pagesA[k]
          });
        }
      }
      out.push({
        index: i,
        entryCountBefore: b.entryCount === undefined ? null : b.entryCount,
        entryCountAfter: a.entryCount === undefined ? null : a.entryCount,
        pageNumbersBefore: pagesB.slice(0, 60),
        pageNumbersAfter: pagesA.slice(0, 60),
        pageNumbersChanged: changedPages.length > 0,
        changedPageCount: changedPages.length,
        changedPages: changedPages.slice(0, 30),
        changedPagesTruncated: changedPages.length > 30,
        addedEntries: titlesA.filter(t => titlesB.indexOf(t) < 0).slice(0, 20),
        removedEntries: titlesB.filter(t => titlesA.indexOf(t) < 0).slice(0, 20),
        error: a.error || b.error
      });
    }
    return out;
  }

  /** 节的页眉/页脚 story 槽位：只在文档真的启用"首页不同/奇偶页不同"时才去碰 2/3 号槽位。 */
  function wordSectionStorySlots(doc, sectionIndex) {
    const slots = [1];
    try {
      const ps = doc.Sections.Item(sectionIndex).PageSetup;
      if (Number(ps.DifferentFirstPageHeaderFooter) !== 0) slots.push(2);
      if (Number(ps.OddAndEvenPagesHeaderFooter) !== 0) slots.push(3);
    } catch (e) {}
    return slots;
  }

  function wordStorySnapshotAndUpdate(doc, sectionIndex, kind, slot, doUpdate, warnings) {
    const entry = { story: kind + ".Item(" + slot + ")", ok: false };
    try {
      const section = doc.Sections.Item(sectionIndex);
      const story = kind === "Headers" ? section.Headers.Item(slot) : section.Footers.Item(slot);
      const fields = story.Range.Fields;
      entry.fieldCountBefore = fields.Count;
      entry.fieldCodesBefore = [];
      for (let k = 1; k <= fields.Count; k++) {
        entry.fieldCodesBefore.push((fields.Item(k).Code ? fields.Item(k).Code.Text : "").replace(/[\r\n\x07]/g, " ").trim());
      }
      if (doUpdate) {
        const t0 = Date.now();
        fields.Update();
        entry.elapsedMs = Date.now() - t0;
      }
      entry.fieldCountAfter = story.Range.Fields.Count;
      entry.text = (story.Range.Text || "").replace(/[\r\n\x07]/g, "").slice(0, 120);
      entry.ok = true;
    } catch (e) {
      entry.error = e.message;
      if (doUpdate) warnings.push(`第 ${sectionIndex} 节 ${kind}.Item(${slot}) 域更新失败：${e.message}`);
    }
    return entry;
  }

  /** 按出现次数查一段文本，返回命中的 Range（与生产其它函数一致的 Find 用法）。 */
  function wordFindTextRangeIn(doc, text, occurrence) {
    const want = Number(occurrence) || 1;
    const scan = (rng) => {
      let n = 0;
      try {
        const f = rng.Find;
        f.ClearFormatting();
        f.Text = text;
        f.MatchCase = false;
        f.MatchWholeWord = false;
        f.MatchWildcards = false;
        f.Forward = true;
        f.Wrap = 0; // wdFindStop
        while (f.Execute()) {
          n++;
          if (n >= want) return f.Parent;
        }
      } catch (e) {}
      return null;
    };
    const hit = scan(doc.Content);
    if (hit) return hit;
    if (doc.Tables) {
      for (let t = 1; t <= doc.Tables.Count; t++) {
        try {
          const h = scan(doc.Tables.Item(t).Range);
          if (h) return h;
        } catch (e) {}
      }
    }
    return null;
  }

  function wordExistingBookmarkNames(doc, limit) {
    const names = [];
    try {
      for (let i = 1; i <= doc.Bookmarks.Count && names.length < (limit || 30); i++) names.push(doc.Bookmarks.Item(i).Name);
    } catch (e) {}
    return names;
  }

  /**
   * 交叉引用插入：REF（引用书签文字）/ PAGEREF（引用书签所在页码）。
   * 支持先用 anchorText 现场建书签，再插域——书签是交叉引用的唯一前提（宿主无 CrossReference）。
   */
  function wordInsertCrossReference(app, doc, params) {
    const {
      targetBookmark, bookmarkName, anchorText, anchorOccurrence,
      crossReferenceType, crossReferenceTemplate, insertLocation, paragraphIndex, prefixText
    } = params || {};
    const warnings = [];
    const name = String(bookmarkName || targetBookmark || "").trim();
    let bookmark = null;

    if (anchorText !== undefined && anchorText !== null && String(anchorText) !== "") {
      if (!name) throw new Error("用 anchorText 现建书签时必须同时给 bookmarkName（书签名，不能含空格）");
      if (/\s/.test(name)) throw new Error(`书签名 [${name}] 含空格，宿主 Word 书签名不允许空格`);
      const range = wordFindTextRangeIn(doc, String(anchorText), anchorOccurrence);
      if (!range) throw new Error(`未在文档中找到锚点文本 [${anchorText}]（第 ${Number(anchorOccurrence) || 1} 处），未建书签也未插入交叉引用`);
      const rangeInfo = { start: range.Start, end: range.End, text: (range.Text || "").replace(/[\r\n\x07]/g, "") };
      try { doc.Bookmarks.Add(name, range); } catch (e) { throw new Error(`Bookmarks.Add("${name}") 失败：${e.message}`); }
      let exists = false;
      try { exists = doc.Bookmarks.Exists(name); } catch (e) {}
      if (!exists) throw new Error(`Bookmarks.Add("${name}") 执行后书签仍不存在，宿主未创建；已中止，未插入交叉引用`);
      bookmark = { name: name, created: true, range: rangeInfo, text: rangeInfo.text, empty: rangeInfo.start === rangeInfo.end, exists: true };
    } else {
      if (!name) throw new Error("必须提供 bookmarkName 或 targetBookmark —— 本宿主没有 CrossReference 对象，交叉引用**只能**按书签");
      let exists = false;
      try { exists = doc.Bookmarks.Exists(name); } catch (e) {}
      if (!exists) {
        const names = wordExistingBookmarkNames(doc);
        throw new Error(
          `书签 [${name}] 在此文档中不存在，未插入任何交叉引用（REF/PAGEREF 引用不到书签时结果恒为"错误！未定义书签。"）。` +
          (names.length ? `现有书签：${names.join("、")}` : "当前文档没有任何书签；可传 anchorText + bookmarkName 让本工具先建书签。")
        );
      }
      const bm = doc.Bookmarks.Item(name);
      const r = bm.Range;
      bookmark = {
        name: name, created: false, exists: true,
        range: { start: r.Start, end: r.End },
        text: (r.Text || "").replace(/[\r\n\x07]/g, ""),
        empty: r.Start === r.End
      };
    }
    if (bookmark.empty) {
      warnings.push(`书签 [${bookmark.name}] 是空书签（起止位置相同）：真机实测 REF 域结果为空字符串、PAGEREF 仍能给出页码。`);
    }

    // 插入位置
    const loc = insertLocation || "end";
    let pos = null;
    if (loc === "start") pos = 0;
    else if (loc === "selection") {
      const wordApp = getWordApp() || app;
      try {
        const sel = wordApp && wordApp.Selection && wordApp.Selection.Range;
        if (sel) pos = sel.Start === sel.End ? sel.Start : sel.End;
      } catch (e) {}
      if (pos === null) { warnings.push("当前没有可用选区，交叉引用回退插入到文档末尾。"); }
    } else if (loc === "after_paragraph") {
      const idx = Number(paragraphIndex);
      if (!idx || idx < 1 || idx > doc.Paragraphs.Count) {
        throw new Error(`段落索引越界：第 ${paragraphIndex} 段不存在（当前共 ${doc.Paragraphs.Count} 段）`);
      }
      pos = doc.Paragraphs.Item(idx).Range.End - 1;
    }
    if (pos === null) pos = doc.Content.End > 1 ? doc.Content.End - 1 : 0;

    // 模板：{ref} / {page}
    const kind = crossReferenceType || "both";
    if (["ref", "pageref", "both"].indexOf(kind) < 0) {
      throw new Error(`未知的 crossReferenceType: ${kind}（支持 ref, pageref, both）`);
    }
    const template = kind === "both"
      ? String(crossReferenceTemplate || "{ref}（第 {page} 页）")
      : (kind === "ref" ? "{ref}" : "{page}");
    const segments = [];
    const re = /\{(ref|page)\}/g;
    let last = 0;
    let m;
    while ((m = re.exec(template)) !== null) {
      if (m.index > last) segments.push({ text: template.slice(last, m.index) });
      segments.push({ field: m[1] });
      last = m.index + m[0].length;
    }
    if (last < template.length) segments.push({ text: template.slice(last) });

    const inserted = [];
    const errors = [];
    if (prefixText) {
      try { doc.Range(pos, pos).InsertAfter(String(prefixText)); pos += String(prefixText).length; } catch (e) { errors.push(`前缀写入失败：${e.message}`); }
    }
    for (const seg of segments) {
      try {
        if (seg.text) {
          doc.Range(pos, pos).InsertAfter(seg.text);
          pos += seg.text.length;
          continue;
        }
        const type = seg.field === "ref" ? 3 : 37; // wdFieldRef / wdFieldPageRef
        const f = doc.Fields.Add(doc.Range(pos, pos), type, bookmark.name + " \\h", false);
        if (f === null || f === undefined) {
          errors.push(`${seg.field === "ref" ? "REF" : "PAGEREF"} 域插入返回 null（宿主未创建），该段被跳过`);
          continue;
        }
        // 关键：+1 跳过域结束标记，否则后续内容落在域内部，下次 Update 会被清掉（真机实测）
        pos = f.Result.End + 1;
        inserted.push({ field: seg.field === "ref" ? "REF" : "PAGEREF", type: type, code: (f.Code.Text || "").replace(/[\r\n\x07]/g, " ").trim(), result: wordNormalizeFieldText(f.Result.Text) });
      } catch (e) {
        errors.push(`${seg.field ? (seg.field === "ref" ? "REF" : "PAGEREF") + " 域" : "文本段"}插入失败：${e.message}`);
      }
    }

    let updateOk = false;
    try { doc.Fields.Update(); updateOk = true; } catch (e) { warnings.push(`插入后 Fields.Update() 失败：${e.message}`); }

    // 读回：重新按域码定位本次插入的域，拿更新后的真实结果
    for (const item of inserted) {
      try {
        for (let k = 1; k <= doc.Fields.Count; k++) {
          const f = doc.Fields.Item(k);
          const code = (f.Code ? f.Code.Text : "").replace(/[\r\n\x07]/g, " ").trim();
          if (code === item.code) {
            item.resultAfterUpdate = wordNormalizeFieldText(f.Result.Text);
            item.resolved = item.resultAfterUpdate !== "" && !/错误|Error!|未定义书签/.test(item.resultAfterUpdate);
            item.fieldIndex = k;
            break;
          }
        }
      } catch (e) {}
      if (item.resolved === undefined) {
        item.resolved = false;
        warnings.push(`${item.field} 域在更新后未能重新定位，无法确认结果`);
      }
    }

    return {
      success: true,
      documentName: doc.Name,
      action: "insert_cross_reference",
      bookmark: bookmark,
      insertLocation: loc,
      insertedFields: inserted,
      fieldsUpdateAfterInsert: updateOk,
      errors: errors.length ? errors : undefined,
      warnings: warnings,
      hostLimitations: [
        "本宿主没有 CrossReference 对象（Application/Document.CrossReference 实测 undefined），交叉引用只能按**书签**：不能按'标题/图表编号'这类 Word 内置引用类型插入。",
        "REF/PAGEREF 引用不存在的书签时结果恒为 '错误！未定义书签。'；空书签的 REF 结果为空字符串（均已实测）。",
        "域插入后立即就有结果，但页码要在文档完成分页后才准确；若结果可疑请再调 action='update' 刷新一次。"
      ],
      message: `已在 [${doc.Name}] 插入 ${inserted.length} 个交叉引用域（书签 [${bookmark.name}]${bookmark.created ? "，本次新建" : ""}）` +
        (errors.length ? `；有 ${errors.length} 条失败，见 errors` : "") +
        (warnings.length ? `；有 ${warnings.length} 条告警，见 warnings` : "")
    };
  }

  function wordUpdateFields(app, params) {
    const { documentName, action, scope, sectionIndex, tocMode, includeHeadersFooters } = params || {};
    const doc = getWordDocument(app, documentName);
    const act = action || "update";

    if (act === "insert_cross_reference") return wordInsertCrossReference(app, doc, params);
    if (act !== "update") {
      throw new Error(`未知的 Word 域操作: ${action}（支持 update、insert_cross_reference）`);
    }

    const scopeVal = scope || "all";
    if (["all", "toc", "section"].indexOf(scopeVal) < 0) {
      throw new Error(`未知的 scope: ${scopeVal}（支持 all、toc、section；scope='section' 时必须配 sectionIndex）`);
    }
    const mode = tocMode || "full";
    if (["full", "page_numbers"].indexOf(mode) < 0) {
      throw new Error(`未知的 tocMode: ${tocMode}（支持 full、page_numbers）`);
    }
    const warnings = [];
    const errors = [];
    const startedAt = Date.now();

    const sectionCount = (() => { try { return doc.Sections.Count; } catch (e) { return 0; } })();
    let sectionTargets = [];
    if (scopeVal === "section") {
      const idx = Number(sectionIndex);
      if (!idx || idx < 1 || idx > sectionCount) {
        throw new Error(`sectionIndex 越界：收到 ${sectionIndex}，当前文档共 ${sectionCount} 个节`);
      }
      sectionTargets = [idx];
    } else if (scopeVal === "all") {
      for (let i = 1; i <= sectionCount; i++) sectionTargets.push(i);
    }

    const fieldsBefore = wordFieldSnapshot(doc, 200);
    const tocBefore = wordTocSnapshot(doc);

    // 1) 正文域
    const bodyUpdate = { applied: false, scope: scopeVal };
    if (scopeVal === "all") {
      try {
        const t0 = Date.now();
        const ret = doc.Fields.Update();
        bodyUpdate.applied = true;
        bodyUpdate.elapsedMs = Date.now() - t0;
        bodyUpdate.hostReturnValue = typeof ret === "number" ? ret : undefined;
        bodyUpdate.note = "宿主 Fields.Update() 的返回值实测恒为 0，不能当计数用；本工具按快照差异统计。";
        bodyUpdate.coversTableOfContents = "宿主实测：doc.Fields.Update() 会连目录条目一起重建（新增标题会被收录），tocMode='full' 时无需再单独调 toc.Update()。";
      } catch (e) {
        errors.push(`正文域更新失败：${e.message}`);
      }
    } else if (scopeVal === "section") {
      for (const idx of sectionTargets) {
        const entry = { sectionIndex: idx, ok: false };
        try {
          const t0 = Date.now();
          const ret = doc.Sections.Item(idx).Range.Fields.Update();
          entry.fieldsCount = doc.Sections.Item(idx).Range.Fields.Count;
          entry.hostReturnValue = typeof ret === "number" ? ret : undefined;
          entry.elapsedMs = Date.now() - t0;
          entry.ok = true;
        } catch (e) { entry.error = e.message; errors.push(`第 ${idx} 节正文域更新失败：${e.message}`); }
        bodyUpdate.sections = (bodyUpdate.sections || []).concat([entry]);
      }
    }

    // 2) 目录
    const tocUpdate = [];
    if (scopeVal === "all" || scopeVal === "toc") {
      let count = 0;
      try { count = doc.TablesOfContents.Count; } catch (e) { errors.push(`读取目录数量失败：${e.message}`); }
      if (count === 0) {
        warnings.push("文档里没有任何目录（TablesOfContents.Count = 0），本次没有可更新的目录对象。");
      }
      for (let i = 1; i <= count; i++) {
        const item = { index: i, mode: mode, ok: false };
        if (scopeVal === "all" && mode === "full") {
          item.skippedBecause = "doc.Fields.Update() 已重建全部目录（含条目与页码），无需重复调用 toc.Update()";
          item.ok = true;
          tocUpdate.push(item);
          continue;
        }
        try {
          const t0 = Date.now();
          const toc = doc.TablesOfContents.Item(i);
          if (mode === "page_numbers") toc.UpdatePageNumbers();
          else toc.Update();
          item.elapsedMs = Date.now() - t0;
          item.ok = true;
        } catch (e) {
          item.error = e.message;
          errors.push(`第 ${i} 个目录更新失败：${e.message}`);
        }
        tocUpdate.push(item);
      }
    }

    // 3) 页眉/页脚域（逐节逐 story；doc.Fields 不含页眉页脚 story）
    const storyUpdates = [];
    if (includeHeadersFooters !== false && (scopeVal === "all" || scopeVal === "section")) {
      for (const idx of sectionTargets) {
        for (const kind of ["Headers", "Footers"]) {
          for (const slot of wordSectionStorySlots(doc, idx)) {
            storyUpdates.push(wordStorySnapshotAndUpdate(doc, idx, kind, slot, true, warnings));
          }
        }
      }
    }

    const fieldsAfter = wordFieldSnapshot(doc, 200);
    const tocAfter = wordTocSnapshot(doc);
    const diff = wordDiffFieldSnapshots(fieldsBefore, fieldsAfter, 60);
    const tocComparison = wordCompareTocs(tocBefore, tocAfter);
    const pageNumbersChanged = tocComparison.some(t => t.pageNumbersChanged);

    return {
      success: true,
      documentName: doc.Name,
      action: "update",
      scope: scopeVal,
      sectionIndex: scopeVal === "section" ? Number(sectionIndex) : undefined,
      tocMode: mode,
      elapsedMs: Date.now() - startedAt,
      fields: {
        countBefore: fieldsBefore.total,
        countAfter: fieldsAfter.total,
        updatedFields: diff.compared,
        changedFields: diff.changed,
        changes: diff.changes,
        changesTruncated: diff.changesTruncated
      },
      bodyUpdate: bodyUpdate,
      tocUpdate: tocUpdate,
      stories: storyUpdates,
      tablesOfContents: {
        countBefore: tocBefore.count,
        countAfter: tocAfter.count,
        pageNumbersChanged: pageNumbersChanged,
        comparison: tocComparison
      },
      warnings: warnings,
      errors: errors.length ? errors : undefined,
      hostLimitations: [
        "doc.Fields.Update() 的返回值在本宿主实测恒为 0，不能当'更新了几个域'的计数；本工具用更新前后逐域结果快照自己统计 updatedFields / changedFields。",
        "doc.Fields **不含页眉/页脚 story 里的域**，所以页眉页脚域由本工具逐节 Headers/Footers.Item(k).Range.Fields.Update() 单独更新（includeHeadersFooters=false 可关闭）。",
        "域更新无弹窗、无交互（实测 63 个域约 100ms），但目录页码依赖文档已完成分页：刚大批量改完内容时页码可能滞后，建议稍后重跑一次本工具。",
        "交叉引用只能按书签（宿主无 CrossReference 对象）；用 action='insert_cross_reference' 可先建书签再插 REF/PAGEREF。"
      ],
      message:
        `已更新 [${doc.Name}] 的域（scope=${scopeVal}，tocMode=${mode}）：正文域 ${diff.compared} 个中 ${diff.changed} 个结果发生变化；` +
        `${tocAfter.count} 个目录，页码${pageNumbersChanged ? "**已变化**" : "未变化"}` +
        (errors.length ? `；有 ${errors.length} 条错误，见 errors` : "") +
        (warnings.length ? `；有 ${warnings.length} 条告警，见 warnings` : "")
    };
  }

  // ===========================================================================
  // CAP-05：Word 内容控件全类型（RPC: word_manage_content_controls）
  //
  // 真机实测（WPS for Mac 12.x，2026-09-22，全部在临时新建文档上做，原文档只读）：
  //   - `doc.ContentControls.Add(Type, Range)` 可用；实测**可创建**：richText(0)、plainText(1)、
  //     picture(2)、comboBox(3)、dropdownList(4)、buildingBlockGallery(5)、date(6)、checkBox(8)。
  //   - **不支持**：group(7)、repeatingSection(9) —— `Add` 返回 null（不抛错），本工具显式报错。
  //   - **没有 `ListItems`**：`cc.ListItems` 实测 `undefined`（Word 桌面版的 ListItems 在 WPS for Mac 上不存在），
  //     下拉/组合框的选项要用 `cc.DropdownListEntries.Add(text, value)`；`Item(i).Delete()` 可删。
  //     新建的下拉/组合框自带一条**占位条目**（Text=占位文案、Value=""），本工具会先删掉它再写调用方的选项。
  //   - 值写入：plainText/richText/comboBox → `cc.Range.Text`；checkBox → `cc.Checked`（读回 Range.Text 是 ☐/☒）；
  //     date → `cc.Range.Text` + `cc.DateDisplayFormat` + `cc.DateDisplayLocale`；dropdownList → 必须
  //     `DropdownListEntries.Item(k).Select()`，**给 Range.Text 赋一个不在选项里的值是静默无效的**（实测无报错、值不变），
  //     本工具因此会核对选项并如实告警。
  //   - 占位符只能用 `cc.SetPlaceholderText(undefined, undefined, text)`；直接写 `cc.PlaceholderText.Text` 实测**无效**。
  //   - 嵌套：目标范围已在另一个内容控件内时 `Add` 返回 null（本宿主不支持嵌套控件），本工具显式报错。
  //   - 包裹语义：plainText/date 等把原范围**包起来**（文本保留）；dropdown/checkbox 是在原范围**之前**
  //     插入控件并留下原文本，本工具会把残留文本删掉（否则文档里会多出一份"标记原文"）。
  //   - `ContentControls.Item("标题")` 按名字取实测返回 null（无按 Title/Tag 查找），本工具改为遍历匹配。
  // ===========================================================================

  const WORD_CC_TYPES = {
    richText: 0, plainText: 1, picture: 2, comboBox: 3, dropdownList: 4,
    buildingBlockGallery: 5, date: 6, group: 7, checkBox: 8, repeatingSection: 9
  };
  const WORD_CC_TYPE_NAMES = {};
  for (const k in WORD_CC_TYPES) WORD_CC_TYPE_NAMES[WORD_CC_TYPES[k]] = k;
  // 实测不支持创建的类型（Add 返回 null）
  const WORD_CC_UNSUPPORTED = {
    group: "本宿主 WPS for Mac 实测 ContentControls.Add(7, range) 返回 null（不抛错），分组控件无法创建。",
    repeatingSection: "本宿主 WPS for Mac 实测 ContentControls.Add(9, range) 返回 null（不抛错），重复节控件无法创建。"
  };
  const WORD_CC_LIMITED = {
    picture: "可创建（Type=2），但本宿主没有图片填充通路，控件内容只是占位符号，不具备可填写语义。",
    buildingBlockGallery: "可创建（Type=5），但本宿主无法写入构建基块，控件内容只是占位文案，不具备可填写语义。"
  };

  function wordCcTypeNumber(type) {
    if (typeof type === "number" && WORD_CC_TYPE_NAMES[type] !== undefined) return type;
    const key = String(type === undefined || type === null ? "" : type).trim();
    if (WORD_CC_TYPES[key] !== undefined) return WORD_CC_TYPES[key];
    const asNum = Number(key);
    if (!isNaN(asNum) && WORD_CC_TYPE_NAMES[asNum] !== undefined) return asNum;
    throw new Error(`未知的内容控件类型: ${type}（可用：${Object.keys(WORD_CC_TYPES).join("、")}；其中 group、repeatingSection 本宿主不支持）`);
  }

  function wordCcEntries(cc) {
    const arr = [];
    try {
      const col = cc.DropdownListEntries;
      if (!col) return arr;
      for (let i = 1; i <= col.Count; i++) {
        const e = col.Item(i);
        arr.push({
          text: (() => { try { return e.Text; } catch (x) { return undefined; } })(),
          value: (() => { try { return e.Value; } catch (x) { return undefined; } })()
        });
      }
    } catch (e) {}
    return arr;
  }

  /** 单个内容控件的读回值（按类型取"当前值"这一项真正的语义）。 */
  function wordCcValue(cc, type) {
    if (type === 8) {
      return {
        checked: (() => { try { return Boolean(cc.Checked); } catch (e) { return undefined; } })(),
        text: (() => { try { return (cc.Range.Text || "").replace(/[\r\n\x07]/g, ""); } catch (e) { return undefined; } })()
      };
    }
    const v = {
      text: (() => { try { return (cc.Range.Text || "").replace(/[\r\n\x07]/g, ""); } catch (e) { return undefined; } })()
    };
    if (type === 6) {
      v.dateDisplayFormat = (() => { try { return cc.DateDisplayFormat; } catch (e) { return undefined; } })();
      v.dateDisplayLocale = (() => { try { return cc.DateDisplayLocale; } catch (e) { return undefined; } })();
    }
    if (type === 3 || type === 4) v.entries = wordCcEntries(cc);
    return v;
  }

  function wordCcDescribe(cc, index) {
    const type = (() => { try { return Number(cc.Type); } catch (e) { return undefined; } })();
    return {
      index: index,
      type: type,
      typeName: WORD_CC_TYPE_NAMES[type] || String(type),
      title: (() => { try { return cc.Title; } catch (e) { return undefined; } })(),
      tag: (() => { try { return cc.Tag; } catch (e) { return undefined; } })(),
      showingPlaceholder: (() => { try { return Boolean(cc.ShowingPlaceholderText); } catch (e) { return undefined; } })(),
      lockContentControl: (() => { try { return Boolean(cc.LockContentControl); } catch (e) { return undefined; } })(),
      lockContents: (() => { try { return Boolean(cc.LockContents); } catch (e) { return undefined; } })(),
      value: wordCcValue(cc, type),
      range: (() => { try { return { start: cc.Range.Start, end: cc.Range.End }; } catch (e) { return null; } })()
    };
  }

  function wordCcList(doc, limit) {
    const max = Number(limit) || 100;
    let count = 0;
    try { count = doc.ContentControls.Count; } catch (e) { return { count: 0, list: [], error: e.message }; }
    const list = [];
    for (let i = 1; i <= Math.min(count, max); i++) {
      try { list.push(wordCcDescribe(doc.ContentControls.Item(i), i)); }
      catch (e) { list.push({ index: i, error: e.message }); }
    }
    return { count: count, list: list, truncated: count > max };
  }

  /** 按 index / tag / title 定位控件；返回 {cc, index}。 */
  function wordCcLocate(doc, params) {
    const { index, tag, title } = params || {};
    const total = doc.ContentControls.Count;
    if (index !== undefined && index !== null && index !== "") {
      const i = Number(index);
      if (!i || i < 1 || i > total) throw new Error(`内容控件索引越界：收到 ${index}，当前共 ${total} 个内容控件`);
      return { cc: doc.ContentControls.Item(i), index: i };
    }
    const matches = [];
    for (let i = 1; i <= total; i++) {
      const c = doc.ContentControls.Item(i);
      let cTag = "", cTitle = "";
      try { cTag = String(c.Tag === undefined || c.Tag === null ? "" : c.Tag); } catch (e) {}
      try { cTitle = String(c.Title === undefined || c.Title === null ? "" : c.Title); } catch (e) {}
      if (tag !== undefined && tag !== null && tag !== "" && cTag === String(tag)) matches.push(i);
      else if (title !== undefined && title !== null && title !== "" && cTitle === String(title)) matches.push(i);
    }
    if (matches.length === 0) {
      throw new Error(`没有匹配的内容控件（tag=${tag === undefined ? "-" : tag}, title=${title === undefined ? "-" : title}）；当前共 ${total} 个。宿主 ContentControls.Item("名字") 实测不可用，本工具按 Title/Tag 遍历匹配。`);
    }
    return { cc: doc.ContentControls.Item(matches[0]), index: matches[0], matchedIndexes: matches };
  }

  function wordCcApplyListItems(cc, listItems, warnings) {
    const col = cc.DropdownListEntries;
    if (!col) { warnings.push("本宿主该控件没有 DropdownListEntries，选项未写入"); return; }
    // 清掉宿主自带的占位条目（Value 为空串），否则它会成为"第 1 个选项"
    if (col.Count >= 1) {
      try {
        const first = col.Item(1);
        const firstValue = first.Value === undefined || first.Value === null ? "" : String(first.Value);
        if (firstValue === "") first.Delete();
      } catch (e) { warnings.push(`占位选项清理失败：${e.message}`); }
    }
    for (const raw of listItems) {
      const text = typeof raw === "string" ? raw : String(raw && raw.text !== undefined ? raw.text : "");
      const value = typeof raw === "string" ? raw : (raw && raw.value !== undefined && raw.value !== null ? String(raw.value) : text);
      if (text === "") continue;
      try { col.Add(text, value); } catch (e) { warnings.push(`选项 [${text}] 写入失败：${e.message}`); }
    }
  }

  function wordManageContentControls(app, params) {
    const {
      documentName, action, type, index, tag, title, value, checked, dateDisplayFormat, dateDisplayLocale,
      listItems, placeholderText, lockContentControl, lockContents, location, markerText, markerOccurrence,
      paragraphIndex, initialText, clearListItems
    } = params || {};
    const doc = getWordDocument(app, documentName);
    const act = action || "list";
    const warnings = [];
    const hostSupport = {
      supportedTypes: ["richText", "plainText", "comboBox", "dropdownList", "date", "checkBox"],
      createOnlyWithoutContent: Object.keys(WORD_CC_LIMITED),
      unsupportedTypes: WORD_CC_UNSUPPORTED,
      listItemsNote: "本宿主没有 cc.ListItems（实测 undefined）；下拉/组合框选项请用 listItems 参数，落到宿主是 DropdownListEntries.Add(text, value)。",
      lookupNote: "宿主 ContentControls.Item(str) 按名字取实测返回 null；本工具的 set_value / delete 用 index 或 tag / title 遍历匹配。"
    };

    if (act === "list") {
      const ccList = wordCcList(doc, params.maxControls);
      return {
        success: true, documentName: doc.Name, action: "list",
        contentControls: ccList,
        hostSupport: hostSupport,
        message: `[${doc.Name}] 共有 ${ccList.count} 个内容控件`
      };
    }

    if (act === "add") {
      const typeNum = wordCcTypeNumber(type);
      const typeName = WORD_CC_TYPE_NAMES[typeNum];
      if (WORD_CC_UNSUPPORTED[typeName]) {
        throw new Error(`内容控件类型 '${typeName}' 本宿主未实现：${WORD_CC_UNSUPPORTED[typeName]}请改用 ${hostSupport.supportedTypes.join("/")}。`);
      }
      if (typeNum === 7 || typeNum === 9) {
        throw new Error(`内容控件类型 '${typeName}' 本宿主未实现（Add 返回 null），请改用 ${hostSupport.supportedTypes.join("/")}。`);
      }
      if (WORD_CC_LIMITED[typeName]) warnings.push(`${typeName}：${WORD_CC_LIMITED[typeName]}`);

      const loc = location || (markerText ? "marker" : "end");
      let range = null;
      if (loc === "marker") {
        if (!markerText) throw new Error("location='marker' 必须提供 markerText（要被替换成内容控件的标记文本，例如 '____'）");
        range = wordFindTextRangeIn(doc, String(markerText), markerOccurrence);
        if (!range) throw new Error(`未在文档中找到标记文本 [${markerText}]（第 ${Number(markerOccurrence) || 1} 处），未创建任何内容控件`);
      } else if (loc === "start") {
        range = doc.Range(0, 0);
      } else if (loc === "selection") {
        const wordApp = getWordApp() || app;
        try {
          const sel = wordApp && wordApp.Selection && wordApp.Selection.Range;
          if (sel) range = sel;
        } catch (e) {}
        if (!range) { warnings.push("当前没有可用选区，内容控件回退创建到文档末尾。"); }
      } else if (loc === "after_paragraph") {
        const idx = Number(paragraphIndex);
        if (!idx || idx < 1 || idx > doc.Paragraphs.Count) {
          throw new Error(`段落索引越界：第 ${paragraphIndex} 段不存在（当前共 ${doc.Paragraphs.Count} 段）`);
        }
        const end = doc.Paragraphs.Item(idx).Range.End - 1;
        range = doc.Range(end, end);
      }
      if (loc === "end" || !range) {
        const e = doc.Content.End > 1 ? doc.Content.End - 1 : 0;
        range = doc.Range(e, e);
        if (initialText) {
          const text = String(initialText);
          range.InsertAfter(text);
          range = doc.Range(e, e + text.length);
        }
      }

      const target = { start: range.Start, end: range.End, markerText: markerText === undefined ? null : String(markerText) };
      let cc = null;
      try { cc = doc.ContentControls.Add(typeNum, range); }
      catch (e) { throw new Error(`ContentControls.Add(${typeNum}) 抛错：${e.message}`); }
      if (cc === null || cc === undefined) {
        throw new Error(
          `ContentControls.Add(${typeNum}, range) 返回 null，未创建任何控件。真机实测两类原因：` +
          `(1) 目标范围已落在另一个内容控件内部——本宿主不支持嵌套内容控件；` +
          `(2) 该类型本宿主不支持（group=7、repeatingSection=9）。` +
          `当前文档已有 ${doc.ContentControls.Count} 个内容控件，可先用 action='list' 查看范围。`
        );
      }

      // 非包裹型控件（下拉/组合框/复选框）会在原范围之前插入内容并留下原文，删掉残留
      let leftoverDeleted = null;
      try {
        const ccEnd = cc.Range.End;
        if (ccEnd < target.end) {
          const rest = doc.Range(ccEnd, target.end);
          const txt = (rest.Text || "").replace(/[\r\n\x07]/g, "");
          rest.Delete();
          leftoverDeleted = txt;
        }
      } catch (e) { warnings.push(`残留标记文本清理失败：${e.message}`); }

      if (title !== undefined) { try { cc.Title = String(title); } catch (e) { warnings.push(`Title 写入失败：${e.message}`); } }
      if (tag !== undefined) { try { cc.Tag = String(tag); } catch (e) { warnings.push(`Tag 写入失败：${e.message}`); } }
      if (placeholderText !== undefined) {
        // 实测：直接写 cc.PlaceholderText.Text 无效，必须走 SetPlaceholderText(BuildingBlock, Range, Text)
        try { cc.SetPlaceholderText(undefined, undefined, String(placeholderText)); }
        catch (e) { warnings.push(`占位符写入失败：${e.message}`); }
      }
      if (lockContentControl !== undefined) { try { cc.LockContentControl = Boolean(lockContentControl); } catch (e) { warnings.push(`LockContentControl 写入失败：${e.message}`); } }
      if (lockContents !== undefined) { try { cc.LockContents = Boolean(lockContents); } catch (e) { warnings.push(`LockContents 写入失败：${e.message}`); } }
      if (dateDisplayFormat !== undefined && typeNum === 6) { try { cc.DateDisplayFormat = String(dateDisplayFormat); } catch (e) { warnings.push(`日期格式写入失败：${e.message}`); } }
      if (dateDisplayLocale !== undefined && typeNum === 6) { try { cc.DateDisplayLocale = Number(dateDisplayLocale); } catch (e) { warnings.push(`日期区域写入失败：${e.message}`); } }

      if ((typeNum === 3 || typeNum === 4) && Array.isArray(listItems) && listItems.length) {
        wordCcApplyListItems(cc, listItems, warnings);
      }

      let valueApplied = null;
      if (checked !== undefined && typeNum === 8) {
        try { cc.Checked = Boolean(checked); valueApplied = Boolean(cc.Checked); }
        catch (e) { warnings.push(`Checked 写入失败：${e.message}`); }
      }
      if (value !== undefined && value !== null && String(value) !== "") {
        const v = String(value);
        if (typeNum === 4) {
          let hit = false;
          try {
            const col = cc.DropdownListEntries;
            for (let i = 1; i <= col.Count; i++) {
              const e = col.Item(i);
              const t = e.Text === undefined ? "" : String(e.Text);
              const val = e.Value === undefined ? "" : String(e.Value);
              if (t === v || val === v) { e.Select(); hit = true; break; }
            }
          } catch (e) { warnings.push(`下拉选中失败：${e.message}`); }
          valueApplied = hit;
          if (!hit) warnings.push(`下拉值 [${v}] 不在选项列表里：宿主对 Range.Text 赋非法值是**静默无效**的（实测无报错、值不变），本次未写入该值。`);
        } else if (typeNum === 8) {
          warnings.push("checkBox 的值请用 checked 参数，value 已忽略。");
        } else {
          try { cc.Range.Text = v; valueApplied = (cc.Range.Text || "").replace(/[\r\n\x07]/g, ""); }
          catch (e) { warnings.push(`值写入失败：${e.message}`); }
        }
      }
      if (typeNum === 6 && value !== undefined && value !== null && String(value) !== "") {
        try { cc.Range.Text = String(value); } catch (e) {}
      }

      // 读回：按控件自身对象读（创建后下标会随文档顺序变化，不能按 Count 取）
      let created = null;
      try { created = wordCcDescribe(cc, null); } catch (e) { warnings.push(`创建后读回失败：${e.message}`); }
      const listAfter = wordCcList(doc, 100);
      // 定位本次创建控件的真实下标（按 range.start 匹配）
      let createdIndex = null;
      if (created && created.range) {
        for (const item of listAfter.list) {
          if (item.range && item.range.start === created.range.start) { createdIndex = item.index; break; }
        }
      }
      if (createdIndex !== null) created.index = createdIndex;

      const tagApplied = tag === undefined ? null : (created && created.tag === String(tag));
      const titleApplied = title === undefined ? null : (created && created.title === String(title));
      if (tagApplied === false) warnings.push(`Tag 写入后读回不一致（期望 [${tag}]，读回 [${created && created.tag}]）：宿主可能丢弃了该属性，请以本返回的读回值为准。`);
      if (titleApplied === false) warnings.push(`Title 写入后读回不一致（期望 [${title}]，读回 [${created && created.title}]）。`);

      return {
        success: true, documentName: doc.Name, action: "add",
        type: typeName, typeCode: typeNum, location: loc,
        targetRange: target,
        leftoverMarkerTextDeleted: leftoverDeleted,
        valueApplied: valueApplied,
        tagApplied: tagApplied,
        titleApplied: titleApplied,
        created: created,
        contentControls: { count: listAfter.count, list: listAfter.list },
        hostSupport: hostSupport,
        warnings: warnings,
        message: `已在 [${doc.Name}] 创建 ${typeName} 内容控件（${loc}${markerText ? "：" + markerText : ""}），文档现有 ${listAfter.count} 个内容控件`
      };
    }

    if (act === "set_value") {
      const found = wordCcLocate(doc, params);
      const cc = found.cc;
      const typeNum = (() => { try { return Number(cc.Type); } catch (e) { return undefined; } })();
      const typeName = WORD_CC_TYPE_NAMES[typeNum] || String(typeNum);
      const applied = {};

      if (title !== undefined) { try { cc.Title = String(title); applied.title = cc.Title; } catch (e) { warnings.push(`Title 写入失败：${e.message}`); } }
      if (tag !== undefined) { try { cc.Tag = String(tag); applied.tag = cc.Tag; } catch (e) { warnings.push(`Tag 写入失败：${e.message}`); } }
      if (placeholderText !== undefined) {
        try { cc.SetPlaceholderText(undefined, undefined, String(placeholderText)); applied.placeholderApplied = true; }
        catch (e) { warnings.push(`占位符写入失败：${e.message}`); }
      }
      if (lockContentControl !== undefined) { try { cc.LockContentControl = Boolean(lockContentControl); applied.lockContentControl = cc.LockContentControl; } catch (e) { warnings.push(`LockContentControl 写入失败：${e.message}`); } }
      if (lockContents !== undefined) { try { cc.LockContents = Boolean(lockContents); applied.lockContents = cc.LockContents; } catch (e) { warnings.push(`LockContents 写入失败：${e.message}`); } }
      if (dateDisplayFormat !== undefined && typeNum === 6) { try { cc.DateDisplayFormat = String(dateDisplayFormat); applied.dateDisplayFormat = cc.DateDisplayFormat; } catch (e) { warnings.push(`日期格式写入失败：${e.message}`); } }
      if (clearListItems && (typeNum === 3 || typeNum === 4)) {
        try {
          const col = cc.DropdownListEntries;
          while (col.Count > 0) col.Item(col.Count).Delete();
          applied.listItemsCleared = true;
        } catch (e) { warnings.push(`选项清空失败：${e.message}`); }
      }
      if (Array.isArray(listItems) && listItems.length && (typeNum === 3 || typeNum === 4)) {
        wordCcApplyListItems(cc, listItems, warnings);
        applied.listItemsApplied = listItems.length;
      }
      if (checked !== undefined) {
        if (typeNum === 8) { try { cc.Checked = Boolean(checked); applied.checked = Boolean(cc.Checked); } catch (e) { warnings.push(`Checked 写入失败：${e.message}`); } }
        else warnings.push(`控件类型是 ${typeName}，checked 参数已忽略（只有 checkBox 支持）。`);
      }
      if (value !== undefined && value !== null && String(value) !== "") {
        const v = String(value);
        if (typeNum === 8) {
          warnings.push("checkBox 的值请用 checked 参数，value 已忽略。");
        } else if (typeNum === 4) {
          let hit = false;
          try {
            const col = cc.DropdownListEntries;
            for (let i = 1; i <= col.Count; i++) {
              const e = col.Item(i);
              const t = e.Text === undefined ? "" : String(e.Text);
              const val = e.Value === undefined ? "" : String(e.Value);
              if (t === v || val === v) { e.Select(); hit = true; break; }
            }
          } catch (e) { warnings.push(`下拉选中失败：${e.message}`); }
          applied.value = hit ? v : null;
          if (!hit) warnings.push(`下拉值 [${v}] 不在选项列表里：宿主静默无效（实测无报错、值不变），本次未写入。`);
        } else {
          try { cc.Range.Text = v; applied.value = (cc.Range.Text || "").replace(/[\r\n\x07]/g, ""); }
          catch (e) { warnings.push(`值写入失败：${e.message}`); }
        }
      }

      const readBack = wordCcDescribe(cc, found.index);
      return {
        success: true, documentName: doc.Name, action: "set_value",
        type: typeName, matchedIndex: found.index,
        matchedIndexes: found.matchedIndexes,
        applied: applied,
        contentControl: readBack,
        hostSupport: hostSupport,
        warnings: warnings,
        message: `已更新 [${doc.Name}] 第 ${found.index} 个内容控件（${typeName}）`
      };
    }

    if (act === "delete") {
      if (params.all === true) {
        const before = (() => { try { return doc.ContentControls.Count; } catch (e) { return 0; } })();
        const deleted = [];
        for (let i = before; i >= 1; i--) {
          try { doc.ContentControls.Item(i).Delete(); deleted.push(i); }
          catch (e) { warnings.push(`删除第 ${i} 个控件失败：${e.message}`); }
        }
        return {
          success: true, documentName: doc.Name, action: "delete", deleteAll: true,
          deletedCount: deleted.length, countBefore: before,
          contentControls: wordCcList(doc, 100),
          warnings: warnings,
          message: `已删除 [${doc.Name}] 的 ${deleted.length} 个内容控件（原有 ${before} 个）`
        };
      }
      const found = wordCcLocate(doc, params);
      const before = wordCcDescribe(found.cc, found.index);
      found.cc.Delete();
      const after = wordCcList(doc, 100);
      return {
        success: true, documentName: doc.Name, action: "delete",
        deleted: before,
        contentControls: { count: after.count, list: after.list },
        hostSupport: hostSupport,
        warnings: warnings,
        message: `已删除 [${doc.Name}] 第 ${found.index} 个内容控件（${before.typeName}，删除后剩余 ${after.count} 个）`
      };
    }

    throw new Error(`未知的内容控件操作: ${action}（支持 list, add, set_value, delete）`);
  }
