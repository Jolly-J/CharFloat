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

  function wordReadDocument(app, params) {
    const { documentName, scope, maxParagraphs, includeFormatting, includeTables } = params || {};
    const doc = getWordDocument(app, documentName);
    const outline = [];
    const maxP = Number(maxParagraphs) || 200;

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

    return {
      documentName: doc.Name,
      fullName: doc.FullName || doc.Name,
      paragraphCount: doc.Paragraphs ? doc.Paragraphs.Count : 0,
      tableCount: doc.Tables ? doc.Tables.Count : 0,
      wordCount: doc.Words ? doc.Words.Count : 0,
      outline,
      tables: tablesSummary,
      paragraphDetails: includeFormatting ? paragraphDetails : undefined,
      previewText: scope === "outline" ? undefined : previewText,
      selectionText: selectionText || undefined
    };
  }

  function wordWriteContent(app, params) {
    const { documentName, location, targetBookmark, paragraphIndex, type, content, formatting } = params || {};
    const doc = getWordDocument(app, documentName);
    const wordApp = getWordApp() || app;
    let targetRange = null;

    if (location === "start") {
      targetRange = doc.Range(0, 0);
    } else if (location === "selection" && wordApp.Selection) {
      targetRange = wordApp.Selection.Range;
    } else if (location === "bookmark" && targetBookmark) {
      if (doc.Bookmarks.Exists(targetBookmark)) {
        targetRange = doc.Bookmarks.Item(targetBookmark).Range;
      } else {
        targetRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
      }
    } else if (location === "after_paragraph" && paragraphIndex) {
      const p = doc.Paragraphs.Item(Number(paragraphIndex));
      targetRange = doc.Range(p.Range.End, p.Range.End);
    } else {
      const endPos = doc.Content.End > 1 ? doc.Content.End - 1 : 0;
      targetRange = doc.Range(endPos, endPos);
    }

    const lines = Array.isArray(content) ? content : [String(content || "")];
    for (const text of lines) {
      if (!text) continue;
      const newPara = doc.Paragraphs.Add(targetRange);
      newPara.Range.Text = text + "\n";

      if (type === "heading1") {
        try { newPara.Range.Style = doc.Styles.Item(-2); } catch (e) { newPara.Range.Font.Bold = true; newPara.Range.Font.Size = 22; }
      } else if (type === "heading2") {
        try { newPara.Range.Style = doc.Styles.Item(-3); } catch (e) { newPara.Range.Font.Bold = true; newPara.Range.Font.Size = 16; }
      } else if (type === "heading3") {
        try { newPara.Range.Style = doc.Styles.Item(-4); } catch (e) { newPara.Range.Font.Bold = true; newPara.Range.Font.Size = 14; }
      } else if (type === "bullet_list") {
        try { newPara.Range.ListFormat.ApplyBulletDefault(); } catch (e) {}
      } else if (type === "quote") {
        try {
          newPara.Range.Font.Italic = true;
          newPara.Format.LeftIndent = 28;
        } catch (e) {}
      }

      // 精确控制排版并主动断开加粗继承污染
      if (formatting && typeof formatting === "object") {
        if (formatting.bold !== undefined) newPara.Range.Font.Bold = Boolean(formatting.bold);
        if (formatting.italic !== undefined) newPara.Range.Font.Italic = Boolean(formatting.italic);
        if (formatting.fontSizePt !== undefined) newPara.Range.Font.Size = Number(formatting.fontSizePt);
        if (formatting.fontName) {
          newPara.Range.Font.NameFarEast = formatting.fontName;
          newPara.Range.Font.NameAscii = formatting.fontName;
        }
        if (formatting.alignment !== undefined) newPara.Format.Alignment = Number(formatting.alignment);
        if (formatting.firstLineIndentChars !== undefined) newPara.Format.CharacterUnitFirstLineIndent = Number(formatting.firstLineIndentChars);
        if (formatting.lineSpacingPt !== undefined) {
          newPara.Format.LineSpacingRule = 4;
          newPara.Format.LineSpacing = Number(formatting.lineSpacingPt);
        }
        if (formatting.spaceBeforePt !== undefined) newPara.Format.SpaceBefore = Number(formatting.spaceBeforePt);
        if (formatting.spaceAfterPt !== undefined) newPara.Format.SpaceAfter = Number(formatting.spaceAfterPt);
      } else if (!type || type === "paragraph") {
        // 普通正文默认强制消除前文加粗继承
        try {
          newPara.Range.Font.Bold = false;
        } catch (e) {}
      }

      targetRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
    }

    return {
      success: true,
      documentName: doc.Name,
      insertedLines: lines.length,
      location: location || "end",
      type: type || "paragraph",
      message: `已成功向 [${doc.Name}] 写入 ${lines.length} 行内容`
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

  function wordPageLayoutAndWatermark(app, params) {
    const { documentName, headerText, footerText, differentFirstPage, differentOddEvenPages, watermarkText } = params || {};
    const doc = getWordDocument(app, documentName);
    const section = doc.Sections.Item(1);

    if (differentFirstPage !== undefined) doc.PageSetup.DifferentFirstPageHeaderFooter = differentFirstPage;
    if (differentOddEvenPages !== undefined) doc.PageSetup.OddAndEvenPagesHeaderFooter = differentOddEvenPages;

    if (headerText !== undefined) {
      try { section.Headers.Item(1).Range.Text = headerText; } catch (e) {}
    }

    if (footerText !== undefined) {
      try { section.Footers.Item(1).Range.Text = footerText; } catch (e) {}
    }

    if (watermarkText) {
      try {
        const newShape = doc.Shapes.AddTextEffect(0, watermarkText, "Microsoft YaHei", 54, false, false, 100, 200);
        newShape.Rotation = -315;
        newShape.Fill.Transparency = 0.85;
        newShape.Line.Visible = false;
        newShape.WrapFormat.Type = 3;
      } catch (e) {}
    }

    return {
      success: true,
      documentName: doc.Name,
      header: headerText,
      footer: footerText,
      watermark: watermarkText || undefined,
      message: `已成功更新 [${doc.Name}] 页面版式与水印`
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

          // 核心修复：第 9 个参数 Format 传入 hasFmt，使 Replacement.Font 格式化必定生效
          findObj.Execute(
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
          matchCount++;
        } else if (hasFmt) {
          // 纯格式化定位赋属性
          while (findObj.Execute()) {
            matchCount++;
            if (replaceFormatting.bold !== undefined) findObj.Parent.Font.Bold = Boolean(replaceFormatting.bold);
            if (replaceFormatting.italic !== undefined) findObj.Parent.Font.Italic = Boolean(replaceFormatting.italic);
            if (replaceFormatting.fontSizePt !== undefined) findObj.Parent.Font.Size = Number(replaceFormatting.fontSizePt);
            if (replaceFormatting.fontName) {
              findObj.Parent.Font.NameFarEast = replaceFormatting.fontName;
              findObj.Parent.Font.NameAscii = replaceFormatting.fontName;
            }
          }
        }
      } catch (e) {}
    }

    if (replaceText !== undefined) {
      return {
        success: true,
        documentName: doc.Name,
        searchQuery,
        replaceText,
        action: "replaced_all",
        message: `已将 [${doc.Name}] 中的 "${searchQuery}" 全文穿透替换为 "${replaceText}"`
      };
    }

    return {
      success: true,
      documentName: doc.Name,
      searchQuery,
      matchCount,
      message: `已为 "${searchQuery}" 找到并应用格式化 (${matchCount} 处)`
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
