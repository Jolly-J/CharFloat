  // ---------------------------------------------------------------------------
  // ppt.js — PPT 宿主写入与 RPC 实现
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
  // ==========================================
  // PowerPoint (演示) 模块 7 大核心操作实现
  // ==========================================

  function hexToPptColor(hex) {
    if (!hex) return 0;
    const clean = hex.replace("#", "");
    if (clean.length !== 6) return 0;
    const r = parseInt(clean.substring(0, 2), 16);
    const g = parseInt(clean.substring(2, 4), 16);
    const b = parseInt(clean.substring(4, 6), 16);
    return (b << 16) | (g << 8) | r;
  }

  // Adapt only newly generated shapes; existing slide content stays untouched.
  // 几何/字号/容量计算全部委托给 ppt-layout.js 的纯函数，本函数只读写宿主对象。
  function fitGeneratedPptShapes(slide, firstIndex, page, warnings) {
    const ratio = pptScaleForPage(page);
    const sx = ratio.sx, sy = ratio.sy, scale = ratio.scale;
    for (let i = firstIndex; i <= slide.Shapes.Count; i++) {
      const shape = slide.Shapes.Item(i);
      shape.Left *= sx; shape.Top *= sy;
      shape.Width *= sx; shape.Height *= sy;
      if (!shape.HasTextFrame || !shape.TextFrame.HasText) continue;
      const frame = shape.TextFrame, range = frame.TextRange;
      frame.AutoSize = 0;
      frame.WordWrap = true;
      const preferred = Number(range.Font.Size) * scale;
      const minimum = Math.min(preferred, 12 * scale);
      const innerW = Math.max(1, shape.Width - Number(frame.MarginLeft || 0) - Number(frame.MarginRight || 0));
      const innerH = Math.max(1, shape.Height - Number(frame.MarginTop || 0) - Number(frame.MarginBottom || 0));
      const text = String(range.Text || "");
      const size = pptFitFontSize(text, preferred, minimum, scale, innerW, innerH);
      range.Font.Size = size;
      if (pptEstimateLines(text, size, innerW) * size * 1.25 > innerH) warnings.push({ shapeId: shape.Id, reason: "文字可能溢出，请缩短内容或扩大文本框并检查预览" });
    }
  }

  function addSlideHeader(slide, title, themeColor) {
    const titleBox = slide.Shapes.AddTextbox(1, 50, 30, 620, 50);
    titleBox.TextFrame.TextRange.Text = title || "主题要点";
    titleBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
    titleBox.TextFrame.TextRange.Font.Size = 24;
    titleBox.TextFrame.TextRange.Font.Bold = true;
    titleBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor(themeColor || "#0F4C81");

    try {
      const line = slide.Shapes.AddLine(50, 75, 120, 75);
      line.Line.ForeColor.RGB = hexToPptColor(themeColor || "#0F4C81");
      line.Line.Weight = 2.5;
    } catch (e) {}
  }

  function renderPptCards(slide, cards, colCount, themeColor, customTop, customH) {
    const cols = Math.max(1, Math.min(Number(colCount) || 3, 4));
    if (cards.length > cols) throw new Error("卡片数量超过分栏数，请拆分为多页，避免内容被截断");
    const count = cards.length;
    const totalW = 620;
    const startX = 50;
    const startY = customTop === undefined ? 100 : Number(customTop);
    const cardH = customH === undefined ? 260 : Number(customH);
    if (!Number.isFinite(startY) || !Number.isFinite(cardH) || startY < 0 || cardH <= 0 || startY + cardH > 405) throw new Error("卡片位置或高度超出页面");
    const gap = 16;
    const cardW = (totalW - (cols - 1) * gap) / cols;

    for (let i = 0; i < count; i++) {
      const item = cards[i];
      const x = startX + i * (cardW + gap);
      const accent = item.accentColor || themeColor || "#0F4C81";

      // 1. 底卡 (圆角矩形 5)
      try {
        const bgShape = slide.Shapes.AddShape(5, x, startY, cardW, cardH);
        bgShape.Fill.Solid();
        bgShape.Fill.ForeColor.RGB = hexToPptColor("#F8FAFC");
        bgShape.Line.ForeColor.RGB = hexToPptColor("#E2E8F0");
        bgShape.Line.Weight = 1;
      } catch (e) {}

      // 2. 顶条强调线
      try {
        const topBar = slide.Shapes.AddShape(1, x + 10, startY + 12, 36, 4);
        topBar.Fill.Solid();
        topBar.Fill.ForeColor.RGB = hexToPptColor(accent);
        topBar.Line.Visible = false;
      } catch (e) {}

      // 3. 卡片标签 (Tag)
      if (item.tag) {
        const tagBox = slide.Shapes.AddTextbox(1, x + 10, startY + 20, cardW - 20, 22);
        tagBox.TextFrame.TextRange.Text = item.tag.toUpperCase();
        tagBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
        tagBox.TextFrame.TextRange.Font.Size = 10;
        tagBox.TextFrame.TextRange.Font.Bold = true;
        tagBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor(accent);
      }

      // 4. 卡片标题
      const titleY = item.tag ? startY + 42 : startY + 22;
      const cardTitleBox = slide.Shapes.AddTextbox(1, x + 10, titleY, cardW - 20, 36);
      cardTitleBox.TextFrame.TextRange.Text = item.title || `要点 ${i + 1}`;
      cardTitleBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
      cardTitleBox.TextFrame.TextRange.Font.Size = cols >= 4 ? 14 : 16;
      cardTitleBox.TextFrame.TextRange.Font.Bold = true;
      cardTitleBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#0F172A");

      // 5. 卡片正文描述
      const descY = titleY + 36;
      const descBox = slide.Shapes.AddTextbox(1, x + 10, descY, cardW - 20, cardH - (descY - startY) - 10);
      descBox.TextFrame.TextRange.Text = item.description || "";
      descBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
      descBox.TextFrame.TextRange.Font.Size = cols >= 4 ? 11 : 12;
      descBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#475569");
      descBox.TextFrame.WordWrap = true;
    }
  }

  function renderPptChart(slide, chartSpec, left, top, width, height) {
    const { chartType, categories, series, title, hasLegend, showDataLabels } = chartSpec || {};
    let typeCode = 51; // xlColumnClustered 柱状图
    if (chartType === "line") typeCode = 4; // xlLine 折线图
    else if (chartType === "pie") typeCode = 5; // xlPie 饼图
    else if (chartType === "bar") typeCode = 57; // xlBarClustered 条形图
    else if (chartType === "column_stacked" || chartType === "stacked_column") typeCode = 52;
    else if (chartType === "bar_stacked" || chartType === "stacked_bar") typeCode = 58;
    else if (chartType === "bar_of_pie" || chartType === "pie_bar") typeCode = 68;
    else if (typeof chartType === "number") typeCode = chartType;

    let chartShape = null;
    let addError = null;
    try {
      chartShape = slide.Shapes.AddChart(typeCode, left, top, width, height);
    } catch (e) {
      addError = e.message;
    }
    if (!chartShape) {
      try {
        chartShape = slide.Shapes.AddChart2(-1, typeCode, left, top, width, height);
      } catch (e) {
        addError = (addError ? addError + "；" : "") + e.message;
      }
    }
    // 本机 WPS 实测：`AddChart` / `AddChart2` 都是 function，但**返回 null 且不创建任何形状**（问题台账 ISS-80）。
    // 原代码不检查返回值，随后 `chartShape.Chart` 抛错并被包装成"图表已创建，但数据配置未完成"——
    // 让调用方以为图已经建出来了，实际什么都没建。
    if (!chartShape) {
      throw new Error(
        "本宿主未能创建 PPT 原生图表：AddChart/AddChart2 未返回图表对象" +
        "（已实测本机 WPS 上二者返回 null 且不创建形状）" +
        (addError ? `；宿主返回：${addError}` : "") +
        "。替代方案：用矢量形状自行绘制，或改用 WPS 表格的原生图表。"
      );
    }

    try {
      const chart = chartShape.Chart;
      if (title) {
        try {
          chart.HasTitle = true;
          chart.ChartTitle.Text = title;
        } catch (e) { throw e; }
      }
      if (hasLegend !== undefined) {
        try { chart.HasLegend = Boolean(hasLegend); } catch (e) { throw e; }
      }

      if (Array.isArray(series) && series.length > 0) {
        const sc = chart.SeriesCollection();
        while (sc.Count > series.length) {
          try { sc.Item(sc.Count).Delete(); } catch (e) { throw e; }
        }
        for (let i = 0; i < series.length; i++) {
          const sData = series[i];
          let sObj = null;
          if (i + 1 <= sc.Count) {
            sObj = sc.Item(i + 1);
          } else {
            try { sObj = sc.NewSeries(); } catch (e) { throw e; }
          }
          if (!sObj) throw new Error("无法创建目标数据系列");

          if (sData.name) {
            try { sObj.Name = sData.name; } catch (e) { throw e; }
          }
          if (categories && categories.length > 0 && i === 0) {
            try { sObj.XValues = categories; } catch (e) { throw e; }
          }
          if (Array.isArray(sData.values)) {
            try { sObj.Values = sData.values; } catch (e) { throw e; }
            // 写后读回：本机 WPS 上 `Values = [...]` **不抛错也不生效**（问题台账 ISS-75），
            // 图表会显示宿主默认数据而调用方毫无察觉。这里核对是否真的落上。
            let readBackRaw = null;
            try { readBackRaw = sObj.Values; } catch (e) { readBackRaw = null; }
            if (readBackRaw !== null && readBackRaw !== undefined) {
              let readBack = [];
              try {
                const n = typeof readBackRaw.Count === "number" ? readBackRaw.Count : readBackRaw.length;
                for (let k = 0; k < n; k++) readBack.push(readBackRaw[k]);
              } catch (e) { readBack = []; }
              const want = sData.values.map(Number);
              const got = readBack.map(Number);
              const same = got.length === want.length &&
                want.every((value, index) => !Number.isFinite(value) || !Number.isFinite(got[index]) || value === got[index]);
              if (!same) {
                throw new Error(
                  `第 ${i + 1} 个数据系列的 Values 未生效：写入 ${JSON.stringify(want).slice(0, 70)}，读回 ${JSON.stringify(got).slice(0, 70)}。` +
                  `宿主忽略了数组赋值（实测本机 WPS 的 ChartData.Workbook 为 null）；请改用矢量形状绘制图表。`
                );
              }
            }
          }
          if (sData.chartType) {
            const sType = sData.chartType;
            try {
              if (sType === "line" || sType === 4) sObj.ChartType = 4;
              else if (sType === "line_markers" || sType === 65) sObj.ChartType = 65;
              else if (typeof sType === "number") sObj.ChartType = sType;
            } catch (e) { throw e; }
          }
          if (sData.axisGroup === 2 || sData.secondaryAxis) {
            try { sObj.AxisGroup = 2; } catch (e) { throw e; }
          }
          if (sData.color) {
            try {
              sObj.Format.Fill.Solid();
              sObj.Format.Fill.ForeColor.RGB = hexToPptColor(sData.color);
            } catch (e) { throw e; }
          }
          if (sData.hasDataLabels || showDataLabels) {
            try { sObj.HasDataLabels = true; } catch (e) { throw e; }
          }
        }
      } else if (categories && categories.length > 0) {
        try {
          const sc = chart.SeriesCollection();
          if (sc.Count > 0) {
            sc.Item(1).XValues = categories;
          }
        } catch (e) { throw e; }
      }
    } catch (e) {
      // 配置失败时把**半成品形状**删掉再抛错：否则会留下一张数据空白/错乱的图表（问题台账 ISS-81）。
      let cleaned = false;
      try {
        chartShape.Delete();
        cleaned = true;
      } catch (delErr) {}
      throw new Error(
        `PPT 原生图表已创建但配置失败${cleaned ? "（已删除半成品形状）" : "（半成品形状删除失败，请手动清理）"}：${e.message}`
      );
    }

    return chartShape;
  }

  function pptReadPresentation(app, params) {
    const { presentationName, includeNotes, maxSlides } = params || {};
    const pres = getPptPresentation(app, presentationName);
    const pptApp = getPptApp() || app;
    const slides = [];
    const count = pres.Slides ? pres.Slides.Count : 0;
    const maxS = Number(maxSlides) || 50;

    for (let i = 1; i <= Math.min(count, maxS); i++) {
      try {
        const slide = pres.Slides.Item(i);
        let title = `Slide ${i}`;
        const snippets = [];
        let notes = "";

        const shapeCount = slide.Shapes ? slide.Shapes.Count : 0;
        for (let s = 1; s <= shapeCount; s++) {
          const shp = slide.Shapes.Item(s);
          if (shp.HasTextFrame && shp.TextFrame.HasText) {
            const text = (shp.TextFrame.TextRange.Text || "").trim().replace(/[\r\n\x07]/g, " ");
            if (shp.Type === 14 || s === 1) {
              if (text && title === `Slide ${i}`) title = text;
            }
            if (text) snippets.push(text);
          }
        }

        if (includeNotes !== false) {
          try {
            if (slide.NotesPage && slide.NotesPage.Shapes) {
              const notesShape = slide.NotesPage.Shapes.Placeholders.Item(2);
              if (notesShape && notesShape.HasTextFrame && notesShape.TextFrame.HasText) {
                notes = (notesShape.TextFrame.TextRange.Text || "").trim();
              }
            }
          } catch (e) {}
        }

        slides.push({
          index: i,
          title,
          shapeCount,
          notes: notes || undefined,
          textSnippets: snippets.slice(0, 5)
        });
      } catch (e) {}
    }

    let activeIdx = 1;
    try {
      if (pptApp.ActiveWindow && pptApp.ActiveWindow.Selection && pptApp.ActiveWindow.Selection.SlideRange) {
        activeIdx = pptApp.ActiveWindow.Selection.SlideRange.SlideIndex;
      }
    } catch (e) {}

    return {
      presentationName: pres.Name,
      fullName: pres.FullName || pres.Name,
      ...pptPageSize(pres),
      slideCount: count,
      activeSlideIndex: activeIdx,
      slides
    };
  }

  function pptGenerateDeck(app, params) {
    const { presentationName, themeColor, themePreset, slides } = params || {};
    if (!Array.isArray(slides) || slides.length === 0) {
      throw new Error("缺少必要参数: slides 数组");
    }
    const pres = getPptPresentation(app, presentationName);
    const baseColor = themeColor || (themePreset === "tech_purple" ? "#4B38B3" : "#0F4C81");
    const createdIndices = [];
    const page = pptPageSize(pres);
    const warnings = [];
    // Design coordinates are mapped to the actual page before returning.
    const pageWidth = 720;
    const pageHeight = 405;

    for (let i = 0; i < slides.length; i++) {
      const spec = slides[i];
      const slide = pres.Slides.Add(pres.Slides.Count + 1, 12);
      const slideIdx = slide.SlideIndex;
      createdIndices.push(slideIdx);

      if (spec.layout === "title") {
        try {
          const bg = slide.Shapes.AddShape(1, 0, 0, pageWidth, pageHeight);
          bg.Fill.Solid();
          bg.Fill.ForeColor.RGB = hexToPptColor(baseColor);
          bg.Line.Visible = false;
        } catch (e) {}

        const titleBox = slide.Shapes.AddTextbox(1, 60, 140, pageWidth - 120, 80);
        titleBox.TextFrame.TextRange.Text = spec.title || "演示文稿";
        titleBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
        titleBox.TextFrame.TextRange.Font.Size = 36;
        titleBox.TextFrame.TextRange.Font.Bold = true;
        titleBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#FFFFFF");

        if (spec.subtitle) {
          const subBox = slide.Shapes.AddTextbox(1, 60, 230, pageWidth - 120, 40);
          subBox.TextFrame.TextRange.Text = spec.subtitle;
          subBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
          subBox.TextFrame.TextRange.Font.Size = 18;
          subBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#E2E8F0");
        }
      } else if (spec.layout === "cards_2" || spec.layout === "cards_3" || spec.layout === "cards_4" || spec.cards) {
        addSlideHeader(slide, spec.title, baseColor);
        const colCount = spec.layout === "cards_2" ? 2 : (spec.layout === "cards_4" ? 4 : (spec.cards ? Math.min(spec.cards.length, 4) : 3));
        renderPptCards(slide, spec.cards || [], colCount, baseColor);
      } else if (spec.layout === "chart" && spec.chart) {
        addSlideHeader(slide, spec.title, baseColor);
        renderPptChart(slide, spec.chart, 60, 90, pageWidth - 120, pageHeight - 120);
      } else if (spec.layout === "end") {
        try {
          const bg = slide.Shapes.AddShape(1, 0, 0, pageWidth, pageHeight);
          bg.Fill.Solid();
          bg.Fill.ForeColor.RGB = hexToPptColor(baseColor);
          bg.Line.Visible = false;
        } catch (e) {}

        const titleBox = slide.Shapes.AddTextbox(1, 60, 160, pageWidth - 120, 80);
        titleBox.TextFrame.TextRange.Text = spec.title || "THANK YOU";
        titleBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
        titleBox.TextFrame.TextRange.Font.Size = 40;
        titleBox.TextFrame.TextRange.Font.Bold = true;
        titleBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#FFFFFF");
        titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2;
      } else {
        addSlideHeader(slide, spec.title, baseColor);
        if (Array.isArray(spec.bulletPoints) && spec.bulletPoints.length > 0) {
          const contentBox = slide.Shapes.AddTextbox(1, 60, 100, pageWidth - 120, pageHeight - 140);
          contentBox.TextFrame.TextRange.Text = spec.bulletPoints.join("\n");
          contentBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
          contentBox.TextFrame.TextRange.Font.Size = 16;
          contentBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#334155");
        } else {
          // content 布局缺正文时，原实现静默只出标题、不报任何问题（问题台账 ISS-54）。
          // 这里显式告警，避免调用方以为这一页已经做完了。
          warnings.push({
            slideIndex: slideIdx,
            reason: "content 布局未提供 bulletPoints：本页只有标题、没有正文。请补 bulletPoints，或改用 cards/chart 布局"
          });
        }
      }

      fitGeneratedPptShapes(slide, 1, page, warnings);

      if (spec.notes) {
        try {
          if (slide.NotesPage && slide.NotesPage.Shapes) {
            const notesShape = slide.NotesPage.Shapes.Placeholders.Item(2);
            if (notesShape && notesShape.HasTextFrame) {
              notesShape.TextFrame.TextRange.Text = spec.notes;
            }
          }
        } catch (e) {}
      }
    }

    return {
      success: true,
      ...page,
      layoutWarnings: warnings,
      visualVerificationRequired: true,
      presentationName: pres.Name,
      createdSlidesCount: slides.length,
      createdSlideIndices: createdIndices,
      themeColor: baseColor,
      message: `已成功基于大纲批量生成 ${slides.length} 页专业商业演示胶片！`
    };
  }

  // 幻灯片页码校验：宿主对越界页码会抛内部 JS 错误（`Cannot read properties of null (reading 'Delete')`），
  // 调用方看不出是哪一页越界（问题台账 ISS-79）。这里统一前置校验并给中文上下文。
  function pptRequireSlideIndex(pres, slideIndex, actionLabel) {
    const total = pres.Slides.Count;
    const idx = Number(slideIndex);
    if (!slideIndex || !Number.isFinite(idx) || Math.floor(idx) !== idx) {
      throw new Error(`${actionLabel} 需要提供整数 slideIndex（1-based）；当前文稿共 ${total} 页`);
    }
    if (total === 0) {
      throw new Error(`${actionLabel} 失败：当前文稿一页都没有，请先用 action='add' 新增幻灯片`);
    }
    if (idx < 1 || idx > total) {
      throw new Error(`${actionLabel} 的 slideIndex=${idx} 越界：当前文稿 [${pres.Name}] 共 ${total} 页，有效范围 1~${total}`);
    }
    return idx;
  }

  function pptSavePresentation(app, params) {
    const { presentationName, filePath, format } = params || {};
    const pres = getPptPresentation(app, presentationName);
    const fmt = String(format || "").toLowerCase();
    const target = filePath ? String(filePath) : null;

    if (target && (fmt === "pdf" || target.toLowerCase().endsWith(".pdf"))) {
      // ppSaveAsPDF = 32；ExportAsFixedFormat(Path, FixedFormatType=2 表示 PDF)
      try {
        pres.ExportAsFixedFormat(target, 2);
      } catch (e) {
        pres.SaveAs(target, 32);
      }
      return {
        success: true,
        presentationName: pres.Name,
        savedPath: target,
        format: "pdf",
        message: `演示文稿 [${pres.Name}] 已导出为 PDF: ${target}`
      };
    }

    if (target) {
      try {
        pres.SaveAs(target);
      } catch (e) {
        pres.SaveAs(target, 24);
      }
      return {
        success: true,
        presentationName: pres.Name,
        savedPath: target,
        format: "pptx",
        message: `演示文稿 [${pres.Name}] 已另存为: ${target}`
      };
    }

    // 未命名的演示文稿：先判 Path 再决定要不要调 Save。
    // 实测（WPS for Mac 12.0）：对未命名文稿调用 `pres.Save()` **不抛错**，可能弹出"另存为"对话框；
    // 在宿主里弹模态框会卡住后续 RPC，因此这里不盲目调用，直接要求 filePath。
    let savePath = null;
    try { savePath = pres.Path ? String(pres.Path) : null; } catch (e) { savePath = null; }
    if (!savePath) {
      return {
        success: false,
        presentationName: pres.Name,
        savedPath: null,
        hostError: "Presentation.Path 为空：该文稿尚未保存到磁盘",
        attempts: [{ action: "save", ok: false, hostError: "文稿没有文件路径，未调用 Presentation.Save（未命名文稿的 Save 可能弹另存为对话框并阻塞自动化）" }],
        message: `演示文稿 [${pres.Name}] 尚未保存到磁盘，无法原地保存；请提供 filePath 走另存为（例如 /Users/.../方案.pptx）`
      };
    }

    pres.Save();
    const verified = (() => { try { return Number(pres.Saved) !== 0; } catch (e) { return null; } })();
    const fullName = (() => { try { return pres.FullName || null; } catch (e) { return null; } })();
    return {
      success: true,
      presentationName: pres.Name,
      savedPath: fullName || (savePath.replace(/[\\/]+$/, "") + "/" + pres.Name),
      format: "pptx",
      hostError: verified === false ? "保存后 Presentation.Saved 读回 false，落盘结果未确认" : undefined,
      verifiedSavedFlag: verified,
      message: `演示文稿 [${pres.Name}] 已原地保存（保存标记 Saved=${verified}）`
    };
  }

  function pptManageSlides(app, params) {
    const { presentationName, action, slideIndex, targetIndex, layoutIndex, backgroundColor, filePath, format, isVisible } = params || {};

    // 新建演示文稿：不需要预先存在的目标文稿，先处理
    if (action === "new_presentation") {
      const pptApp = getPptApp() || app;
      if (!pptApp) throw new Error("WPS 演示 (PowerPoint) 未就绪");
      const created = pptApp.Presentations.Add();
      try { if (isVisible !== false && created.Application) created.Application.Visible = true; } catch (e) {}
      let savedPath = null;
      if (filePath) {
        try { created.SaveAs(String(filePath)); savedPath = String(filePath); } catch (e) { savedPath = null; }
      }
      return {
        success: true,
        presentationName: created.Name,
        slideCount: created.Slides ? created.Slides.Count : 0,
        savedPath: savedPath,
        appended: false,
        message: savedPath
          ? `已新建演示文稿 [${created.Name}] 并另存为 ${savedPath}`
          : `已新建演示文稿 [${created.Name}]（尚未保存到磁盘；后续写入请显式传 presentationName）`
      };
    }

    const pres = getPptPresentation(app, presentationName);

    switch (action) {
      case "save":
      case "save_as":
        return pptSavePresentation(app, { presentationName: presentationName, filePath: filePath, format: format });
      case "add": {
        const idx = slideIndex ? Number(slideIndex) : (pres.Slides.Count + 1);
        if (slideIndex && (idx < 1 || idx > pres.Slides.Count + 1)) {
          throw new Error(`add 的 slideIndex=${idx} 越界：当前共 ${pres.Slides.Count} 页，可在 1~${pres.Slides.Count + 1} 之间插入`);
        }
        const lIndex = Number(layoutIndex) || 12;
        const newSlide = pres.Slides.Add(idx, lIndex);
        return {
          success: true,
          presentationName: pres.Name,
          slideIndex: newSlide.SlideIndex,
          slideCount: pres.Slides.Count,
          appended: true,
          idempotent: false,
          message: `已在位置 ${newSlide.SlideIndex} 新增幻灯片（追加型操作，重复调用会继续新增）`
        };
      }
      case "delete": {
        const idx = pptRequireSlideIndex(pres, slideIndex, "delete");
        const slide = pres.Slides.Item(idx);
        slide.Delete();
        return {
          success: true,
          presentationName: pres.Name,
          deletedIndex: idx,
          slideCount: pres.Slides.Count,
          message: `已成功删除第 ${idx} 页幻灯片（当前剩 ${pres.Slides.Count} 页）`
        };
      }
      case "move": {
        const idx = pptRequireSlideIndex(pres, slideIndex, "move");
        const to = Number(targetIndex);
        if (!targetIndex || !Number.isFinite(to) || to < 1 || to > pres.Slides.Count) {
          throw new Error(`move 的 targetIndex=${targetIndex} 越界：有效范围 1~${pres.Slides.Count}`);
        }
        const slide = pres.Slides.Item(idx);
        slide.MoveTo(to);
        return { success: true, presentationName: pres.Name, from: idx, to: to, slideCount: pres.Slides.Count, message: `幻灯片已从第 ${idx} 页移动到第 ${to} 页` };
      }
      case "duplicate": {
        const idx = pptRequireSlideIndex(pres, slideIndex, "duplicate");
        const slide = pres.Slides.Item(idx);
        slide.Duplicate();
        return {
          success: true,
          presentationName: pres.Name,
          originalIndex: idx,
          slideCount: pres.Slides.Count,
          appended: true,
          idempotent: false,
          message: `已克隆第 ${idx} 页幻灯片（追加型操作，重复调用会继续克隆）`
        };
      }
      case "set_background": {
        const idx = pptRequireSlideIndex(pres, slideIndex, "set_background");
        if (!backgroundColor) throw new Error("set_background 操作必须提供 backgroundColor");
        const slide = pres.Slides.Item(idx);
        const total = pres.Slides.Count;

        // 读回背景色的辅助：不同宿主返回的 RGB 可能是 number 也可能是其它形态，统一成 "R,G,B"
        const readRgb = (target) => {
          try {
            const rgb = Number(target.Background.Fill.ForeColor.RGB);
            if (!Number.isFinite(rgb)) return null;
            const b = rgb & 0xff, g = (rgb >> 8) & 0xff, r = (rgb >> 16) & 0xff;
            return r + "," + g + "," + b;
          } catch (e) { return null; }
        };
        const neighbourIndex = idx === 1 ? Math.min(2, total) : idx - 1;
        const neighbourBefore = neighbourIndex !== idx ? readRgb(pres.Slides.Item(neighbourIndex)) : null;

        // 关键：该页若仍"跟随母版背景"，`slide.Background` 可能指向母版对象，写入会**串改全部页**
        // （问题台账 ISS-87，受控复现：设第 1 页后第 2 页也变红）。先显式断开与母版的关联再写。
        try { slide.FollowMasterBackground = false; } catch (e) {}
        slide.Background.Fill.Solid();
        slide.Background.Fill.ForeColor.RGB = hexToPptColor(backgroundColor);

        // 写后校验：目标页确实变了，且相邻页**没有被串改**
        const applied = readRgb(slide);
        const neighbourAfter = neighbourIndex !== idx ? readRgb(pres.Slides.Item(neighbourIndex)) : null;
        if (neighbourBefore !== null && neighbourAfter !== null && neighbourBefore !== neighbourAfter) {
          throw new Error(
            `设置背景时串改了相邻页：第 ${neighbourIndex} 页背景由 ${neighbourBefore} 变成了 ${neighbourAfter}。` +
            `已尝试先断开 FollowMasterBackground；请检查该稿的母版/版式是否被直接修改。`
          );
        }
        return {
          success: true,
          presentationName: pres.Name,
          slideIndex: idx,
          backgroundColor,
          appliedRgb: applied,
          neighbourChecked: neighbourIndex !== idx ? { slideIndex: neighbourIndex, before: neighbourBefore, after: neighbourAfter } : null,
          message: `已将第 ${slideIndex} 页背景设为 ${backgroundColor}`
        };
      }
      default:
        throw new Error(`未知的 PPT 页面操作: ${action}`);
    }
  }

  function pptAddBusinessCards(app, params) {
    const { presentationName, slideIndex, columnCount, cards, topY, cardHeight } = params || {};
    if (!Array.isArray(cards) || cards.length === 0) throw new Error("缺少 cards 数组");
    const pres = getPptPresentation(app, presentationName);
    const idx = Number(slideIndex) || (pres.Slides.Count > 0 ? 1 : 1);
    const slide = pres.Slides.Item(idx);
    const cols = columnCount || (cards.length === 2 ? 2 : (cards.length === 4 ? 4 : 3));

    const page = pptPageSize(pres);
    const warnings = [];
    const firstIndex = slide.Shapes.Count + 1;
    renderPptCards(slide, cards, cols, "#0F4C81",
      topY === undefined ? undefined : Number(topY) * 405 / page.pageHeight,
      cardHeight === undefined ? undefined : Number(cardHeight) * 405 / page.pageHeight);
    fitGeneratedPptShapes(slide, firstIndex, page, warnings);
    return {
      success: true,
      presentationName: pres.Name,
      slideIndex: idx,
      ...page,
      layoutWarnings: warnings,
      visualVerificationRequired: true,
      columns: cols,
      cardsCount: cards.length,
      message: `已在第 ${idx} 页成功排版 ${cards.length} 张现代化商业信息卡片`
    };
  }

  function pptInsertNativeChart(app, params) {
    const { presentationName, slideIndex, chartType, title, categories, series, left, top, width, height } = params || {};
    const pres = getPptPresentation(app, presentationName);
    const idx = Number(slideIndex) || (pres.Slides.Count > 0 ? 1 : 1);
    const slide = pres.Slides.Item(idx);

    const page = pptPageSize(pres);
    const l = left !== undefined ? Number(left) : page.pageWidth / 12;
    const t = top !== undefined ? Number(top) : page.pageHeight * 90 / 405;
    const w = width !== undefined ? Number(width) : page.pageWidth * 600 / 720;
    const h = height !== undefined ? Number(height) : page.pageHeight * 280 / 405;

    renderPptChart(slide, { chartType: chartType || "column", title, categories, series }, l, t, w, h);
    return {
      success: true,
      presentationName: pres.Name,
      slideIndex: idx,
      chartType: chartType || "column",
      message: `已在第 ${idx} 页成功插入原生矢量图表`
    };
  }

  function getShapeTypeName(typeCode) {
    switch (typeCode) {
      case 1: return "shape"; // msoAutoShape
      case 3: return "chart"; // msoChart
      case 6: return "group"; // msoGroup
      case 13: return "picture"; // msoPicture
      case 14: return "placeholder"; // msoPlaceholder
      case 17: return "textbox"; // msoTextBox
      case 19: return "table"; // msoTable
      case 24: return "smartArt"; // msoSmartArt
      default: return `type_${typeCode}`;
    }
  }

  function pptGetSlideShapes(app, params) {
    const { presentationName, slideIndex } = params || {};
    const pres = getPptPresentation(app, presentationName);
    const pptApp = getPptApp() || app;
    let idx = Number(slideIndex);
    if (!idx || isNaN(idx)) {
      try {
        if (pptApp.ActiveWindow && pptApp.ActiveWindow.Selection && pptApp.ActiveWindow.Selection.SlideRange) {
          idx = pptApp.ActiveWindow.Selection.SlideRange.SlideIndex;
        }
      } catch (e) {}
      if (!idx) idx = 1;
    }

    const slide = pres.Slides.Item(idx);
    const count = slide.Shapes ? slide.Shapes.Count : 0;
    const shapes = [];

    for (let s = 1; s <= count; s++) {
      try {
        const shp = slide.Shapes.Item(s);
        const hasText = Boolean(shp.HasTextFrame && shp.TextFrame.HasText);
        let textContent = "";
        if (hasText) {
          textContent = (shp.TextFrame.TextRange.Text || "").trim();
        }

        const hasTable = Boolean(shp.HasTable);
        let tableMeta = null;
        if (hasTable && shp.Table) {
          tableMeta = {
            rows: shp.Table.Rows ? shp.Table.Rows.Count : 0,
            columns: shp.Table.Columns ? shp.Table.Columns.Count : 0
          };
        }

        const hasChart = Boolean(shp.HasChart);

        shapes.push({
          shapeIndex: s,
          shapeId: shp.Id,
          name: shp.Name || `Shape_${s}`,
          typeCode: shp.Type,
          typeName: getShapeTypeName(shp.Type),
          left: Math.round(Number(shp.Left) * 10) / 10,
          top: Math.round(Number(shp.Top) * 10) / 10,
          width: Math.round(Number(shp.Width) * 10) / 10,
          height: Math.round(Number(shp.Height) * 10) / 10,
          rotation: shp.Rotation || 0,
          zOrderPosition: shp.ZOrderPosition || s,
          hasText,
          fontSize: hasText ? Number(shp.TextFrame.TextRange.Font.Size) : undefined,
          text: textContent ? (textContent.length > 200 ? textContent.slice(0, 200) + "..." : textContent) : undefined,
          hasTable,
          table: tableMeta || undefined,
          hasChart
        });
      } catch (e) {
        log(`读取第 ${idx} 页第 ${s} 个形状元数据失败: ${e.message}`);
      }
    }

    return {
      success: true,
      presentationName: pres.Name,
      slideIndex: idx,
      ...pptPageSize(pres),
      shapeCount: count,
      shapes,
      // 版式与占位符读回（ISS-44）：宿主目前无法区分矩形/圆角/椭圆（typeCode 都是 1），
      // 但占位符（typeCode 14）与所在版式名可读，用于核对"是否套用了正确版式"。
      layout: (() => {
        try {
          return {
            name: slide.CustomLayout ? slide.CustomLayout.Name : undefined,
            layoutIndex: Number(slide.Layout),
            placeholderCount: slide.Shapes.Placeholders ? slide.Shapes.Placeholders.Count : shapes.filter(s => s.typeCode === 14).length,
            placeholders: (() => {
              try {
                const list = [];
                const phCount = slide.Shapes.Placeholders.Count;
                for (let p = 1; p <= phCount; p++) {
                  const ph = slide.Shapes.Placeholders.Item(p);
                  list.push({
                    shapeId: ph.Id,
                    name: ph.Name,
                    placeholderType: (() => { try { return Number(ph.PlaceholderFormat.Type); } catch (e) { return undefined; } })(),
                    hasText: Boolean(ph.HasTextFrame && ph.TextFrame.HasText),
                    text: (() => { try { return ph.TextFrame.HasText ? String(ph.TextFrame.TextRange.Text).slice(0, 200) : ""; } catch (e) { return undefined; } })()
                  });
                }
                return list;
              } catch (e) { return []; }
            })()
          };
        } catch (e) {
          return { error: e.message };
        }
      })(),
      message: `已成功获取第 ${idx} 页幻灯片中全部 ${shapes.length} 个形状的几何与属性信息`
    };
  }

  function findShapeOnSlide(slide, shapeIdOrIndex) {
    if (shapeIdOrIndex === undefined || shapeIdOrIndex === null) return null;
    const count = slide.Shapes.Count;
    // 1. 如果传入数字且在范围内，先尝试直接按索引或 ID 获取
    if (typeof shapeIdOrIndex === "number" || /^\d+$/.test(String(shapeIdOrIndex))) {
      const num = Number(shapeIdOrIndex);
      // 先遍历按 Id 匹配
      for (let s = 1; s <= count; s++) {
        const item = slide.Shapes.Item(s);
        if (item.Id === num) return item;
      }
      // 再按索引匹配
      if (num >= 1 && num <= count) {
        return slide.Shapes.Item(num);
      }
    }
    // 2. 按名称查找
    for (let s = 1; s <= count; s++) {
      const item = slide.Shapes.Item(s);
      if (item.Name === String(shapeIdOrIndex)) return item;
    }
    // 3. 兜底直接 Item
    try {
      return slide.Shapes.Item(shapeIdOrIndex);
    } catch (e) {
      return null;
    }
  }

  function pptManageShapesAndMedia(app, params) {
    const {
      presentationName,
      slideIndex,
      action,
      shapeId,
      shapeId1,
      shapeId2,
      shapeType,
      text,
      left,
      top,
      width,
      height,
      fontSize,
      fontColor,
      fontBold,
      alignment,
      fillColor,
      lineColor,
      rotation,
      zOrderAction,
      alignType
    } = params || {};

    const pres = getPptPresentation(app, presentationName);
    const pptApp = getPptApp() || app;
    let idx = Number(slideIndex);
    if (!idx || isNaN(idx)) {
      try {
        if (pptApp.ActiveWindow && pptApp.ActiveWindow.Selection && pptApp.ActiveWindow.Selection.SlideRange) {
          idx = pptApp.ActiveWindow.Selection.SlideRange.SlideIndex;
        }
      } catch (e) {}
      if (!idx) idx = 1;
    }
    const slide = pres.Slides.Item(idx);

    const l = left !== undefined ? Number(left) : 100;
    const t = top !== undefined ? Number(top) : 100;
    const w = width !== undefined ? Number(width) : 200;
    const h = height !== undefined ? Number(height) : 100;

    switch (action) {
      case "add_textbox": {
        const tb = slide.Shapes.AddTextbox(1, l, t, w, h);
        const tr = tb.TextFrame.TextRange;
        tr.Text = text || "新文本框";
        tr.Font.Name = "Microsoft YaHei";
        if (fontSize) tr.Font.Size = Number(fontSize);
        if (fontColor) tr.Font.Color.RGB = hexToPptColor(fontColor);
        if (fontBold !== undefined) tr.Font.Bold = Boolean(fontBold);
        if (alignment !== undefined) {
          const alignMap = { left: 1, center: 2, right: 3, justify: 4 };
          tr.ParagraphFormat.Alignment = alignMap[alignment] || 1;
        }
        return { success: true, shapeId: tb.Id, left: tb.Left, top: tb.Top, width: tb.Width, height: tb.Height, fontSize: Number(tr.Font.Size), unit: "pt", message: "已成功添加文本框" };
      }

      case "add_shape": {
        let typeCode = 1;
        if (shapeType === "rounded_rectangle") typeCode = 5;
        else if (shapeType === "oval") typeCode = 9;
        else if (shapeType === "arrow") typeCode = 13;

        const shp = slide.Shapes.AddShape(typeCode, l, t, w, h);
        if (text) {
          shp.TextFrame.TextRange.Text = text;
          shp.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
          if (fontSize) shp.TextFrame.TextRange.Font.Size = Number(fontSize);
        }
        if (fillColor) {
          shp.Fill.Solid();
          shp.Fill.ForeColor.RGB = hexToPptColor(fillColor);
        }
        if (lineColor) {
          shp.Line.ForeColor.RGB = hexToPptColor(lineColor);
        }
        return { success: true, shapeId: shp.Id, left: shp.Left, top: shp.Top, width: shp.Width, height: shp.Height, unit: "pt", message: "已成功添加形状" };
      }

      case "update_shape": {
        const targetId = shapeId;
        if (targetId === undefined || targetId === null) throw new Error("update_shape 操作必须提供 shapeId");
        const shp = findShapeOnSlide(slide, targetId);
        if (!shp) throw new Error(`未在第 ${idx} 页找到形状: ${targetId}`);

        if (left !== undefined) shp.Left = Number(left);
        if (top !== undefined) shp.Top = Number(top);
        if (width !== undefined) shp.Width = Number(width);
        if (height !== undefined) shp.Height = Number(height);
        if (rotation !== undefined) shp.Rotation = Number(rotation);

        if (text !== undefined && shp.HasTextFrame) {
          shp.TextFrame.TextRange.Text = text;
        }
        if (shp.HasTextFrame && shp.TextFrame.HasText) {
          const tr = shp.TextFrame.TextRange;
          if (fontSize !== undefined) tr.Font.Size = Number(fontSize);
          if (fontColor !== undefined) tr.Font.Color.RGB = hexToPptColor(fontColor);
          if (fontBold !== undefined) tr.Font.Bold = Boolean(fontBold);
          if (alignment !== undefined) {
            const alignMap = { left: 1, center: 2, right: 3, justify: 4 };
            tr.ParagraphFormat.Alignment = alignMap[alignment] || 1;
          }
        }
        if (fillColor !== undefined) {
          shp.Fill.Solid();
          shp.Fill.ForeColor.RGB = hexToPptColor(fillColor);
        }
        if (lineColor !== undefined) {
          shp.Line.ForeColor.RGB = hexToPptColor(lineColor);
        }

        return {
          success: true,
          shapeId: shp.Id,
          left: shp.Left,
          top: shp.Top,
          width: shp.Width,
          height: shp.Height,
          message: `已成功更新形状 [${shp.Id}] 的属性`
        };
      }

      case "swap_shapes": {
        const id1 = shapeId1 !== undefined ? shapeId1 : shapeId;
        const id2 = shapeId2;
        if (!id1 || !id2) throw new Error("swap_shapes 必须提供 shapeId1 和 shapeId2 两个目标形状标识");

        const shp1 = findShapeOnSlide(slide, id1);
        const shp2 = findShapeOnSlide(slide, id2);
        if (!shp1) throw new Error(`未找到第一个形状: ${id1}`);
        if (!shp2) throw new Error(`未找到第二个形状: ${id2}`);

        const top1 = Number(shp1.Top);
        const height1 = Number(shp1.Height);
        const left1 = Number(shp1.Left);

        const top2 = Number(shp2.Top);
        const height2 = Number(shp2.Height);
        const left2 = Number(shp2.Left);

        // 智能垂直互换（保持上下文视觉流）
        if (top1 < top2) {
          // shp1 在上方，shp2 在下方
          const gap = top2 - (top1 + height1);
          const effectiveGap = gap > 0 ? gap : 15;
          // 将 shp2 移到上方原 shp1 的 Top
          shp2.Top = top1;
          // 将 shp1 移到 shp2 下方
          shp1.Top = top1 + height2 + effectiveGap;
        } else {
          // shp2 在上方，shp1 在下方
          const gap = top1 - (top2 + height2);
          const effectiveGap = gap > 0 ? gap : 15;
          shp1.Top = top2;
          shp2.Top = top2 + height1 + effectiveGap;
        }

        return {
          success: true,
          slideIndex: idx,
          shape1: { id: shp1.Id, oldTop: top1, newTop: shp1.Top, height: height1 },
          shape2: { id: shp2.Id, oldTop: top2, newTop: shp2.Top, height: height2 },
          message: `已成功将形状 [${shp1.Id}] 与 [${shp2.Id}] 在第 ${idx} 页进行精准垂直互换！`
        };
      }

      case "set_z_order": {
        if (!shapeId) throw new Error("set_z_order 必须提供 shapeId");
        const shp = findShapeOnSlide(slide, shapeId);
        if (!shp) throw new Error(`未找到形状: ${shapeId}`);

        // 0: msoBringToFront, 1: msoSendToBack, 2: msoBringForward, 3: msoSendBackward
        const zMap = {
          bring_to_front: 0,
          send_to_back: 1,
          bring_forward: 2,
          send_backward: 3
        };
        const cmd = zMap[zOrderAction || "bring_to_front"];
        if (cmd !== undefined) {
          shp.ZOrder(cmd);
        }
        return { success: true, shapeId: shp.Id, zOrderAction, message: `已成功调整形状 [${shp.Id}] 的图层层级` };
      }

      case "align_shapes": {
        const ids = params.shapeIds || (shapeId1 && shapeId2 ? [shapeId1, shapeId2] : []);
        if (!Array.isArray(ids) || ids.length === 0) throw new Error("align_shapes 必须提供 shapeIds 数组");
        const alignMode = alignType || "center"; // left | center | right | top | middle | bottom

        let refVal = null;
        for (let i = 0; i < ids.length; i++) {
          const shp = findShapeOnSlide(slide, ids[i]);
          if (!shp) continue;
          if (i === 0) {
            if (alignMode === "left") refVal = shp.Left;
            else if (alignMode === "center") refVal = shp.Left + shp.Width / 2;
            else if (alignMode === "right") refVal = shp.Left + shp.Width;
            else if (alignMode === "top") refVal = shp.Top;
            else if (alignMode === "middle") refVal = shp.Top + shp.Height / 2;
            else if (alignMode === "bottom") refVal = shp.Top + shp.Height;
          } else {
            if (alignMode === "left") shp.Left = refVal;
            else if (alignMode === "center") shp.Left = refVal - shp.Width / 2;
            else if (alignMode === "right") shp.Left = refVal - shp.Width;
            else if (alignMode === "top") shp.Top = refVal;
            else if (alignMode === "middle") shp.Top = refVal - shp.Height / 2;
            else if (alignMode === "bottom") shp.Top = refVal - shp.Height;
          }
        }
        return { success: true, alignMode, alignedCount: ids.length, message: `已完成 ${ids.length} 个形状的 [${alignMode}] 对齐` };
      }

      case "delete_shape": {
        if (!shapeId) throw new Error("delete_shape 操作必须提供 shapeId");
        const shp = findShapeOnSlide(slide, shapeId);
        if (!shp) throw new Error(`未找到形状: ${shapeId}`);
        const delId = shp.Id;
        shp.Delete();
        return { success: true, message: `已成功删除形状: ${delId}` };
      }

      default:
        throw new Error(`未知的形状/多媒体操作: ${action}`);
    }
  }

  function pptManageTable(app, params) {
    const {
      presentationName,
      slideIndex,
      action,
      shapeId,
      tableIndex,
      rows,
      columns,
      columnWidths,
      rowHeights,
      left,
      top,
      width,
      height,
      data,
      row,
      column,
      text,
      fontSize,
      fontColor,
      fontBold,
      fillColor,
      headerFillColor,
      headerFontSize,
      bodyFontSize,
      borderColor,
      zebra
    } = params || {};

    const pres = getPptPresentation(app, presentationName);
    const pptApp = getPptApp() || app;
    let idx = Number(slideIndex);
    if (!idx || isNaN(idx)) {
      try {
        if (pptApp.ActiveWindow && pptApp.ActiveWindow.Selection && pptApp.ActiveWindow.Selection.SlideRange) {
          idx = pptApp.ActiveWindow.Selection.SlideRange.SlideIndex;
        }
      } catch (e) {}
      if (!idx) idx = 1;
    }
    const slide = pres.Slides.Item(idx);

    function findTableShape() {
      if (shapeId) {
        const shp = findShapeOnSlide(slide, shapeId);
        if (shp && shp.HasTable && shp.Table) return shp;
      }
      let tCount = 0;
      const targetTableIdx = Number(tableIndex) || 1;
      for (let s = 1; s <= slide.Shapes.Count; s++) {
        const shp = slide.Shapes.Item(s);
        if (shp.HasTable && shp.Table) {
          tCount++;
          if (tCount === targetTableIdx) return shp;
        }
      }
      return null;
    }

    function applyCellBorders(tbl, bColorHex) {
      const bColor = hexToPptColor(bColorHex || "#333333");
      const whiteColor = hexToPptColor("#FFFFFF");
      const rowCount = tbl.Rows.Count;
      const colCount = tbl.Columns.Count;

      for (let r = 1; r <= rowCount; r++) {
        for (let c = 1; c <= colCount; c++) {
          try {
            const cell = tbl.Cell(r, c);
            for (let b = 1; b <= 4; b++) {
              const border = cell.Borders.Item(b);
              border.Visible = true;
              border.Weight = 1;

              if (r === 1) {
                // 表头行：上下外框与最左/最右外侧使用黑色闭合线，内部纵向使用高保真白色分割线
                if (b === 1 || b === 3) {
                  border.ForeColor.RGB = bColor;
                  border.Weight = 1;
                } else if (b === 2) {
                  if (c === 1) {
                    border.ForeColor.RGB = bColor;
                    border.Weight = 1;
                  } else {
                    border.ForeColor.RGB = whiteColor;
                    border.Weight = 1.5;
                  }
                } else if (b === 4) {
                  if (c === colCount) {
                    border.ForeColor.RGB = bColor;
                    border.Weight = 1;
                  } else {
                    border.ForeColor.RGB = whiteColor;
                    border.Weight = 1.5;
                  }
                }
              } else {
                border.ForeColor.RGB = bColor;
                border.Weight = 1;
              }
            }
          } catch (e) {}
        }
      }
    }

    switch (action) {
      case "create_table": {
        const rCount = Number(rows) || 3;
        const cCount = Number(columns) || 3;
        const l = left !== undefined ? Number(left) : 32.6;
        const t = top !== undefined ? Number(top) : 58;
        const w = width !== undefined ? Number(width) : 895;
        const h = height !== undefined ? Number(height) : 435;

        const tableShape = slide.Shapes.AddTable(rCount, cCount, l, t, w, h);
        const tbl = tableShape.Table;

        try {
          tbl.ApplyStyle("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", false);
        } catch (e) {}
        try {
          tbl.FirstRow = false;
          tbl.BandedRows = false;
          tbl.BandedColumns = false;
        } catch (e) {}

        // 设置列宽
        if (Array.isArray(columnWidths) && columnWidths.length > 0) {
          for (let c = 0; c < Math.min(columnWidths.length, tbl.Columns.Count); c++) {
            try {
              tbl.Columns.Item(c + 1).Width = Number(columnWidths[c]);
            } catch (e) {}
          }
        }

        // 填充数据与富文本分段加粗
        if (Array.isArray(data) && data.length > 0) {
          for (let r = 0; r < Math.min(data.length, tbl.Rows.Count); r++) {
            const rowData = data[r];
            if (Array.isArray(rowData)) {
              for (let c = 0; c < Math.min(rowData.length, tbl.Columns.Count); c++) {
                const cell = tbl.Cell(r + 1, c + 1);
                const cellText = String(rowData[c] || "");
                const tr = cell.Shape.TextFrame.TextRange;
                tr.Text = cellText;
                tr.Font.Name = "Microsoft YaHei";

                const hFontSize = Number(headerFontSize) || 16;
                const bFontSize = Number(bodyFontSize) || 14;

                // 垂直居中与内边距规范
                try {
                  cell.Shape.TextFrame.VerticalAnchor = 3; // msoAnchorMiddle 垂直居中
                  cell.Shape.TextFrame.MarginLeft = 5;
                  cell.Shape.TextFrame.MarginRight = 5;
                  cell.Shape.TextFrame.MarginTop = 3;
                  cell.Shape.TextFrame.MarginBottom = 3;
                  cell.Shape.TextFrame.WordWrap = true;
                } catch (e) {}

                if (r === 0) {
                  // 表头标题：16pt 白色加粗居中
                  cell.Shape.Fill.Solid();
                  cell.Shape.Fill.ForeColor.RGB = hexToPptColor(headerFillColor || "#0072C6");
                  tr.Font.Size = hFontSize;
                  tr.Font.Bold = true;
                  tr.Font.Color.RGB = hexToPptColor("#FFFFFF");
                  tr.ParagraphFormat.Alignment = 2; // 居中
                } else {
                  // 数据行底色纯白
                  cell.Shape.Fill.Solid();
                  cell.Shape.Fill.ForeColor.RGB = 16777215; // 0xFFFFFF 纯白
                  tr.Font.Color.RGB = hexToPptColor("#000000");

                  if (c === 0) {
                    // 第一列：指标名称 14pt 居中加粗
                    tr.Font.Size = bFontSize;
                    tr.Font.Bold = true;
                    tr.ParagraphFormat.Alignment = 2;
                  } else if (c === 3) {
                    // 第四列：判定状态列 15.5pt 加粗居中
                    tr.Font.Size = bFontSize + 1.5;
                    tr.Font.Bold = true;
                    tr.ParagraphFormat.Alignment = 2;
                    if (cellText === "达成") {
                      tr.Font.Color.RGB = hexToPptColor("#009132");
                    } else if (cellText === "进行中") {
                      tr.Font.Color.RGB = hexToPptColor("#E36C09");
                    } else if (cellText === "未开展") {
                      tr.Font.Color.RGB = hexToPptColor("#FF202E");
                    }
                  } else {
                    // 内容列：统一 14pt，首句分段精准加粗，破折号居中
                    tr.Font.Size = bFontSize;
                    if (cellText === "——") {
                      tr.ParagraphFormat.Alignment = 2;
                    } else {
                      tr.ParagraphFormat.Alignment = 1;
                    }

                    const sepIdx = cellText.search(/[；;:：]/);
                    if (sepIdx > 0 && sepIdx < 35) {
                      try {
                        const part1 = tr.Characters(1, sepIdx + 1);
                        part1.Font.Bold = true;
                        part1.Font.Size = bFontSize;
                        const part2 = tr.Characters(sepIdx + 2);
                        part2.Font.Bold = false;
                        part2.Font.Size = bFontSize;
                      } catch (e) {}
                    } else if (cellText.startsWith("最高认证效率")) {
                      try {
                        tr.Font.Bold = true;
                        tr.Font.Size = bFontSize;
                      } catch (e) {}
                    } else {
                      tr.Font.Bold = false;
                      tr.Font.Size = bFontSize;
                    }
                  }
                }
              }
            }
          }
        }

        // 动态阶梯行高矩阵优化
        if (Array.isArray(rowHeights) && rowHeights.length > 0) {
          for (let r = 0; r < Math.min(rowHeights.length, tbl.Rows.Count); r++) {
            try {
              tbl.Rows.Item(r + 1).Height = Number(rowHeights[r]);
            } catch (e) {}
          }
        } else {
          try {
            tbl.Rows.Item(1).Height = 30;
            for (let r = 2; r <= tbl.Rows.Count; r++) {
              tbl.Rows.Item(r).Height = 46;
            }
          } catch (e) {}
        }

        // 应用网格实线边框
        applyCellBorders(tbl, borderColor || "#333333");

        return {
          success: true,
          shapeId: tableShape.Id,
          slideIndex: idx,
          rows: rCount,
          columns: cCount,
          message: `已在第 ${idx} 页成功创建 8 行 5 列表格（含细实线网格、精准列宽与分段加粗）`
        };
      }

      case "read_table": {
        const shp = findTableShape();
        if (!shp) throw new Error(`未在第 ${idx} 页找到表格形状`);
        const tbl = shp.Table;
        const tableData = [];
        const rCount = tbl.Rows.Count;
        const cCount = tbl.Columns.Count;

        for (let r = 1; r <= rCount; r++) {
          const rowArr = [];
          for (let c = 1; c <= cCount; c++) {
            try {
              const cell = tbl.Cell(r, c);
              const txt = cell.Shape && cell.Shape.HasTextFrame && cell.Shape.TextFrame.HasText
                ? cell.Shape.TextFrame.TextRange.Text.trim()
                : "";
              rowArr.push(txt);
            } catch (e) {
              rowArr.push("");
            }
          }
          tableData.push(rowArr);
        }

        return {
          success: true,
          shapeId: shp.Id,
          slideIndex: idx,
          rows: rCount,
          columns: cCount,
          data: tableData,
          message: `已成功读取第 ${idx} 页表格数据 (${rCount} 行 ${cCount} 列)`
        };
      }

      case "set_cell_text": {
        const shp = findTableShape();
        if (!shp) throw new Error(`未在第 ${idx} 页找到表格形状`);
        const tbl = shp.Table;
        const rIdx = Number(row) || 1;
        const cIdx = Number(column) || 1;
        const cell = tbl.Cell(rIdx, cIdx);

        if (text !== undefined) {
          cell.Shape.TextFrame.TextRange.Text = String(text);
          cell.Shape.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
        }
        if (fontSize) cell.Shape.TextFrame.TextRange.Font.Size = Number(fontSize);
        if (fontColor) cell.Shape.TextFrame.TextRange.Font.Color.RGB = hexToPptColor(fontColor);
        if (fontBold !== undefined) cell.Shape.TextFrame.TextRange.Font.Bold = Boolean(fontBold);
        if (fillColor) {
          cell.Shape.Fill.Solid();
          cell.Shape.Fill.ForeColor.RGB = hexToPptColor(fillColor);
        }

        return {
          success: true,
          shapeId: shp.Id,
          row: rIdx,
          column: cIdx,
          message: `已成功更新单元格 [${rIdx}, ${cIdx}] 的内容与格式`
        };
      }

      case "style_table": {
        const shp = findTableShape();
        if (!shp) throw new Error(`未在第 ${idx} 页找到表格形状`);
        const tbl = shp.Table;
        const hColor = headerFillColor || "#0072C6";

        for (let r = 1; r <= tbl.Rows.Count; r++) {
          for (let c = 1; c <= tbl.Columns.Count; c++) {
            const cell = tbl.Cell(r, c);
            if (r === 1) {
              cell.Shape.Fill.Solid();
              cell.Shape.Fill.ForeColor.RGB = hexToPptColor(hColor);
              if (cell.Shape.HasTextFrame && cell.Shape.TextFrame.HasText) {
                cell.Shape.TextFrame.TextRange.Font.Bold = true;
                cell.Shape.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#FFFFFF");
              }
            } else {
              cell.Shape.Fill.Solid();
              cell.Shape.Fill.ForeColor.RGB = hexToPptColor("#FFFFFF");
            }
          }
        }

        applyCellBorders(tbl, borderColor || "#333333");

        return {
          success: true,
          shapeId: shp.Id,
          headerColor: hColor,
          message: `已成功应用专业商业表格配色主题与边框网格`
        };
      }

      default:
        throw new Error(`未知的表格操作: ${action}`);
    }
  }

  function pptCaptureSlidePreview(app, params) {
    const { presentationName, slideIndex } = params || {};
    const pres = getPptPresentation(app, presentationName);
    const pptApp = getPptApp() || app;
    const total = pres.Slides.Count;
    let idx = Number(slideIndex);
    if (!idx || isNaN(idx)) {
      try {
        if (pptApp.ActiveWindow && pptApp.ActiveWindow.Selection && pptApp.ActiveWindow.Selection.SlideRange) {
          idx = pptApp.ActiveWindow.Selection.SlideRange.SlideIndex;
        }
      } catch (e) {}
      if (!idx) idx = 1;
    }
    if (total === 0) {
      throw new Error(`capture_slide_preview 失败：演示文稿 [${pres.Name}] 一页都没有，没有可导出的幻灯片`);
    }
    if (Math.floor(idx) !== idx || idx < 1 || idx > total) {
      throw new Error(`capture_slide_preview 的 slideIndex=${slideIndex} 越界：当前文稿 [${pres.Name}] 共 ${total} 页，有效范围 1~${total}`);
    }
    const slide = pres.Slides.Item(idx);

    const primaryPath = params.outputPath ? String(params.outputPath) : null;
    if (!primaryPath) throw new Error("缺少 Bridge 指定的预览输出路径（outputPath）");
    // 同一 API 用脚本 `slide.Export(path,"PNG",1280,720)` 实测可用，但工具路径历史上 100% 报
    // "PPT 未生成预览"（问题台账 ISS-76/ISS-86）：桥接侧把"文件没落盘"折成了兜底文案，
    // 宿主原始错误被丢掉。这里改为**两条导出路径依次尝试**，并把每次尝试的宿主原始报错全部回传。
    const attempts = [];
    const tryExport = (path, withSize) => {
      try {
        if (withSize) slide.Export(path, "PNG", 1280, 720);
        else slide.Export(path, "PNG");
        attempts.push({ path: path, scale: withSize ? "1280x720" : "default", ok: true });
        return true;
      } catch (e) {
        attempts.push({ path: path, scale: withSize ? "1280x720" : "default", ok: false, hostError: String(e && e.message ? e.message : e) });
        return false;
      }
    };

    const exportedPrimary = tryExport(primaryPath, true) || tryExport(primaryPath, false);
    const exported = [primaryPath];

    // 可选第二落点：调用方显式指定的 outputPath（网关转发的临时路径为 userOutputPath，
    // 两者不同时一并写出，便于调用方自行取图）
    const altPath = params.userOutputPath ? String(params.userOutputPath) : null;
    if (altPath && altPath !== primaryPath) {
      tryExport(altPath, true) || tryExport(altPath, false);
      exported.push(altPath);
    }

    const hostErrors = attempts.filter(a => !a.ok).map(a => `${a.path}（${a.scale}）: ${a.hostError}`);
    // 加载项无文件系统访问，无法确认落盘；把原始返回完整回传，由桥接侧判定文件是否存在（ISS-86）
    return {
      success: exportedPrimary,
      presentationName: pres.Name,
      slideIndex: idx,
      slideCount: total,
      imagePath: primaryPath,
      exportedPaths: exported,
      hasImage: exportedPrimary,
      hostError: hostErrors.length ? hostErrors.join(" | ") : undefined,
      attempts: attempts,
      message: exportedPrimary
        ? `已导出第 ${idx} 页幻灯片预览图至: ${primaryPath}`
        : `第 ${idx} 页幻灯片导出未成功；宿主对 ${attempts.length} 次导出尝试的原始报错见 hostError 与 attempts`
    };
  }
