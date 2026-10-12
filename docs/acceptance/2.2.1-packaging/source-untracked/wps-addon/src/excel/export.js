  // export.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  // 10.25 打印与分页设置（CAP-01）
  /**
   * 读写工作表的打印设置：打印区域、打印标题行列、纸张、方向、页边距、页眉页脚、
   * 网格线/居中/缩放，以及手动分页符。
   *
   * 依据：真机探测确认 `PageSetup` 上这些属性都存在（PrintArea / PrintTitleRows / Orientation /
   * PaperSize / Zoom / CenterHorizontally / PrintGridlines / 各边距 / CenterHeader / LeftFooter /
   * FitToPagesWide），`HPageBreaks` / `VPageBreaks` 也可用。
   * 此前 AI 做不出"可直接打印装订"的报表。
   *
   * `action: "read"` 只读回当前设置；默认 `"apply"` 先写后**逐项读回核对**。
   */
  function configurePrintLayout(app, params) {
    const {
      sheetName, workbookName, action = "apply",
      printArea, printTitleRows, printTitleColumns,
      orientation, paperSize, zoom, fitToPagesWide, fitToPagesTall,
      centerHorizontally, centerVertically, printGridlines,
      leftMargin, rightMargin, topMargin, bottomMargin, headerMargin, footerMargin,
      centerHeader, leftHeader, rightHeader, centerFooter, leftFooter, rightFooter,
      clearPrintArea, addHorizontalPageBreak, addVerticalPageBreak, clearPageBreaks
    } = params || {};

    const sheet = getWorksheet(app, sheetName, workbookName);
    const ps = sheet.PageSetup;

    // VBA 常量：方向 xlPortrait=1 / xlLandscape=2；纸张 xlPaperLetter=1 / xlPaperA4=9 / xlPaperA3=8 …
    const ORIENT = { portrait: 1, landscape: 2 };
    const PAPER = { letter: 1, a4: 9, a3: 8, legal: 5, b5: 13 };
    const applied = {};
    const warnings = [];

    if (action !== "read") {
      if (clearPrintArea) {
        ps.PrintArea = "";
        ps.PrintTitleRows = "";
        ps.PrintTitleColumns = "";
        applied.clearedPrintArea = true;
      }
      if (printArea !== undefined) { ps.PrintArea = printArea === null ? "" : String(printArea); applied.printArea = String(printArea); }
      if (printTitleRows !== undefined) { ps.PrintTitleRows = printTitleRows === null ? "" : String(printTitleRows); applied.printTitleRows = String(printTitleRows); }
      if (printTitleColumns !== undefined) { ps.PrintTitleColumns = printTitleColumns === null ? "" : String(printTitleColumns); applied.printTitleColumns = String(printTitleColumns); }
      if (orientation !== undefined) {
        const code = typeof orientation === "number" ? orientation : ORIENT[String(orientation).toLowerCase()];
        if (code === undefined) throw new Error(`不支持的 orientation: ${orientation}（可用 portrait / landscape，或 1 / 2）`);
        ps.Orientation = code;
        applied.orientation = code;
      }
      if (paperSize !== undefined) {
        const code = typeof paperSize === "number" ? paperSize : PAPER[String(paperSize).toLowerCase()];
        if (code === undefined) throw new Error(`不支持的 paperSize: ${paperSize}（可用 a4 / a3 / letter / legal / b5，或 VBA 常量）`);
        ps.PaperSize = code;
        applied.paperSize = code;
      }
      if (zoom !== undefined) { ps.Zoom = Number(zoom); applied.zoom = Number(zoom); }
      if (fitToPagesWide !== undefined) { ps.FitToPagesWide = Number(fitToPagesWide); applied.fitToPagesWide = Number(fitToPagesWide); }
      if (fitToPagesTall !== undefined) { ps.FitToPagesTall = Number(fitToPagesTall); applied.fitToPagesTall = Number(fitToPagesTall); }
      if (centerHorizontally !== undefined) { ps.CenterHorizontally = Boolean(centerHorizontally); applied.centerHorizontally = Boolean(centerHorizontally); }
      if (centerVertically !== undefined) { ps.CenterVertically = Boolean(centerVertically); applied.centerVertically = Boolean(centerVertically); }
      if (printGridlines !== undefined) { ps.PrintGridlines = Boolean(printGridlines); applied.printGridlines = Boolean(printGridlines); }
      for (const [key, value] of Object.entries({ leftMargin, rightMargin, topMargin, bottomMargin, headerMargin, footerMargin })) {
        if (value === undefined) continue;
        ps[key.charAt(0).toUpperCase() + key.slice(1)] = Number(value);
        applied[key] = Number(value);
      }
      for (const [key, value] of Object.entries({ centerHeader, leftHeader, rightHeader, centerFooter, leftFooter, rightFooter })) {
        if (value === undefined) continue;
        ps[key.charAt(0).toUpperCase() + key.slice(1)] = String(value);
        applied[key] = String(value);
      }

      if (clearPageBreaks) {
        try { sheet.ResetAllPageBreaks(); applied.clearedPageBreaks = true; } catch (e) { warnings.push(`清除分页符失败: ${e.message}`); }
      }
      if (Number.isFinite(Number(addHorizontalPageBreak))) {
        try { sheet.HPageBreaks.Add(sheet.Rows.Item(Number(addHorizontalPageBreak))); } catch (e) { warnings.push(`添加水平分页符失败: ${e.message}`); }
      }
      if (addVerticalPageBreak !== undefined) {
        if (!Number.isFinite(Number(addVerticalPageBreak))) {
          throw new Error(`addVerticalPageBreak 需传列号（数字），收到 ${JSON.stringify(addVerticalPageBreak)}；列字母请先换算成序号`);
        }
        try { sheet.VPageBreaks.Add(sheet.Columns.Item(Number(addVerticalPageBreak))); } catch (e) { warnings.push(`添加垂直分页符失败: ${e.message}`); }
      }
    }

    // 读回：打印设置最怕"设了但没生效"
    const safe = (fn) => { try { const v = fn(); return v === undefined ? null : v; } catch (e) { return null; } };
    const settings = {
      printArea: safe(() => String(ps.PrintArea || "")),
      printTitleRows: safe(() => String(ps.PrintTitleRows || "")),
      printTitleColumns: safe(() => String(ps.PrintTitleColumns || "")),
      orientation: safe(() => Number(ps.Orientation)),
      paperSize: safe(() => Number(ps.PaperSize)),
      zoom: safe(() => Number(ps.Zoom)),
      fitToPagesWide: safe(() => Number(ps.FitToPagesWide)),
      fitToPagesTall: safe(() => Number(ps.FitToPagesTall)),
      centerHorizontally: safe(() => Boolean(ps.CenterHorizontally)),
      centerVertically: safe(() => Boolean(ps.CenterVertically)),
      printGridlines: safe(() => Boolean(ps.PrintGridlines)),
      margins: {
        left: safe(() => Number(ps.LeftMargin)), right: safe(() => Number(ps.RightMargin)),
        top: safe(() => Number(ps.TopMargin)), bottom: safe(() => Number(ps.BottomMargin)),
        header: safe(() => Number(ps.HeaderMargin)), footer: safe(() => Number(ps.FooterMargin))
      },
      headers: {
        center: safe(() => String(ps.CenterHeader || "")), left: safe(() => String(ps.LeftHeader || "")),
        right: safe(() => String(ps.RightHeader || ""))
      },
      footers: {
        center: safe(() => String(ps.CenterFooter || "")), left: safe(() => String(ps.LeftFooter || "")),
        right: safe(() => String(ps.RightFooter || ""))
      },
      pageBreaks: {
        horizontal: safe(() => Number(sheet.HPageBreaks.Count)),
        vertical: safe(() => Number(sheet.VPageBreaks.Count))
      }
    };
    settings.orientationName = settings.orientation === 1 ? "portrait" : settings.orientation === 2 ? "landscape" : null;

    // 打印区域写后核对（最容易"设了却没生效"的一项）
    if (action !== "read" && printArea !== undefined && printArea !== null && String(printArea) !== "") {
      const norm = (s) => String(s || "").toLowerCase().replace(/\$/g, "");
      if (norm(printArea) !== norm(settings.printArea)) {
        throw new Error(`打印区域设置未生效：请求 "${printArea}"，读回 "${settings.printArea}"`);
      }
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      action,
      applied,
      settings,
      warnings,
      message: action === "read"
        ? `已读回 [${sheet.Name}] 的打印设置（打印区域 ${settings.printArea || "未设置"}，${settings.orientationName || "方向未知"}）`
        : `已应用并读回核对 [${sheet.Name}] 的打印设置`
    };
  }

  // 10.26 导出工作表/工作簿为 PDF（CAP-02）
  /**
   * 用宿主的 `ExportAsFixedFormat` 导出 PDF（xlTypePDF = 0）。
   *
   * **加载项没有文件系统访问，无法确认落盘**——本函数只负责发起导出并回传路径与耗时，
   * **落盘校验由桥接侧完成**（与 Word 预览同一套机制）。输出目录必须是宿主可写位置，
   * 见 `src/bridge/runtime.ts` 的 `previewDir()`：WPS 是沙箱应用，`~/.wps-bridge` 之类写不进去。
   */
  function exportSheetPdf(app, params) {
    const { sheetName, workbookName, outputPath, scope = "workbook", quality = "standard" } = params || {};
    if (!outputPath) throw new Error("缺少必要参数: outputPath（由桥接指定的输出路径）");
    // Quality: xlQualityStandard = 0 / xlQualityMinimum = 1
    const qualityCode = String(quality).toLowerCase() === "minimum" ? 1 : 0;
    const target = scope === "sheet" ? getWorksheet(app, sheetName, workbookName) : getWorkbook(app, workbookName);

    const startedAt = Date.now();
    try {
      // ExportAsFixedFormat(Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish)
      target.ExportAsFixedFormat(0, String(outputPath), qualityCode, true, false, undefined, undefined, false);
    } catch (e) {
      return { success: false, outputPath: String(outputPath), scope, hostError: e.message, message: `导出 PDF 失败：${e.message}` };
    }
    return {
      success: true,
      outputPath: String(outputPath),
      scope,
      elapsedMs: Date.now() - startedAt,
      hostCannotVerify: true,
      message: `宿主已接受导出请求（${scope === "sheet" ? "工作表 " + (sheetName || "活动表") : "整个工作簿"} → ${outputPath}）；文件是否真的写出由桥接侧校验`
    };
  }

  // 14. 捕获工作表或指定区域的渲染预览图 (原生 JSA 高保真推入剪贴板)
  function captureSheetPreview(app, params) {
    const { sheetName, address, range, workbookName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);

    try {
      sheet.Activate();
    } catch (e) {}

    const targetAddr = address || range;
    let targetRange;
    if (targetAddr) {
      targetRange = sheet.Range(targetAddr);
    } else {
      // 默认范围必须是"**整张画布**"，不能只用 UsedRange——
      // UsedRange 只统计**有数据的单元格，不含形状**。当一页全是矢量图形、
      // 单元格没有任何值时，UsedRange 会退化成 $A$1，截出来只有 146x50 的废图
      //（真机踩到过）。所以这里把"所有形状的外框"并入 UsedRange。
      const ur = sheet.UsedRange;
      let r1 = ur.Row, c1 = ur.Column;
      let r2 = r1 + ur.Rows.Count - 1, c2 = c1 + ur.Columns.Count - 1;
      try {
        const shapes = sheet.Shapes;
        for (let i = 1; i <= shapes.Count; i++) {
          const sh = shapes.Item(i);
          let tl = null, br = null;
          try { tl = sh.TopLeftCell; br = sh.BottomRightCell; } catch (e) {}
          if (!tl || !br) continue;
          r1 = Math.min(r1, tl.Row); c1 = Math.min(c1, tl.Column);
          r2 = Math.max(r2, br.Row); c2 = Math.max(c2, br.Column);
        }
      } catch (e) {}
      // 注意：WPS JSA 的 Worksheet **没有 `Cells()` 方法**（Excel VBA 有），
      // 调它会报 "sheet.Cells is not a function"。这里自己拼 A1 地址。
      const colName = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
      targetRange = sheet.Range(`${colName(c1)}${r1}:${colName(c2)}${r2}`);
    }

    // 执行高保真选区图形渲染并推入系统剪贴板 (1 = xlScreen, -4147 = xlBitmap)
    targetRange.CopyPicture(1, -4147);

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: targetRange.Address()
    };
  }
