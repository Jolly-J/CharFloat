  // format.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  // 10. 高级格式与排版引擎 (美学升级)
  function formatCells(app, params) {
    // 本次调用的告警（原先函数内没有这个变量，我加读回核对时漏了声明 → ReferenceError）
    const warnings = [];
    const {
      sheetName,
      address,
      workbookName,
      fontName,
      fontSize,
      bold,
      fontColor,
      backgroundColor,
      horizontalAlignment,
      verticalAlignment = "center",
      numberFormat,
      rowHeight,
      wrapText,
      borders
    } = params;

    if (!address) throw new Error("缺少 address 参数");

    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);

    // 1. 字体与大小
    if (fontName) range.Font.Name = fontName;
    if (fontSize) range.Font.Size = fontSize;
    if (bold !== undefined) range.Font.Bold = !!bold;

    // 2. 颜色
    // 非法颜色值原先被 `if (bgr !== null)` 静默吞掉——调用方以为设上了（excel-tester L-1）。
    // 现在解析不出来就写 warnings，让"没生效"可见。
    if (backgroundColor) {
      const bgr = hexToExcelColor(backgroundColor);
      if (bgr !== null) range.Interior.Color = bgr;
      else warnings.push(`backgroundColor 无法解析为颜色: ${JSON.stringify(backgroundColor)}（应为 #RRGGBB 或 RRGGBB），该项已跳过`);
    }
    if (fontColor) {
      const bgr = hexToExcelColor(fontColor);
      if (bgr !== null) range.Font.Color = bgr;
      else warnings.push(`fontColor 无法解析为颜色: ${JSON.stringify(fontColor)}（应为 #RRGGBB 或 RRGGBB），该项已跳过`);
    }

    // 3. 对齐
    if (horizontalAlignment) {
      if (horizontalAlignment === "center") range.HorizontalAlignment = -4108; // xlCenter
      else if (horizontalAlignment === "left") range.HorizontalAlignment = -4131; // xlLeft
      else if (horizontalAlignment === "right") range.HorizontalAlignment = -4152; // xlRight
    }
    if (verticalAlignment !== undefined && verticalAlignment !== null && verticalAlignment !== "") {
      // xlTop=-4160 / xlCenter=-4108 / xlBottom=-4107 / xlJustify=-4130
      const VA = { top: -4160, center: -4108, middle: -4108, bottom: -4107, justify: -4130 };
      const code = VA[String(verticalAlignment).toLowerCase()];
      if (code === undefined) {
        warnings.push(`不认识的 verticalAlignment: ${verticalAlignment}（可用 top / center / bottom / justify）`);
      } else {
        range.VerticalAlignment = code;
        // **读回核对**：原先只处理 center，top/bottom 直接落空，
        // 而返回体照样回显请求值——"谎报成功"（excel-tester H-2）。
        let actual = null;
        try { actual = Number(range.VerticalAlignment); } catch (e) {}
        if (actual !== null && actual !== code) {
          warnings.push(`verticalAlignment 未生效：请求 ${verticalAlignment}(${code})，宿主读回 ${actual}`);
        }
      }
    }

    // 4. 换行与行高
    if (wrapText !== undefined) {
      range.WrapText = !!wrapText;
    }
    if (rowHeight) {
      // 本机 WPS 上给跨多行的 `range.RowHeight` 赋值会被静默忽略（问题台账 ISS-22）。
      // 改为逐行写 Rows.Item(n).RowHeight，并读回校验。
      // 注意：本函数的区域变量名是 `range`（不是 targetRange）——写成 targetRange 会抛
      // ReferenceError，而 npm test 未覆盖该路径，只会在真机调用时暴露。
      const firstRow = range.Row;
      const rowTotal = range.Rows.Count;
      for (let i = 0; i < rowTotal; i++) {
        sheet.Rows.Item(firstRow + i).RowHeight = Number(rowHeight);
      }
      let readBackHeight = null;
      try { readBackHeight = Number(sheet.Rows.Item(firstRow).RowHeight); } catch (e) {}
      if (readBackHeight !== null && Number.isFinite(readBackHeight) && Math.abs(readBackHeight - Number(rowHeight)) > 0.6) {
        throw new Error(
          `行高设置未生效：请求 ${rowHeight}，第 ${firstRow} 行读回 ${readBackHeight}。` +
          `请确认该行未被合并单元格或工作表保护锁定。`
        );
      }
    }

    // 5. 数字格式
    if (numberFormat) {
      range.NumberFormat = numberFormat;
    }

    // 6. 精细浅灰边框
    if (borders) {
      try {
        const borderBgr = hexToExcelColor(typeof borders === "string" ? borders : "#CBD5E1");
        // 边框: xlEdgeLeft=7, xlEdgeTop=8, xlEdgeBottom=9, xlEdgeRight=10, xlInsideHorizontal=11, xlInsideVertical=12
        [7, 8, 9, 10, 11, 12].forEach((bId) => {
          try {
            const b = range.Borders.Item(bId);
            b.LineStyle = 1; // xlContinuous
            b.Weight = 2; // xlThin
            if (borderBgr !== null) b.Color = borderBgr;
          } catch (e) {}
        });
      } catch (e) {}
    }

    // 7. 合并 / 拆分单元格
    // 执行后必须读回校验：原来 catch 里只 log，失败也会返回 success —— 调用方以为合并成功（问题台账 ISS-42）。
    if (params.merge === true || params.unmerge === true) {
      const wantMerged = params.merge === true;
      const actionLabel = wantMerged ? "合并" : "取消合并";
      let failure = null;
      try {
        if (wantMerged) range.Merge(); else range.UnMerge();
      } catch (e) {
        failure = e.message;
      }
      let mergedNow = null;
      try { mergedNow = range.MergeCells === true; } catch (e) { mergedNow = null; }
      if (failure) throw new Error(`${actionLabel}失败：${failure}`);
      if (mergedNow !== null && mergedNow !== wantMerged) {
        throw new Error(`${actionLabel}未生效（读回 merged=${mergedNow}）；若目标区域与已有合并区部分重叠，请先取消原有合并`);
      }
    }

    return {
      warnings,
      success: true,
      address: range.Address(),
      appliedStyles: params
    };
  }

  // 10.05 单元格内**局部**格式（富文本）—— CAP-03
  /**
   * 只给单元格里的一段文字加格式（如"一句话里只把某段加粗/变红/改字号"）。
   *
   * 依据：WPS 宿主有 `Range.Characters(Start, Length)`（真机探测确认：返回对象且 `Count` 可用）。
   * 在此之前 AI 只能整格统一格式，做不到局部强调——这是"精细操控文档"最典型的诉求之一。
   *
   * 定位方式二选一：`find`（按文本找第 occurrence 处）或 `start`（1 基起点）+ `length`。
   * 写完逐项读回核对，任何一项没落上都报错，不做静默降级。
   */
  function formatTextSegment(app, params) {
    const {
      sheetName, workbookName, address,
      find, occurrence = 1, start, length,
      bold, italic, underline, fontColor, fontSize, fontName
    } = params || {};
    if (!address) throw new Error("缺少必要参数: address (如 'A1')");

    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);

    let text = "";
    try { text = range.Value2 === null || range.Value2 === undefined ? "" : String(range.Value2); } catch (e) {}
    if (!text) throw new Error(`单元格 ${range.Address ? range.Address() : address} 没有文本内容，无法做局部格式`);

    let begin = null;
    let segLen;
    if (find !== undefined && find !== null && String(find) !== "") {
      const query = String(find);
      const want = Math.max(1, Number(occurrence) || 1);
      let from = 0, hit = 0, idx = -1;
      for (;;) {
        idx = text.indexOf(query, from);
        if (idx < 0) break;
        hit++;
        if (hit === want) break;
        from = idx + 1;
      }
      if (idx < 0) {
        throw new Error(
          `在 ${address} 中找不到第 ${want} 处 "${query}"（共找到 ${hit} 处，单元格文本长度 ${text.length}）。` +
          `当前文本前 60 字：${text.slice(0, 60)}`
        );
      }
      begin = idx + 1; // Characters 是 1 基
      segLen = query.length;
    } else {
      if (!Number.isFinite(Number(start))) {
        throw new Error("需要提供 find（按文本定位）或 start（1 基起始位置，配合 length 使用）");
      }
      begin = Number(start);
      if (Number.isFinite(Number(length))) segLen = Number(length);
    }

    const chars = segLen === undefined ? range.Characters(begin) : range.Characters(begin, segLen);
    if (!chars) throw new Error(`无法定位字符片段（start=${begin}, length=${segLen}）：宿主未返回对象`);

    const requested = {};
    // VBA/WPS 的布尔格式用 -1/0（msoTrue/msoFalse）；传 JS 布尔在部分宿主上会被忽略
    if (bold !== undefined) { chars.Font.Bold = bold ? -1 : 0; requested.bold = Boolean(bold); }
    if (italic !== undefined) { chars.Font.Italic = italic ? -1 : 0; requested.italic = Boolean(italic); }
    if (underline !== undefined) { chars.Font.Underline = underline ? 2 : -4142; requested.underline = Boolean(underline); }
    if (fontSize !== undefined) { chars.Font.Size = Number(fontSize); requested.fontSize = Number(fontSize); }
    if (fontName) { chars.Font.Name = String(fontName); requested.fontName = String(fontName); }
    if (fontColor) {
      const bgr = hexToExcelColor(fontColor);
      if (bgr === null) throw new Error(`fontColor 无法解析: ${fontColor}（应为 #RRGGBB）`);
      chars.Font.Color = bgr;
      requested.fontColor = String(fontColor).toUpperCase();
    }

    // 读回核对：局部格式最怕"看着像生效了"
    const readBack = {};
    const mismatches = [];
    const num = (v) => (typeof v === "boolean" ? (v ? -1 : 0) : Number(v));
    try { readBack.segmentText = chars.Text === undefined ? null : String(chars.Text); } catch (e) { readBack.segmentText = null; }
    if (requested.bold !== undefined) {
      readBack.bold = num(safeRead(() => chars.Font.Bold, null)) !== 0;
      if (readBack.bold !== requested.bold) mismatches.push(`bold 请求 ${requested.bold} 读回 ${readBack.bold}`);
    }
    if (requested.italic !== undefined) {
      readBack.italic = num(safeRead(() => chars.Font.Italic, null)) !== 0;
      if (readBack.italic !== requested.italic) mismatches.push(`italic 请求 ${requested.italic} 读回 ${readBack.italic}`);
    }
    if (requested.fontSize !== undefined) {
      readBack.fontSize = Number(safeRead(() => chars.Font.Size, null));
      if (Math.abs(readBack.fontSize - requested.fontSize) > 0.26) mismatches.push(`fontSize 请求 ${requested.fontSize} 读回 ${readBack.fontSize}`);
    }
    if (requested.fontName !== undefined) {
      readBack.fontName = safeRead(() => chars.Font.Name, null);
      if (String(readBack.fontName) !== requested.fontName) mismatches.push(`fontName 请求 ${requested.fontName} 读回 ${readBack.fontName}`);
    }
    if (requested.fontColor !== undefined) {
      readBack.fontColor = excelColorToHex(safeRead(() => chars.Font.Color, null));
      if (String(readBack.fontColor).toUpperCase() !== requested.fontColor) mismatches.push(`fontColor 请求 ${requested.fontColor} 读回 ${readBack.fontColor}`);
    }
    if (requested.underline !== undefined) {
      readBack.underline = num(safeRead(() => chars.Font.Underline, null)) !== -4142;
      if (readBack.underline !== requested.underline) mismatches.push(`underline 请求 ${requested.underline} 读回 ${readBack.underline}`);
    }

    if (mismatches.length > 0) {
      throw new Error(
        `局部格式未完全生效：${mismatches.join("；")}。` +
        `该单元格文本为 "${text.slice(0, 60)}"，片段起点 ${begin}、长度 ${segLen === undefined ? "至末尾" : segLen}。`
      );
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address ? range.Address() : address,
      cellText: text.slice(0, 120),
      segmentStart: begin,
      segmentLength: segLen === undefined ? text.length - begin + 1 : segLen,
      segmentText: readBack.segmentText,
      requested,
      readBack,
      message: `已将 [${sheet.Name}] ${address} 的第 ${begin} 个字符起 ${segLen === undefined ? "到末尾" : segLen + " 个字符"}「${readBack.segmentText ?? ""}」设置为指定格式，并读回核对通过`
    };
  }

  // 10.1 条件格式与数据条/色阶
  function addConditionalFormatting(app, params) {
    // 本次调用的告警（原先函数内没有这个变量，我加读回核对时漏了声明 → ReferenceError）
    const warnings = [];
    // CAP-30 条件格式读回：此前只有"加"没有"读"，AI 无法回答
    // "这块区域现在挂了哪些规则"，改完也只能靠人看。
    // 覆盖单元格值 / 公式 / 色阶 / 数据条 / Top10 / 图标集 / 重复值 / 唯一值 / 文本 / 空值。
    if ((params || {}).action === "read") {
      const { sheetName, address, workbookName } = params || {};
      const sheet = getWorksheet(app, sheetName, workbookName);
      const TYPE = {
        1: "cell_value", 2: "formula", 3: "color_scale", 4: "data_bar", 5: "top10",
        6: "icon_set", 8: "unique_values", 9: "text_contains", 10: "blanks", 11: "time_period",
        12: "above_average", 13: "no_blanks", 16: "duplicate_values"
      };
      const OPER = { 1: "between", 2: "not_between", 3: "equal", 4: "not_equal", 5: "greater_than", 6: "less_than", 7: "greater_equal", 8: "less_equal" };
      const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
      const toHex = (v) => (v === null || v === undefined || v < 0 || v === 16777215) ? null
        : "#" + [v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff].map(n => n.toString(16).padStart(2, "0")).join("").toUpperCase();

      const ranges = address ? [sheet.Range(address)] : (() => {
        // 整表：宿主没有"枚举所有条件格式区域"的 API，按行扫描已用区域
        const ur = sheet.UsedRange, out = [];
        const colName = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
        const rows = Math.min(ur.Rows.Count, 2000), cols = Math.min(ur.Columns.Count, 100);
        for (let r = ur.Row; r < ur.Row + rows; r++) {
          for (let c = ur.Column; c < ur.Column + cols; c++) out.push(sheet.Range(`${colName(c)}${r}`));
        }
        return out;
      })();

      const items = [];
      for (const rng of ranges) {
        let fcs = null;
        try { fcs = rng.FormatConditions; } catch (e) { continue; }
        const n = g(() => Number(fcs.Count), 0);
        if (!n) continue;
        const entry = { address: g(() => rng.Address(), null), rules: [] };
        for (let i = 1; i <= n; i++) {
          const fc = (() => { try { return fcs.Item(i); } catch (e) { return null; } })();
          if (!fc) continue;
          const tc = g(() => Number(fc.Type), null);
          const rule = {
            index: i,
            type: TYPE[tc] || `unknown(${tc})`,
            typeCode: tc,
            operator: OPER[g(() => Number(fc.Operator), null)] || null,
            formula1: g(() => fc.Formula1, null),
            formula2: g(() => fc.Formula2, null),
            priority: g(() => Number(fc.Priority), null),
            stopIfTrue: g(() => Boolean(fc.StopIfTrue), null),
            fillColor: toHex(g(() => Number(fc.Interior.Color), null)),
            fontColor: toHex(g(() => Number(fc.Font.Color), null)),
            fontBold: g(() => Boolean(fc.Font.Bold), null),
            // 图标集 / 数据条 / 色阶的特有属性（宿主不支持的会落到 null）
            iconSet: g(() => String(fc.IconSet), null),
            showIconOnly: g(() => Boolean(fc.ShowIconOnly), null),
            barColor: toHex(g(() => Number(fc.BarColor), null)),
            text: g(() => fc.Text, null)
          };
          entry.rules.push(rule);
        }
        items.push(entry);
      }
      const total = items.reduce((s, x) => s + x.rules.length, 0);
      return {
        success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name,
        areaCount: items.length, ruleCount: total, areas: items, warnings: [],
        message: `工作表 [${sheet.Name}] 上 ${items.length} 个区域共 ${total} 条条件格式规则`
      };
    }
    const {
      sheetName,
      address,
      workbookName,
      ruleType = "cell_value",
      operator = "less_than",
      formula1,
      formula2,
      backgroundColor,
      fontColor,
      barColor,
      colorScaleMin,
      colorScaleMax,
      clearExisting = false,
      // CAP-06 新增：图标集 / Top-N / 重复值 / 公式 / 文字包含
      iconSet,
      iconThresholds,
      topBottom,
      topRank,
      topPercent,
      containsText
    } = params || {};

    if (!address) throw new Error("缺少 address 参数");
    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);

    // ⚠️ 与 set_data_validation 同一类问题（Lead 追修「先校验、后动手」）：
    // `clearExisting` 会**先删掉该区域既有条件格式**，而原有的参数校验分散在下面各分支里
    // （未知 ruleType / 不支持的 iconSet / text_contains 缺 containsText / formula 缺 formula1），
    // 于是"拼错一个参数 + clearExisting"会先把用户的规则清干净、再报错。
    // 这里把这几项校验**整体前移**到任何修改之前，报错文案与分支内保持一致。
    const ICON_SET_CODES = {
      "3_arrows": 1, "3_arrows_gray": 2, "3_flags": 3,
      "3_traffic_lights": 4, "3_traffic_lights_rimmed": 5,
      "3_signs": 6, "3_symbols": 7, "3_symbols_circled": 8,
      "4_arrows": 9, "4_arrows_gray": 10, "4_red_to_black": 11,
      "4_ratings": 12, "4_traffic_lights": 13,
      "5_arrows": 14, "5_arrows_gray": 15, "5_quarters": 16,
      "5_ratings": 17, "5_boxes": 18
    };
    const KNOWN_RULE_TYPES = ["cell_value", "data_bar", "color_scale", "icon_set", "top10", "duplicate_values", "unique_values", "formula", "text_contains", "clear"];
    const untouched = "；本次未做任何修改（原有条件格式保持不变，若传了 clearExisting 也**没有**执行清除）";
    if (KNOWN_RULE_TYPES.indexOf(String(ruleType)) < 0) {
      throw new Error(
        `未知的条件格式类型: ${ruleType}（支持 cell_value, data_bar, color_scale, icon_set, ` +
        `top10, duplicate_values, unique_values, formula, text_contains, clear）` + untouched
      );
    }
    if (ruleType === "icon_set") {
      const iconSetName = iconSet || "3_traffic_lights";
      if (ICON_SET_CODES[iconSetName] === undefined) {
        throw new Error(`不支持的 iconSet: ${iconSetName}（可用: ${Object.keys(ICON_SET_CODES).join(" / ")}）` + untouched);
      }
    }
    if (ruleType === "text_contains" && !containsText) {
      throw new Error("text_contains 规则必须提供 containsText" + untouched);
    }
    if (ruleType === "formula" && !formula1) {
      throw new Error("formula 规则必须提供 formula1（如 '=A1>100'）" + untouched);
    }
    if (ruleType === "cell_value" && operator === "between" && formula2 === undefined) {
      // between 少了上限时宿主 Add 会抛错，而那时条件格式已经被 Delete 掉了（同一破坏路径）
      throw new Error("operator='between' 必须同时提供 formula1 与 formula2（上限）" + untouched);
    }

    if (clearExisting) {
      try { range.FormatConditions.Delete(); } catch (e) {}
    }

    if (ruleType === "cell_value") {
      let xlOp = 6; // xlLess
      if (operator === "greater_than") xlOp = 5;
      else if (operator === "less_than") xlOp = 6;
      else if (operator === "equal") xlOp = 3;
      else if (operator === "between") xlOp = 1;

      const f1 = formula1 !== undefined ? String(formula1) : "0";
      const f2 = formula2 !== undefined ? String(formula2) : undefined;
      const fc = range.FormatConditions.Add(1, xlOp, f1, f2);

      if (backgroundColor) {
        const bg = hexToExcelColor(backgroundColor);
        if (bg !== null) fc.Interior.Color = bg;
      }
      if (fontColor) {
        const fg = hexToExcelColor(fontColor);
        if (fg !== null) fc.Font.Color = fg;
      }
    } else if (ruleType === "data_bar") {
      const db = range.FormatConditions.AddDatabar();
      if (barColor && db.BarColor) {
        const bc = hexToExcelColor(barColor);
        if (bc !== null) db.BarColor.Color = bc;
      }
    } else if (ruleType === "color_scale") {
      const cs = range.FormatConditions.AddColorScale(2);
      if (colorScaleMin && cs.ColorScaleCriteria) {
        const c1 = hexToExcelColor(colorScaleMin);
        if (c1 !== null) cs.ColorScaleCriteria.Item(1).FormatColor.Color = c1;
      }
      if (colorScaleMax && cs.ColorScaleCriteria) {
        const c2 = hexToExcelColor(colorScaleMax);
        if (c2 !== null) cs.ColorScaleCriteria.Item(2).FormatColor.Color = c2;
      }
    } else if (ruleType === "icon_set") {
      // CAP-06：红黄绿灯这类"业务信号"是最常见的报表诉求。
      // 真机探测：宿主有 AddIconSetCondition（AddTextString 不存在，文字规则走公式规则）。
      // ICON_SET_CODES 已在函数开头（校验前移时）声明，这里直接复用。
      const iconSetName = iconSet || "3_traffic_lights";
      const code = ICON_SET_CODES[iconSetName];
      if (code === undefined) {
        // 兜底（正常路径已在前面拦下）
        throw new Error(`不支持的 iconSet: ${iconSetName}（可用: ${Object.keys(ICON_SET_CODES).join(" / ")}）`);
      }
      const icons = range.FormatConditions.AddIconSetCondition();
      icons.IconSet = app.IconSets ? app.IconSets.Item(code) : code;
      // 阈值：不给就用宿主默认；给了就逐档设置（最多 5 档）
      if (Array.isArray(iconThresholds) && iconThresholds.length > 0) {
        for (let i = 0; i < Math.min(iconThresholds.length, 5); i++) {
          const spec = iconThresholds[i];
          const crit = icons.IconCriteria.Item(i + 2); // 第 1 档是"最低值"，从第 2 档开始可设阈值
          if (spec && typeof spec === "object") {
            if (spec.type !== undefined) crit.Type = Number(spec.type);
            if (spec.operator !== undefined) crit.Operator = Number(spec.operator);
            if (spec.value !== undefined) crit.Value = spec.value;
          } else if (spec !== undefined && spec !== null) {
            crit.Value = spec;
          }
        }
      }
      if (backgroundColor) {
        const bgc = hexToExcelColor(backgroundColor);
        if (bgc !== null) icons.Interior.Color = bgc;
      }
    } else if (ruleType === "top10") {
      // CAP-06：Top/Bottom N 或百分比
      const top = range.FormatConditions.AddTop10();
      if (topBottom !== undefined) top.TopBottom = Number(topBottom);   // 1=xlTop 2=xlBottom
      if (topRank !== undefined) top.Rank = Number(topRank);
      if (topPercent !== undefined) top.Percent = topPercent ? -1 : 0;
      if (backgroundColor) {
        const bgc = hexToExcelColor(backgroundColor);
        if (bgc !== null) top.Interior.Color = bgc;
      }
      if (fontColor) {
        const fgc = hexToExcelColor(fontColor);
        if (fgc !== null) top.Font.Color = fgc;
      }
    } else if (ruleType === "duplicate_values" || ruleType === "unique_values") {
      // CAP-06：重复值/唯一值高亮（对账、查重最常用）
      const dv = range.FormatConditions.AddUniqueValues();
      dv.DupeUnique = ruleType === "duplicate_values" ? 1 : 2; // 1=xlDuplicate 2=xlUnique
      if (backgroundColor) {
        const bgc = hexToExcelColor(backgroundColor);
        if (bgc !== null) dv.Interior.Color = bgc;
      }
      if (fontColor) {
        const fgc = hexToExcelColor(fontColor);
        if (fgc !== null) dv.Font.Color = fgc;
      }
    } else if (ruleType === "formula" || ruleType === "text_contains") {
      // CAP-06：公式规则（xlExpression=2）。文字包含没有独立宿主方法（真机确认 AddTextString 不存在），
      // 用标准的 SEARCH 公式实现——这也是业务上更通用的做法。
      let expr = formula1;
      if (ruleType === "text_contains") {
        if (!containsText) throw new Error("text_contains 规则必须提供 containsText");
        const anchor = String(address).split(":")[0];
        expr = `=ISNUMBER(SEARCH("${String(containsText).replace(/"/g, '""')}",${anchor}))`;
      }
      if (!expr) throw new Error("formula 规则必须提供 formula1（如 '=A1>100'）");
      // ⚠️ **必须先把活动单元格移到区域左上角**：
      // `FormatConditions.Add` 的公式用**相对引用**，基准是**当前活动单元格**而不是区域左上角。
      // 若活动单元格在别处，公式会被整体平移——同一调用随环境变语义，规则看着加上了却完全不生效
      // （excel-tester H-1：活动单元格=A1 时，给 M1:M5 加的规则变成 =ISNUMBER(SEARCH("特价",Y1))）。
      let fx;
      try {
        // 把活动单元格移到区域左上角，公式相对引用才以它为基准。
        // ⚠️ 不能用 `range.Cells(1,1)`：WPS JSA 的 **Range/Worksheet 都没有 Cells() 方法**
        // （真机踩到：抛错被吞掉，公式照样被平移）。改用**地址字符串**取左上角。
        const tl = String(address).split(":")[0].replace(/\$/g, "");
        sheet.Range(tl).Select();
      } catch (e) {
        warnings.push(`无法把活动单元格移到规则区域左上角（${e.message}），公式相对引用可能被平移`);
      }
      fx = range.FormatConditions.Add(2, 0, String(expr)); // 2=xlExpression, 0=xlNone
      // 读回公式，核对它是否被平移（这是 H-1 的判据）
      try {
        const readExpr = String(range.FormatConditions.Item(range.FormatConditions.Count).Formula1);
        if (readExpr && readExpr !== String(expr) && String(expr).indexOf("ISNUMBER") >= 0) {
          warnings.push(`规则公式被平移：请求 ${expr}，宿主读回 ${readExpr}`);
        }
      } catch (e) {}
      if (backgroundColor) {
        const bgc = hexToExcelColor(backgroundColor);
        if (bgc !== null) fx.Interior.Color = bgc;
      }
      if (fontColor) {
        const fgc = hexToExcelColor(fontColor);
        if (fgc !== null) fx.Font.Color = fgc;
      }
    } else if (ruleType === "clear") {
      range.FormatConditions.Delete();
    } else {
      throw new Error(
        `未知的条件格式类型: ${ruleType}（支持 cell_value, data_bar, color_scale, icon_set, ` +
        `top10, duplicate_values, unique_values, formula, text_contains, clear）`
      );
    }

    // 读回核对：条件格式是"只写不读"的重灾区，写完必须能确认到底加了几条
    let appliedCount = null;
    const conditions = [];
    try {
      const fc = range.FormatConditions;
      appliedCount = Number(fc.Count);
      for (let i = 1; i <= Math.min(appliedCount, 20); i++) {
        const rule = fc.Item(i);
        conditions.push({
          index: i,
          type: Number(safeRead(() => rule.Type, null)),
          enabled: readConditionEnabled(rule),
          priority: safeRead(() => Number(rule.Priority), null),
          formula1: safeRead(() => (rule.Formula1 === undefined ? null : String(rule.Formula1)), null)
        });
      }
    } catch (e) {}

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      ruleType: ruleType,
      appliedConditionCount: appliedCount,
      conditions,
      message: `已成功在 ${range.Address()} 应用 ${ruleType} 条件格式${appliedCount === null ? "" : `（该区域现有 ${appliedCount} 条规则）`}`
    };
  }
