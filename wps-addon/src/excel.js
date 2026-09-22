  // ---------------------------------------------------------------------------
  // excel.js — WPS 表格（Excel/ET）全部 RPC 实现
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
  // 2. 提取原表设计语言 (Design Token)
  function getStyleToken(app, params) {
    const { sheetName, sampleAddress = "A3", workbookName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(sampleAddress);

    const fontName = range.Font.Name || "微软雅黑";
    const fontSize = range.Font.Size || 11;
    const fontBold = !!range.Font.Bold;
    const fontColor = excelColorToHex(range.Font.Color);
    const headerBg = excelColorToHex(range.Interior.Color);

    // 尝试探测大标题
    let titleFontName = fontName;
    try {
      const titleCell = sheet.Range("A1");
      if (titleCell.Font.Name) titleFontName = titleCell.Font.Name;
    } catch (e) {}

    return {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      fontName: fontName,
      titleFontName: titleFontName,
      sampleFontSize: fontSize,
      sampleBold: fontBold,
      fontColor: fontColor,
      headerBackgroundColor: headerBg || "#1E3A8A"
    };
  }

  // 3. 新建或激活工作表 (显式锁定 workbookName)
  function createWorksheet(app, params) {
    // 表名校验（excel-tester M-4）：非法字符或超长名字宿主会**静默创建 SheetN** 并返回成功，
    // 调用方以为表叫自己给的名字，后续按名定位全部失败。这里先校验再创建。
    if (params && params.sheetName !== undefined && params.sheetName !== null && String(params.sheetName) !== "") {
      const want = String(params.sheetName);
      if (want.length > 31) throw new Error(`工作表名过长（${want.length} 字符，上限 31）："${want}"`);
      if (/[\\\/\?\*\[\]:]/.test(want)) throw new Error(`工作表名含非法字符（\\ / ? * [ ] :）："${want}"`);
      if (/^'.*'$/.test(want)) throw new Error(`工作表名不能以单引号开头或结尾："${want}"`);
    }
    const { sheetName, workbookName } = params;
    if (!sheetName) throw new Error("缺少 sheetName 参数");

    const wb = getWorkbook(app, workbookName);

    let targetSheet = null;
    try {
      targetSheet = wb.Worksheets.Item(sheetName);
    } catch (e) {}

    let isNew = false;
    if (!targetSheet) {
      const count = wb.Worksheets.Count;
      targetSheet = wb.Worksheets.Add(null, wb.Worksheets.Item(count));
      targetSheet.Name = sheetName;
      isNew = true;
    }

    try {
      wb.Activate();
      targetSheet.Activate();
    } catch (e) {}

    return {
      success: true,
      workbookName: wb.Name,
      sheetName: targetSheet.Name,
      isNew: isNew,
      sheetCount: wb.Worksheets.Count
    };
  }

  // 4. 删除指定工作表
  function deleteWorksheet(app, params) {
    const { sheetName, workbookName } = params;
    if (!sheetName) throw new Error("缺少 sheetName 参数");

    const wb = getWorkbook(app, workbookName);
    let targetSheet = null;
    try {
      targetSheet = wb.Worksheets.Item(sheetName);
    } catch (e) {
      return { success: true, workbookName: wb.Name, message: `工作表 ${sheetName} 已不存在，无需删除` };
    }

    if (wb.Worksheets.Count <= 1) {
      throw new Error(`工作簿 [${wb.Name}] 仅剩 1 个工作表，不能删除最后一个工作表`);
    }

    try { app.DisplayAlerts = false; } catch (e) {}
    try {
      targetSheet.Delete();
    } finally {
      try { app.DisplayAlerts = true; } catch (e) {}
    }

    return {
      success: true,
      workbookName: wb.Name,
      message: `已成功删除工作簿 [${wb.Name}] 中的工作表 [${sheetName}]`,
      sheetCount: wb.Worksheets.Count
    };
  }

  // 5. 清理指定区域
  function clearRange(app, params) {
    const { sheetName, address, workbookName } = params;
    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);
    range.Clear();
    return { success: true, workbookName: sheet.Parent.Name, clearedAddress: range.Address() };
  }

  // 6. 获取大纲
  function getSheetOutline(app, params) {
    const sheetName = typeof params === "string" ? params : params?.sheetName;
    const workbookName = typeof params === "object" ? params?.workbookName : null;
    const sheet = getWorksheet(app, sheetName, workbookName);
    const usedRange = sheet.UsedRange;

    if (!usedRange || usedRange.Rows.Count === 0 || usedRange.Columns.Count === 0) {
      return {
        sheetName: sheet.Name,
        isEmpty: true,
        usedRangeAddress: "",
        rowCount: 0,
        columnCount: 0,
        headerPreview: []
      };
    }

    const address = usedRange.Address();
    const rowCount = usedRange.Rows.Count;
    const colCount = usedRange.Columns.Count;

    const sampleRows = Math.min(rowCount, 3);
    const previewRange = sheet.Range(sheet.Cells.Item(usedRange.Row, usedRange.Column), sheet.Cells.Item(usedRange.Row + sampleRows - 1, usedRange.Column + colCount - 1));
    const rawValues = previewRange.Value2;

    let headerPreview = [];
    if (Array.isArray(rawValues)) {
      headerPreview = rawValues;
    } else if (rawValues !== undefined && rawValues !== null) {
      headerPreview = [[rawValues]];
    }

    return {
      sheetName: sheet.Name,
      isEmpty: false,
      usedRangeAddress: address,
      startRow: usedRange.Row,
      startColumn: usedRange.Column,
      rowCount: rowCount,
      columnCount: colCount,
      headerPreview: headerPreview,
      // ISS-94 / ISS-64：把"能写但读不回来"的工作表级状态一并带回，
      // 让 AI 写完保护/标签色/冻结/筛选/条件格式后能自检，而不是只能相信 success。
      sheetState: readSheetState(app, sheet)
    };
  }

  /**
   * 读回条件格式的"是否启用"。
   *
   * 真机实测：本机 WPS 的 `FormatCondition` **没有 `Enabled` 属性**（读到 `undefined`）。
   * 直接 `Boolean(rule.Enabled)` 会把它写成 `false`，让读回看起来像"规则被禁用了"——
   * **读回说谎比没有读回更糟**。属性不存在时如实返回 null（未知），不猜测。
   */
  function readConditionEnabled(rule) {
    try {
      const raw = rule.Enabled;
      if (raw === undefined || raw === null) return null;
      return Boolean(raw);
    } catch (e) {
      return null;
    }
  }

  /**
   * 工作表级读回（只读，逐项 try/catch）。
   *
   * 这些状态原本都只有"写"没有"读"（问题台账 ISS-94）：调用方 set 完拿不到任何证据。
   * 冻结窗格依赖 `app.ActiveWindow`，**只在目标表处于活动状态时才有意义**，
   * 因此非活动表返回 null 并说明原因，避免拿别人的窗口状态冒充本表状态。
   */
  function readSheetState(app, sheet) {
    const state = {};
    const safe = (fn, fallback) => { try { const v = fn(); return v === undefined ? fallback : v; } catch (e) { return fallback; } };

    state.protection = safe(() => ({
      protectContents: Boolean(sheet.ProtectContents),
      protectDrawingObjects: Boolean(sheet.ProtectDrawingObjects),
      protectionMode: Boolean(sheet.ProtectionMode)
    }), null);

    state.tabColor = safe(() => {
      const color = sheet.Tab && sheet.Tab.Color !== undefined ? Number(sheet.Tab.Color) : null;
      if (color === null || !Number.isFinite(color)) return null;
      return excelColorToHex(color);
    }, null);

    state.autoFilter = safe(() => ({
      filterMode: Boolean(sheet.AutoFilterMode),
      range: sheet.AutoFilter && sheet.AutoFilter.Range ? sheet.AutoFilter.Range.Address() : null
    }), null);

    state.freezePanes = safe(() => {
      const active = app.ActiveWindow;
      if (!active) return { available: false, reason: "宿主没有活动窗口" };
      let isTarget = false;
      try { isTarget = String(active.ActiveSheet && active.ActiveSheet.Name) === String(sheet.Name); } catch (e) { isTarget = false; }
      if (!isTarget) {
        return { available: false, reason: "冻结窗格属窗口状态，只有目标表处于活动状态时才能读；请先激活该表" };
      }
      return {
        available: true,
        freezePanes: Boolean(active.FreezePanes),
        splitRow: safe(() => Number(active.SplitRow), null),
        splitColumn: safe(() => Number(active.SplitColumn), null)
      };
    }, { available: false, reason: "读取活动窗口状态失败" });

    // 条件格式：按已用区域扫描，返回每个区域的规则摘要（类型/优先级/是否启用）
    state.conditionalFormats = safe(() => {
      const rules = [];
      const used = sheet.UsedRange;
      if (!used) return rules;
      const fc = used.FormatConditions;
      const count = fc && typeof fc.Count === "number" ? fc.Count : 0;
      for (let i = 1; i <= Math.min(count, 50); i++) {
        try {
          const rule = fc.Item(i);
          rules.push({
            index: i,
            type: safe(() => Number(rule.Type), null),
            enabled: readConditionEnabled(rule),
            priority: safe(() => Number(rule.Priority), null),
            formula1: safe(() => (rule.Formula1 === undefined ? null : String(rule.Formula1)), null),
            interiorColor: safe(() => {
              const c = rule.Interior && rule.Interior.Color !== undefined ? Number(rule.Interior.Color) : null;
              return c === null || !Number.isFinite(c) ? null : excelColorToHex(c);
            }, null)
          });
        } catch (e) { /* 单条规则读失败不影响整体 */ }
      }
      return { count, scanned: Math.min(count, 50), rules };
    }, null);

    return state;
  }

  // 7. 切片读取数据
  function readRangeData(app, params) {
    const { sheetName, address, workbookName, includeFormulas = true, includeNumberFormats = false } = params;
    if (!address) throw new Error("缺少必要参数: address (例如 'A1:C10')");

    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);

    if (range.Rows.Count * range.Columns.Count > 10000) throw new Error("单次读取上限为 10000 个单元格，请分片读取");
    const values = range.Value2;
    let formulas = null;
    let numberFormats = null;

    if (includeFormulas) {
      formulas = range.Formula;
    }
    if (includeNumberFormats) {
      numberFormats = [];
      for (let r = 1; r <= range.Rows.Count; r++) {
        const row = [];
        for (let c = 1; c <= range.Columns.Count; c++) row.push(safeRead(() => range.Cells.Item(r, c).NumberFormat, null));
        numberFormats.push(row);
      }
    }

    return {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      rowCount: range.Rows.Count,
      columnCount: range.Columns.Count,
      values: normalize2DArray(values, range.Rows.Count, range.Columns.Count),
      formulas: formulas !== null && formulas !== undefined ? normalize2DArray(formulas, range.Rows.Count, range.Columns.Count) : null,
      numberFormats: numberFormats ? normalize2DArray(numberFormats, range.Rows.Count, range.Columns.Count) : null
    };
  }

  function safeRead(reader, fallback) {
    try {
      const value = reader();
      return value === undefined ? fallback : value;
    } catch (e) {
      return fallback;
    }
  }

  function alignmentName(value, vertical) {
    const map = vertical
      ? { "-4160": "top", "-4108": "center", "-4107": "bottom" }
      : { "-4131": "left", "-4108": "center", "-4152": "right", "1": "general" };
    return map[String(value)] || value;
  }

  function readBorders(range) {
    const borderIds = { left: 7, top: 8, bottom: 9, right: 10, insideHorizontal: 11, insideVertical: 12 };
    const result = {};
    Object.keys(borderIds).forEach((name) => {
      const border = safeRead(() => range.Borders.Item(borderIds[name]), null);
      if (!border) return;
      result[name] = {
        lineStyle: safeRead(() => border.LineStyle, null),
        weight: safeRead(() => border.Weight, null),
        color: excelColorToHex(safeRead(() => border.Color, null))
      };
    });
    return result;
  }

  function readStyleFields(range, include) {
    const result = {};
    const wants = (field) => include.indexOf(field) >= 0;
    if (wants("fontName")) result.fontName = safeRead(() => range.Font.Name, null);
    if (wants("fontSize")) result.fontSize = safeRead(() => range.Font.Size, null);
    if (wants("bold")) result.bold = safeRead(() => range.Font.Bold, null);
    if (wants("fontColor")) result.fontColor = excelColorToHex(safeRead(() => range.Font.Color, null));
    if (wants("backgroundColor")) result.backgroundColor = excelColorToHex(safeRead(() => range.Interior.Color, null));
    if (wants("numberFormat")) result.numberFormat = safeRead(() => range.NumberFormat, null);
    if (wants("horizontalAlignment")) result.horizontalAlignment = alignmentName(safeRead(() => range.HorizontalAlignment, null), false);
    if (wants("verticalAlignment")) result.verticalAlignment = alignmentName(safeRead(() => range.VerticalAlignment, null), true);
    if (wants("wrapText")) result.wrapText = safeRead(() => range.WrapText, null);
    if (wants("rowHeight")) result.rowHeight = safeRead(() => range.RowHeight, null);
    if (wants("columnWidth")) result.columnWidth = safeRead(() => range.ColumnWidth, null);
    if (wants("validation")) {
      // ISS-94：数据有效性原来只写不读，AI 设完下拉/范围校验后无法自检。
      result.validation = safeRead(() => {
        const v = range.Validation;
        const type = Number(v.Type);
        return {
          type,
          typeName: ({ 1: "xlValidateWholeNumber", 2: "xlValidateDecimal", 3: "xlValidateList", 4: "xlValidateDate", 5: "xlValidateTime", 6: "xlValidateTextLength", 7: "xlValidateCustom" })[type] || null,
          operator: Number(v.Operator),
          formula1: v.Formula1 === undefined ? null : String(v.Formula1),
          formula2: v.Formula2 === undefined ? null : String(v.Formula2),
          ignoreBlank: Boolean(v.IgnoreBlank),
          inCellDropdown: Boolean(v.InCellDropdown),
          prompt: v.InputMessage === undefined ? null : String(v.InputMessage),
          errorMessage: v.ErrorMessage === undefined ? null : String(v.ErrorMessage)
        };
      }, null);
    }

    if (wants("merged")) result.merged = safeRead(() => range.MergeCells, null);
    if (wants("mergeArea")) {
      const firstCell = safeRead(() => range.Cells.Item(1, 1), null);
      const isMerged = firstCell ? safeRead(() => !!firstCell.MergeCells, false) : false;
      result.mergeArea = isMerged ? safeRead(() => firstCell.MergeArea.Address(), null) : null;
    }
    if (wants("borders")) result.borders = readBorders(range);
    return result;
  }

  function getRangeStyles(app, params) {
    const config = params || {};
    const { sheetName, workbookName, address, mode = "summary" } = config;
    if (!address) throw new Error("缺少必要参数: address (例如 'A1:C10')");
    const allowed = ["fontName", "fontSize", "bold", "fontColor", "backgroundColor", "numberFormat", "horizontalAlignment", "verticalAlignment", "wrapText", "rowHeight", "columnWidth", "merged", "mergeArea", "borders", "validation"];
    const defaults = ["fontName", "fontSize", "bold", "fontColor", "backgroundColor", "numberFormat", "horizontalAlignment", "verticalAlignment", "wrapText", "rowHeight", "columnWidth", "merged", "mergeArea"];
    const include = Array.isArray(config.include) ? config.include.filter((field) => allowed.indexOf(field) >= 0) : defaults;
    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);
    const base = {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      rowCount: range.Rows.Count,
      columnCount: range.Columns.Count,
      mode
    };

    if (mode !== "cells") {
      const styles = readStyleFields(range, include);
      const mixedOrUnavailableFields = Object.keys(styles).filter((field) => {
        if (field === "mergeArea" && styles.merged === false) return false;
        return styles[field] === null;
      });
      return Object.assign(base, { styles, mixedOrUnavailableFields });
    }

    const maxCells = Math.max(1, Math.min(500, Math.trunc(Number(config.maxCells) || 100)));
    const cells = [];
    const totalCells = range.Rows.Count * range.Columns.Count;
    for (let r = 1; r <= range.Rows.Count && cells.length < maxCells; r++) {
      for (let c = 1; c <= range.Columns.Count && cells.length < maxCells; c++) {
        const cell = range.Cells.Item(r, c);
        cells.push(Object.assign({ address: cell.Address() }, readStyleFields(cell, include)));
      }
    }
    return Object.assign(base, { totalCells, returnedCells: cells.length, truncated: totalCells > cells.length, cells });
  }

  // 8. 单元格搜索
  function searchCells(app, params) {
    const { sheetName, query, maxResults = 50, workbookName, address } = params;
    if (!query) throw new Error("缺少搜索关键字: query");

    const sheet = getWorksheet(app, sheetName, workbookName);
    // 可选 address：原实现只能全表检索，报错却让调用方"缩小检索范围"（问题台账 ISS-29）。
    const usedRange = address ? sheet.Range(address) : sheet.UsedRange;
    if (!usedRange) return { matches: [] };

    if (usedRange.Rows.Count * usedRange.Columns.Count > 100000) {
      throw new Error("检索范围超过 100000 单元格，请用 address 指定更小的区域");
    }
    const matches = [];
    const values = normalize2DArray(usedRange.Value2, usedRange.Rows.Count, usedRange.Columns.Count);
    // 说明承诺可搜"文本或公式"，但原实现只读 Value2 —— 按公式搜必然 0 命中且返回 success
    // （问题台账 ISS-29，比报错更危险）。这里同时检索公式串。
    const formulas = normalize2DArray(usedRange.Formula, usedRange.Rows.Count, usedRange.Columns.Count);
    const startRow = usedRange.Row;
    const startCol = usedRange.Column;

    const lowerQuery = String(query).toLowerCase();

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const val = values[r][c];
        const formula = formulas[r] ? formulas[r][c] : null;
        const hitValue = val !== null && val !== undefined && String(val).toLowerCase().includes(lowerQuery);
        const hitFormula = formula !== null && formula !== undefined && String(formula).toLowerCase().includes(lowerQuery);
        if (hitValue || hitFormula) {
          const cellRow = startRow + r;
          const cellCol = startCol + c;
          matches.push({
            address: sheet.Cells.Item(cellRow, cellCol).Address(),
            row: cellRow,
            column: cellCol,
            value: val,
            formula: formula === null || formula === undefined ? "" : String(formula),
            matchedIn: hitValue ? "value" : "formula"
          });
          if (matches.length >= maxResults) break;
        }
      }
      if (matches.length >= maxResults) break;
    }

    return {
      sheetName: sheet.Name,
      query: query,
      searchedRange: usedRange.Address(),
      totalFound: matches.length,
      matches: matches
    };
  }

  // 9. 批量修改数据 (带快照)
  function patchCells(app, params) {
    const { sheetName, address, values, formulas, workbookName } = params;
    if (!address) throw new Error("缺少必要参数: address (例如 'B2:C5')");
    if (!values && !formulas) throw new Error("必须提供 values 或 formulas 进行更新");

    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);
    const rowCount = range.Rows.Count;
    const colCount = range.Columns.Count;
    if (rowCount * colCount > 10000) throw new Error("单次写入上限为 10000 个单元格，请分批处理");
    [values, formulas].forEach(matrix => {
      if (matrix !== undefined && matrix !== null && (!Array.isArray(matrix) || matrix.length !== rowCount || matrix.some(row => !Array.isArray(row) || row.length !== colCount))) throw new Error("写入矩阵必须与目标区域尺寸一致");
    });

    const oldValues = normalize2DArray(range.Value2, rowCount, colCount);
    const oldFormulas = normalize2DArray(range.Formula, rowCount, colCount);

    const beforeSnapshot = {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      rowCount: rowCount,
      columnCount: colCount,
      values: oldValues,
      formulas: oldFormulas
    };

    try {
      if (values !== undefined && values !== null) range.Value2 = values;
      if (formulas !== undefined && formulas !== null) {
        // 空项（''/null/undefined）表示"该单元格只写值、不改公式"，必须**跳过**而不是写入：
        // 给 range.Formula 赋空字符串等于清空单元格，会把上一步刚写入的值一起抹掉，
        // 而且仍返回 success + modifiedCount，属于静默数据丢失（见问题台账 ISS-01 / DP7）。
        // 只有在整张矩阵都没有空项时才用整批赋值，保住批量性能。
        const hasBlank = formulas.some(row => Array.isArray(row) && row.some(cell => cell === '' || cell === null || cell === undefined));
        if (hasBlank) {
          for (let r = 0; r < rowCount; r++) {
            const row = formulas[r];
            if (!Array.isArray(row)) continue;
            for (let c = 0; c < colCount; c++) {
              const cellFormula = row[c];
              if (cellFormula === '' || cellFormula === null || cellFormula === undefined) continue;
              range.Cells.Item(r + 1, c + 1).Formula = cellFormula;
            }
          }
        } else {
          range.Formula = formulas;
        }
      }
    } catch (writeError) {
      try { range.Formula = oldFormulas; }
      catch (restoreError) { throw new Error("写入部分失败且恢复失败，请检查目标区域：" + writeError.message); }
      throw new Error("写入失败，已恢复原值与公式：" + writeError.message);
    }

    const newValues = normalize2DArray(range.Value2, rowCount, colCount);
    const newFormulas = normalize2DArray(range.Formula, rowCount, colCount);

    const diff = [];
    for (let r = 0; r < rowCount; r++) {
      for (let c = 0; c < colCount; c++) {
        const oVal = oldValues[r] ? oldValues[r][c] : null;
        const nVal = newValues[r] ? newValues[r][c] : null;
        const oForm = oldFormulas[r] ? oldFormulas[r][c] : null;
        const nForm = newFormulas[r] ? newFormulas[r][c] : null;

        if (oVal !== nVal || oForm !== nForm) {
          diff.push({
            cell: range.Cells.Item(r + 1, c + 1).Address(),
            oldValue: oVal,
            newValue: nVal,
            oldFormula: oForm,
            newFormula: nForm
          });
        }
      }
    }

    const afterSnapshot = {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      rowCount: rowCount,
      columnCount: colCount,
      values: newValues,
      formulas: newFormulas
    };

    return {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      modifiedCount: diff.length,
      diff: diff,
      beforeSnapshot: beforeSnapshot,
      afterSnapshot: afterSnapshot
    };
  }

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
      const ICON_SET_CODES = {
        "3_arrows": 1, "3_arrows_gray": 2, "3_flags": 3,
        "3_traffic_lights": 4, "3_traffic_lights_rimmed": 5,
        "3_signs": 6, "3_symbols": 7, "3_symbols_circled": 8,
        "4_arrows": 9, "4_arrows_gray": 10, "4_red_to_black": 11,
        "4_ratings": 12, "4_traffic_lights": 13,
        "5_arrows": 14, "5_arrows_gray": 15, "5_quarters": 16,
        "5_ratings": 17, "5_boxes": 18
      };
      const iconSetName = iconSet || "3_traffic_lights";
      const code = ICON_SET_CODES[iconSetName];
      if (code === undefined) {
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

  // 10.2 冻结窗格吸顶
  function freezePanes(app, params) {
    const { sheetName, workbookName, freezeRowIndex, freezeColumnIndex, unfreeze, action = "apply" } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}

    const win = app.ActiveWindow;
    if (!win) throw new Error("无法获取 WPS 活动窗口句柄");

    // CAP-31 冻结窗格读回：只写不读的话，AI 无法确认"到底冻了几行"，
    // 也就没法在多次操作后知道自己当前处于什么状态。
    if (action === "read") {
      const frozen = (() => { try { return Boolean(win.FreezePanes); } catch (e) { return null; } })();
      const splitRow = (() => { try { return Number(win.SplitRow) || 0; } catch (e) { return null; } })();
      const splitColumn = (() => { try { return Number(win.SplitColumn) || 0; } catch (e) { return null; } })();
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        frozen,
        // 对外给"冻结到第几行/列"（1 基，更接近用户说法），同时保留原始 SplitRow/SplitColumn
        freezeRowIndex: splitRow === null ? null : splitRow + 1,
        freezeColumnIndex: splitColumn === null ? null : splitColumn + 1,
        splitRow, splitColumn,
        warnings: [],
        message: frozen
          ? `工作表 [${sheet.Name}] 已冻结：第 ${splitRow} 行之上的 ${splitRow} 行、第 ${splitColumn} 列之左的 ${splitColumn} 列保持不动`
          : `工作表 [${sheet.Name}] 当前未冻结窗格`
      };
    }

    if (unfreeze) {
      win.FreezePanes = false;
      win.SplitRow = 0;
      win.SplitColumn = 0;
      return { success: true, message: "已成功解除当前工作表的窗格冻结" };
    }

    win.FreezePanes = false;
    // 越界值校验：原先 `freezeRowIndex:0` 会走 else 分支导致 SplitRow 残留旧值 → 随机冻结（L-2）
    if (freezeRowIndex !== undefined && freezeRowIndex !== null && !(Number(freezeRowIndex) >= 1)) {
      throw new Error(`freezeRowIndex 必须是 >= 1 的整数（收到 ${freezeRowIndex}）；不冻传 unfreeze:true`);
    }
    if (freezeRowIndex && freezeRowIndex > 1) {
      win.SplitRow = freezeRowIndex - 1;
    }
    if (freezeColumnIndex && freezeColumnIndex > 1) {
      win.SplitColumn = freezeColumnIndex - 1;
    }
    win.FreezePanes = true;

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      freezeRowIndex: freezeRowIndex || 1,
      freezeColumnIndex: freezeColumnIndex || 1,
      message: `已成功锁定第 ${freezeRowIndex ? freezeRowIndex - 1 : 0} 行之上的表头吸顶显示`
    };
  }

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

  // 10.27 表格矢量绘图（CAP-07）
  //
  // 依据真机探测（WPS 12.1.28496）：宿主 `Sheet.Shapes` 全套可用——
  //   AddShape / AddLine / AddTextbox / **AddTextEffect（艺术字）** / AddPicture / BuildFreeform，
  //   分组走 `Shapes.Range([名...]).Group()`（**WPS 表格能分组，Office.js 侧这条未打通**）。
  // 这是 WPS 相对 Microsoft Excel 的**能力优势**：艺术字与自由曲线在 Office.js 侧没有对应 API。
  // 此前我们一个表格绘图工具都没有，AI 只能用 wps_execute_script 手写。

  /** 几何形状枚举（msoAutoShapeType 常用值）。 */
  const AUTO_SHAPE_TYPES = {
    rectangle: 1, rounded_rectangle: 5, oval: 9, ellipse: 9, diamond: 4, triangle: 7,
    right_triangle: 8, pentagon: 51, hexagon: 10, arrow_right: 33, arrow_left: 34,
    arrow_up: 35, arrow_down: 36, chevron: 52, cross: 11, star_5: 12, star_6: 13,
    callout_rounded: 106, cloud: 179, heart: 21, lightning: 22, sun: 24, moon: 23,
    can: 13, cube: 17, donut: 18, flow_chart_process: 61, flow_chart_decision: 63,
    flow_chart_terminator: 68, flow_chart_data: 62, line_horizontal: 130, text_box: 17
  };

  /** 形状文字的默认字体。中文报表里微软雅黑比宿主默认的宋体更清晰、字重更统一。 */
  const DEFAULT_SHAPE_FONT = "微软雅黑";

  function resolveShapeType(raw) {
    if (raw === undefined || raw === null || raw === "") return AUTO_SHAPE_TYPES.rectangle;
    if (typeof raw === "number") return raw;
    const key = String(raw).toLowerCase().trim();
    if (AUTO_SHAPE_TYPES[key] !== undefined) return AUTO_SHAPE_TYPES[key];
    if (/^\d+$/.test(key)) return parseInt(key, 10);
    throw new Error(`不认识的形状类型 "${raw}"。可用：${Object.keys(AUTO_SHAPE_TYPES).join(" / ")}`);
  }

  /**
   * 判断某个十六进制底色偏亮还是偏暗，用于给形状文字选黑字还是白字。
   *
   * 为什么需要：WPS 给**自选图形默认白字**。如果调用方填了白底又不指定字色，
   * 就会得到"白字白底"——文字完全看不见，而且没有任何报错（真机踩到过：
   * 流程框全部隐形）。这里按底色亮度自动选，调用方显式给 fontColor 时以调用方为准。
   */
  function isLightFill(hex) {
    const rgb = hexToExcelColor(hex);
    if (rgb === null || rgb === undefined) return null;   // 无底色 → 交给宿主默认
    const r = rgb & 0xff, g = (rgb >> 8) & 0xff, b = (rgb >> 16) & 0xff;
    return (0.299 * r + 0.587 * g + 0.114 * b) > 150;
  }

  /** 单元格地址取值：宿主方法是原生实现，**必须保留接收者**直接调用（摘出来会丢 this）。 */
  function shapeAddressOf(cell) {
    try { return String(cell.Address()); } catch (e) {
      try { return String(cell.Address); } catch (e2) { return ""; }
    }
  }

  /** 读回单个形状的完整状态。 */
  function readShape(shape, sheet) {
    const safe = (fn, fallback) => { try { const v = fn(); return v === undefined ? fallback : v; } catch (e) { return fallback; } };
    let text = null;
    try { text = String(shape.TextFrame.Characters().Text); } catch (e) { text = null; }
    return {
      name: safe(() => String(shape.Name), null),
      type: safe(() => Number(shape.Type), null),
      autoShapeType: safe(() => Number(shape.AutoShapeType), null),
      left: safe(() => Math.round(Number(shape.Left) * 100) / 100, null),
      top: safe(() => Math.round(Number(shape.Top) * 100) / 100, null),
      width: safe(() => Math.round(Number(shape.Width) * 100) / 100, null),
      height: safe(() => Math.round(Number(shape.Height) * 100) / 100, null),
      rotation: safe(() => Number(shape.Rotation), null),
      visible: safe(() => Number(shape.Visible) !== 0, null),
      fillColor: safe(() => { const v = Number(shape.Fill.ForeColor.RGB); return Number.isFinite(v) && v >= 0 ? excelColorToHex(v) : null; }, null),
      lineColor: safe(() => { const v = Number(shape.Line.ForeColor.RGB); return Number.isFinite(v) && v >= 0 ? excelColorToHex(v) : null; }, null),
      lineWeight: safe(() => Number(shape.Line.Weight), null),
      text,
      // 真实对齐（Excel 对象模型：HorizontalAlignment/VerticalAlignment）。
      // 放进 readShape 而不是只在 addShape 里，这样 list_shapes 也能核对——
      // 「框居中了但字没居中」这类静默失败必须能被读回发现。
      textAlign: safe(() => { const h = Number(shape.TextFrame.HorizontalAlignment); return h === -4108 ? "center" : h === -4152 ? "right" : h === -4131 ? "left" : null; }, null),
      textVAlign: safe(() => { const v = Number(shape.TextFrame.VerticalAlignment); return v === -4108 ? "middle" : v === -4107 ? "bottom" : v === -4160 ? "top" : null; }, null),
      topLeftCell: safe(() => shapeAddressOf(shape.TopLeftCell), null),
      isGroup: safe(() => Number(shape.Type) === 6, false),
      groupItemCount: safe(() => (Number(shape.Type) === 6 && shape.GroupItems ? Number(shape.GroupItems.Count) : null), null)
    };
  }

  /** 找出形状：优先按名字，其次按序号（1 基）。 */
  function findShape(sheet, params) {
    const shapes = sheet.Shapes;
    if (params.name || params.shapeName) {
      const name = String(params.name || params.shapeName);
      for (let i = 1; i <= shapes.Count; i++) {
        const sh = shapes.Item(i);
        try { if (String(sh.Name) === name) return sh; } catch (e) {}
      }
      const available = [];
      for (let i = 1; i <= Math.min(shapes.Count, 30); i++) {
        try { available.push(String(shapes.Item(i).Name)); } catch (e) {}
      }
      throw new Error(`找不到名为 "${name}" 的形状。现有形状（最多 30 个）：${available.join(" / ") || "无"}`);
    }
    if (Number.isFinite(Number(params.shapeIndex))) {
      const idx = Number(params.shapeIndex);
      if (idx < 1 || idx > shapes.Count) throw new Error(`形状序号 ${idx} 越界（当前共 ${shapes.Count} 个）`);
      return shapes.Item(idx);
    }
    throw new Error("需要提供 name（形状名）或 shapeIndex（序号，从 1 开始）");
  }

  /**
   * 画矢量形状：几何形状 / 直线 / 文本框 / **艺术字**。
   * 写完**读回真实几何与样式**，宿主没接受某个属性时写进 warnings，不做假成功。
   */
  function addShape(app, params) {
    const { sheetName, workbookName, kind = "geometric", shapeType, text, name,
            left = 0, top = 0, width = 160, height = 80, rotation,
            x1, y1, x2, y2, fillColor, lineColor, lineWeight,
            wordArtPreset, fontName, fontSize, bold, italic, fontColor,
            textAlign, textVAlign, marginLeft, marginRight, marginTop, marginBottom } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shapes = sheet.Shapes;
    const warnings = [];

    let shape = null;
    const k = String(kind).toLowerCase();
    if (k === "geometric" || k === "autoshape") {
      shape = shapes.AddShape(resolveShapeType(shapeType), Number(left), Number(top), Number(width), Number(height));
    } else if (k === "line" || k === "connector") {
      shape = shapes.AddLine(
        Number(x1 !== undefined ? x1 : left), Number(y1 !== undefined ? y1 : top),
        Number(x2 !== undefined ? x2 : Number(left) + Number(width)),
        Number(y2 !== undefined ? y2 : Number(top) + Number(height))
      );
    } else if (k === "textbox" || k === "text") {
      shape = shapes.AddTextbox(1, Number(left), Number(top), Number(width), Number(height));
      // 文本框默认**不画边框、不填底色**——否则页面上每段文字都套一个框。
      // 调用方显式传 fillColor/lineColor 时才打开对应可见性（见下方）。
      try { shape.Line.Visible = 0; } catch (e) {}
      try { shape.Fill.Visible = 0; } catch (e) {}
      // 关掉自动缩放：宿主默认会把文本框收缩到文字宽度，导致"按给定宽度算的居中"全部偏掉
      // （真机踩到：柱顶数据标签偏离柱心 1~2pt）。宽度按调用方给的固定，版面才可预测。
      try { shape.TextFrame.AutoSize = 0; } catch (e) {}
      try { shape.TextFrame.WordWrap = 0; } catch (e) {}
    } else if (k === "wordart" || k === "texteffect") {
      // 艺术字：WPS 相对 Office.js 的独有能力
      const preset = Number.isFinite(Number(wordArtPreset)) ? Number(wordArtPreset) : 0;
      shape = shapes.AddTextEffect(
        preset, String(text === undefined ? "" : text), String(fontName || "宋体"),
        Number(fontSize || 36), bold === false ? 0 : -1, italic ? -1 : 0, Number(left), Number(top)
      );
    } else {
      throw new Error(`不支持的形状种类: ${kind}（可用 geometric / line / textbox / wordart）`);
    }
    if (!shape) throw new Error("宿主没有返回形状对象（创建失败）");

    if (name) { try { shape.Name = String(name); } catch (e) { warnings.push(`设置名字失败: ${e.message}`); } }
    if (k !== "line" && Number.isFinite(Number(rotation))) {
      try { shape.Rotation = Number(rotation); } catch (e) { warnings.push(`设置旋转失败: ${e.message}`); }
    }
    if (fillColor) {
      const rgb = hexToExcelColor(fillColor);
      if (rgb === null) warnings.push(`fillColor 无法解析: ${fillColor}`);
      else { try { shape.Fill.Visible = -1; shape.Fill.ForeColor.RGB = rgb; } catch (e) { warnings.push(`设置填充失败: ${e.message}`); } }
    }
    if (lineColor) {
      const rgb = hexToExcelColor(lineColor);
      if (rgb === null) warnings.push(`lineColor 无法解析: ${lineColor}`);
      else { try { shape.Line.Visible = -1; shape.Line.ForeColor.RGB = rgb; } catch (e) { warnings.push(`设置线条色失败: ${e.message}`); } }
    }
    if (Number.isFinite(Number(lineWeight))) {
      try { shape.Line.Weight = Number(lineWeight); } catch (e) { warnings.push(`设置线宽失败: ${e.message}`); }
    }
    // 文字：几何形状/文本框都能写字（艺术字的文字在创建时给）。
    // 字号/加粗/斜体/字体/字色对**所有形状类型**生效——原来只对艺术字生效，
    // 于是"画个文本框并给它 22pt 加粗"这类最常见需求被静默忽略。
    if (k !== "wordart" && k !== "texteffect" && text !== undefined && text !== null && String(text) !== "") {
      try {
        const chars = shape.TextFrame.Characters();
        chars.Text = String(text);
        if (Number.isFinite(Number(fontSize))) chars.Font.Size = Number(fontSize);
        if (bold !== undefined) chars.Font.Bold = bold ? -1 : 0;
        if (italic !== undefined) chars.Font.Italic = italic ? -1 : 0;
        chars.Font.Name = String(fontName || DEFAULT_SHAPE_FONT);
        if (fontColor) {
          const fc = hexToExcelColor(fontColor);
          if (fc !== null) chars.Font.Color = fc;
        } else {
          // 未指定字色：按填充色亮度选黑/白，避免"白字白底"隐形
          const light = isLightFill(fillColor);
          if (light === true) chars.Font.Color = hexToExcelColor("#333333");
          else if (light === false) chars.Font.Color = hexToExcelColor("#FFFFFF");
        }
        // 对齐：文本框默认左对齐/顶对齐，自选图形默认居中/垂直居中。
        //
        // ⚠️ 这里是 **Excel 的对象模型**，不是 PowerPoint 的：
        //   `TextFrame.TextRange.ParagraphFormat.Alignment` 在 Excel 里**不存在**
        //   （TextRange=undefined、Characters().ParagraphFormat=undefined），
        //   正确属性是 `TextFrame.HorizontalAlignment` / `VerticalAlignment`。
        // 之前用错的那条路径被 try/catch 吞掉，于是"框居中了、框里的字还是居左"，
        // 既不报错也看不出来（使用者一眼就看出来了）——所以这里改成**失败即告警**，不再静默。
        const tf = shape.TextFrame;
        const H = { left: -4131, center: -4108, right: -4152 };   // xlHAlign*
        const V = { top: -4160, middle: -4108, bottom: -4107 };   // xlVAlign*
        const wantH = textAlign || ((k !== "textbox" && k !== "text") ? "center" : null);
        const wantV = textVAlign || ((k !== "textbox" && k !== "text") ? "middle" : null);
        if (wantH && H[wantH] !== undefined) {
          try { tf.HorizontalAlignment = H[wantH]; } catch (e) { warnings.push(`设置水平对齐失败: ${e.message}`); }
        }
        if (wantV && V[wantV] !== undefined) {
          try { tf.VerticalAlignment = V[wantV]; } catch (e) { warnings.push(`设置垂直对齐失败: ${e.message}`); }
        }
        for (const [key, val] of Object.entries({ MarginLeft: marginLeft, MarginRight: marginRight, MarginTop: marginTop, MarginBottom: marginBottom })) {
          if (Number.isFinite(Number(val))) { try { tf[key] = Number(val); } catch (e) { warnings.push(`设置 ${key} 失败: ${e.message}`); } }
        }
      } catch (e) { warnings.push(`写入文字失败: ${e.message}`); }
    }

    const actual = readShape(shape, sheet);

    // 核对请求与读回：不一致就如实告警（不抛错，因为部分属性宿主可能合法地做了归一）
    const requested = { kind: k, left: Number(left), top: Number(top), width: Number(width), height: Number(height) };
    if (k === "geometric" || k === "autoshape") {
      const tol = 1.5;
      for (const key of ["left", "top", "width", "height"]) {
        if (actual[key] !== null && Math.abs(Number(actual[key]) - requested[key]) > tol) {
          warnings.push(`${key} 请求 ${requested[key]} 读回 ${actual[key]}`);
        }
      }
    }
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      kind: k,
      shape: actual,
      requested,
      shapeCount: shapes.Count,
      warnings,
      message: `已在 [${sheet.Name}] 创建 ${k} 形状「${actual.name || name || ""}」（当前共 ${shapes.Count} 个形状）`
    };
  }

  /** 读回工作表上的全部形状（绘图能力的验收入口）。 */
  function listShapes(app, params) {
    const { sheetName, workbookName, detail = true, filterName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shapes = sheet.Shapes;
    const items = [];
    for (let i = 1; i <= shapes.Count; i++) {
      const sh = shapes.Item(i);
      if (filterName) {
        let nm = null;
        try { nm = String(sh.Name); } catch (e) {}
        if (nm !== String(filterName)) continue;
      }
      items.push(detail ? readShape(sh, sheet) : { name: (() => { try { return String(sh.Name); } catch (e) { return null; } })() });
    }
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      count: items.length,
      shapes: items,
      message: `工作表 [${sheet.Name}] 上共有 ${items.length} 个形状`
    };
  }

  /** 修改形状：位置/尺寸/旋转/填充/线条/文字/可见性，或删除。 */
  function updateShape(app, params) {
    const { sheetName, workbookName, action, left, top, width, height, rotation,
            fillColor, lineColor, lineWeight, text, visible, newName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shape = findShape(sheet, params);
    const warnings = [];

    const act = String(action || "update").toLowerCase();
    if (act === "delete") {
      const nm = (() => { try { return String(shape.Name); } catch (e) { return null; } })();
      shape.Delete();
      return {
        success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name,
        action: "delete", deletedName: nm, shapeCountAfter: sheet.Shapes.Count,
        message: `已删除形状「${nm}」（剩余 ${sheet.Shapes.Count} 个）`
      };
    }

    if (Number.isFinite(Number(left))) shape.Left = Number(left);
    if (Number.isFinite(Number(top))) shape.Top = Number(top);
    if (Number.isFinite(Number(width))) shape.Width = Number(width);
    if (Number.isFinite(Number(height))) shape.Height = Number(height);
    if (Number.isFinite(Number(rotation))) shape.Rotation = Number(rotation);
    if (newName) { try { shape.Name = String(newName); } catch (e) { warnings.push(`改名失败: ${e.message}`); } }
    if (visible !== undefined) { try { shape.Visible = visible ? -1 : 0; } catch (e) { warnings.push(`设置可见性失败: ${e.message}`); } }
    if (fillColor) {
      const rgb = hexToExcelColor(fillColor);
      if (rgb === null) warnings.push(`fillColor 无法解析: ${fillColor}`);
      else { try { shape.Fill.Visible = -1; shape.Fill.ForeColor.RGB = rgb; } catch (e) { warnings.push(`设置填充失败: ${e.message}`); } }
    }
    if (lineColor) {
      const rgb = hexToExcelColor(lineColor);
      if (rgb === null) warnings.push(`lineColor 无法解析: ${lineColor}`);
      else { try { shape.Line.Visible = -1; shape.Line.ForeColor.RGB = rgb; } catch (e) { warnings.push(`设置线条色失败: ${e.message}`); } }
    }
    if (Number.isFinite(Number(lineWeight))) {
      try { shape.Line.Weight = Number(lineWeight); } catch (e) { warnings.push(`设置线宽失败: ${e.message}`); }
    }
    if (text !== undefined && text !== null) {
      try { shape.TextFrame.Characters().Text = String(text); } catch (e) { warnings.push(`写入文字失败: ${e.message}`); }
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      action: "update",
      shape: readShape(shape, sheet),
      warnings,
      message: `已更新形状「${(() => { try { return String(shape.Name); } catch (e) { return ""; } })()}」`
    };
  }

  /** 分组：`Shapes.Range([名...]).Group()`（WPS 表格可用；Office.js 侧未打通）。 */
  function groupShapes(app, params) {
    const { sheetName, workbookName, names, groupName } = params || {};
    const list = Array.isArray(names) ? names.filter(Boolean).map(String) : [];
    if (list.length < 2) throw new Error("group_shapes 需要至少 2 个形状（names 传形状名数组）");
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shapes = sheet.Shapes;
    // 先确认都存在，错误信息才有用
    const existing = [];
    for (let i = 1; i <= shapes.Count; i++) { try { existing.push(String(shapes.Item(i).Name)); } catch (e) {} }
    const missing = list.filter(n => existing.indexOf(n) < 0);
    if (missing.length) throw new Error(`这些形状不存在: ${missing.join(", ")}。现有形状：${existing.join(" / ") || "无"}`);

    const group = shapes.Range(list).Group();
    if (groupName) { try { group.Name = String(groupName); } catch (e) {} }
    const members = [];
    try {
      const gi = group.GroupItems;
      for (let i = 1; i <= gi.Count; i++) { try { members.push(String(gi.Item(i).Name)); } catch (e) {} }
    } catch (e) {}
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      group: readShape(group, sheet),
      requestedMembers: list,
      memberCount: members.length,
      members,
      verified: members.length === list.length,
      warnings: members.length === list.length ? [] : [`请求组合 ${list.length} 个，读回组内 ${members.length} 个`],
      message: `已把 ${list.length} 个形状组合为「${(() => { try { return String(group.Name); } catch (e) { return groupName || ""; } })()}」`
    };
  }

  /** 解散分组，返回释放后的形状名单。 */
  function ungroupShapes(app, params) {
    const { sheetName, workbookName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shape = findShape(sheet, params);
    const type = (() => { try { return Number(shape.Type); } catch (e) { return null; } })();
    if (type !== 6) {
      throw new Error(`形状「${(() => { try { return String(shape.Name); } catch (e) { return ""; } })()}」不是组合（Type=${type}，组合应为 6）`);
    }
    const before = [];
    try { const gi = shape.GroupItems; for (let i = 1; i <= gi.Count; i++) { try { before.push(String(gi.Item(i).Name)); } catch (e) {} } } catch (e) {}
    shape.Ungroup();
    const after = [];
    const shapes = sheet.Shapes;
    for (let i = 1; i <= shapes.Count; i++) { try { after.push(String(shapes.Item(i).Name)); } catch (e) {} }
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      released: before,
      releasedStillPresent: before.filter(n => after.indexOf(n) >= 0),
      shapeCountAfter: shapes.Count,
      message: `已解散组合，释放 ${before.length} 个形状`
    };
  }

  /** 调整层级。VBA ZOrder 常量：0=置顶 1=置底 2=上移一层 3=下移一层。 */
  function setShapeZorder(app, params) {
    const { sheetName, workbookName, zOrder = "bringToFront" } = params || {};
    const Z = { bringtofront: 0, sendtoback: 1, bringforward: 2, sendbackward: 3 };
    const code = Z[String(zOrder).toLowerCase()];
    if (code === undefined) throw new Error(`不支持的 zOrder: ${zOrder}（可用 bringToFront / sendToBack / bringForward / sendBackward）`);
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shape = findShape(sheet, params);
    shape.ZOrder(code);
    // 读回全部形状以核对层级（WPS 上 ZOrderPosition 不可读，改用"绘制顺序"近似：
    // 按 ZOrder(0) 逐个置顶来还原顺序代价太高，这里只如实返回当前形状与形状总数）
    const names = [];
    const shapes = sheet.Shapes;
    for (let i = 1; i <= shapes.Count; i++) { try { names.push(String(shapes.Item(i).Name)); } catch (e) {} }
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      applied: zOrder,
      shape: readShape(shape, sheet),
      shapeNames: names,
      warnings: ["WPS 表格读不到 ZOrderPosition，无法直接回读层级；shapeNames 的顺序是 Shapes 集合顺序，可作为近似参考"],
      message: `已把形状「${(() => { try { return String(shape.Name); } catch (e) { return ""; } })()}」调整为 ${zOrder}`
    };
  }

  // 10.3 整行整列增删、隐藏与高度宽度控制
  function modifyRowsColumns(app, params) {
    const { sheetName, workbookName, targetType, action, index, count = 1, size } = params || {};
    if (!targetType || !action || index === undefined) {
      throw new Error("缺少必要参数: targetType ('row'|'column'), action ('insert'|'delete'|'hide'|'unhide'|'set_size'), index (起始行号或列号/列字母)");
    }
    const sheet = getWorksheet(app, sheetName, workbookName);

    // 解析列号（支持数字如 2 或字母如 "B"）
    function parseColIndex(val) {
      if (typeof val === "number") return val;
      const str = String(val).toUpperCase().trim();
      if (/^\d+$/.test(str)) return parseInt(str, 10);
      let num = 0;
      for (let i = 0; i < str.length; i++) {
        num = num * 26 + (str.charCodeAt(i) - 64);
      }
      return num > 0 ? num : 1;
    }

    if (targetType === "row") {
      const startRow = Number(index);
      const endRow = startRow + Number(count) - 1;
      const rowRange = sheet.Range(`${startRow}:${endRow}`);

      if (action === "insert") {
        rowRange.Insert(-4121); // xlDown
      } else if (action === "delete") {
        rowRange.Delete(-4162); // xlUp
      } else if (action === "hide" || action === "unhide") {
        const want = action === "hide";
        // 本机 WPS 上给 `Range("5:6").Hidden` 赋值是**静默 no-op**（列走 Columns.Item 就正常），
        // 因此改为逐行写 Rows.Item(n).Hidden，并在写完后**读回校验**，未生效即报错（问题台账 ISS-60）。
        for (let r = startRow; r <= endRow; r++) {
          sheet.Rows.Item(r).Hidden = want;
        }
        const readBack = [];
        for (let r = startRow; r <= endRow; r++) {
          readBack.push(sheet.Rows.Item(r).Hidden === true);
        }
        if (readBack.some(applied => applied !== want)) {
          throw new Error(
            `${want ? "隐藏" : "取消隐藏"}第 ${startRow}-${endRow} 行未生效（逐行读回 ${JSON.stringify(readBack)}）。` +
            `已改用 Rows.Item(n).Hidden 逐行写入，若仍失败请检查该行是否被工作表保护。`
          );
        }
      } else if (action === "set_size") {
        if (size === undefined) throw new Error("设置行高时必须传入 size 参数 (单位: 磅值)");
        rowRange.RowHeight = Number(size);
      } else {
        throw new Error(`未知的行操作动作: ${action}`);
      }
    } else if (targetType === "column") {
      const startCol = parseColIndex(index);
      for (let i = 0; i < count; i++) {
        const col = sheet.Columns.Item(startCol + i);
        if (action === "insert") {
          col.Insert(-4161); // xlToRight
        } else if (action === "delete") {
          col.Delete(-4159); // xlToLeft
        } else if (action === "hide") {
          col.Hidden = true;
        } else if (action === "unhide") {
          col.Hidden = false;
        } else if (action === "set_size") {
          if (size === undefined) throw new Error("设置列宽时必须传入 size 参数 (字符宽度)");
          col.ColumnWidth = Number(size);
        } else {
          throw new Error(`未知的列操作动作: ${action}`);
        }
      }
    } else {
      throw new Error(`targetType 必须是 'row' 或 'column'`);
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      targetType,
      action,
      index,
      count,
      size,
      message: `已成功对工作表 [${sheet.Name}] 的第 ${index} ${targetType === "row" ? "行" : "列"}（共 ${count} 项）执行 ${action} 操作`
    };
  }

  // 10.4 单元格批注与审阅备注管理
  function manageCellComments(app, params) {
    const { sheetName, workbookName, address, action = "read", text, author } = params || {};
    if (!address && action !== "clear_all") {
      throw new Error("缺少必要参数: address (如 'C5' 或 'D10:D20')");
    }
    const sheet = getWorksheet(app, sheetName, workbookName);
    const targetRange = address ? sheet.Range(address) : sheet.UsedRange;

    // 单元格地址取值。**必须保留接收者直接调用**：宿主方法是原生实现，
    // 摘出来再调（`const f = cell.Address; f()`）会丢 this，抛
    // "Address called on null or undefined"，被 catch 吞掉后表现为地址为空
    // （真机实测：消息变成"已在单元格  添加批注"，ISS-66 的根因就在这里）。
    const addressOf = (cell) => {
      try {
        return String(cell.Address());
      } catch (e) {
        try {
          // 少数宿主把 Address 暴露成属性而非方法
          return String(cell.Address);
        } catch (e2) {
          return "";
        }
      }
    };

    if (action === "add") {
      if (!text) throw new Error("添加批注时必须提供 text 参数");
      const cell = targetRange.Cells.Item(1, 1);
      // 地址必须在 AddComment **之前**取：真机实测 AddComment 之后该单元格的 Address 取不到，
      // 会让成功消息变成"已在单元格  添加批注"（地址为空）。
      const cellAddressText = addressOf(cell); // 先取地址，AddComment 后取不到
      try {
        if (cell.Comment) cell.Comment.Delete();
      } catch (e) {}
      // 不再把作者拼进正文（ISS-43）：先尝试设置真实 Author，设不上再退回"正文前缀"，
      // 并在返回体里**如实区分**"请求的作者"与"读回的作者"（原来直接回显入参，冒充读回，ISS-61）。
      const comment = cell.AddComment(text);
      comment.Visible = false;
      let authorApplied = false;
      if (author) {
        try {
          comment.Author = author;
          authorApplied = true;
        } catch (e) {
          authorApplied = false;
        }
      }
      let authorOnHost = null;
      try { authorOnHost = comment.Author || null; } catch (e) {}
      if (author && !authorApplied && authorOnHost !== author) {
        try {
          comment.Text(text ? `${author}:\n${text}` : text);
        } catch (e) {}
      }
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        address: cellAddressText,
        action: "add",
        commentText: text,
        authorRequested: author || null,
        authorOnHost,
        authorApplied: authorOnHost === author,
        message: `已在单元格 ${cellAddressText} 添加批注${author && authorOnHost !== author ? `（宿主未接受自定义作者，已把作者写进正文；当前宿主作者为 ${authorOnHost ?? "未知"}）` : ""}`
      };
    } else if (action === "read") {
      const comments = [];
      const rowCount = targetRange.Rows.Count;
      const colCount = targetRange.Columns.Count;
      for (let r = 1; r <= rowCount; r++) {
        for (let c = 1; c <= colCount; c++) {
          const cell = targetRange.Cells.Item(r, c);
          try {
            if (cell.Comment) {
              comments.push({
                address: addressOf(cell),
                text: cell.Comment.Text(),
                author: cell.Comment.Author || undefined
              });
            }
          } catch (e) {}
        }
      }
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        address: targetRange.Address,
        comments,
        totalComments: comments.length
      };
    } else if (action === "delete") {
      const cell = targetRange.Cells.Item(1, 1);
      try {
        if (cell.Comment) cell.Comment.Delete();
      } catch (e) {}
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        address: cell.Address,
        action: "delete",
        message: `已成功删除单元格 ${cell.Address} 上的批注`
      };
    } else if (action === "clear_all") {
      try {
        sheet.Cells.ClearComments();
      } catch (e) {}
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        action: "clear_all",
        message: `已清空工作表 [${sheet.Name}] 中的所有批注`
      };
    }
    throw new Error(`未知的批注操作 action: ${action} (支持 'add' | 'read' | 'delete' | 'clear_all')`);
  }

  // 10.5 全局查找与定位替换
  function findAndReplace(app, params) {
    const { sheetName, workbookName, searchQuery, replaceText, matchCase = false, matchEntireCell = false, searchRange, maxResults = 50 } = params || {};
    if (searchQuery === undefined || searchQuery === null) {
      throw new Error("缺少必要参数: searchQuery (要查找的文本或数值)");
    }
    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = searchRange ? sheet.Range(searchRange) : sheet.UsedRange;
    const queryStr = String(searchQuery);
    const queryLower = queryStr.toLowerCase();

    const rowCount = range.Rows.Count;
    const colCount = range.Columns.Count;
    const values = range.Value2 || [];
    const formulas = range.Formula || [];

    const matches = [];
    let replacedCount = 0;

    for (let r = 0; r < rowCount; r++) {
      const rowVal = Array.isArray(values) ? values[r] : [values];
      const rowForm = Array.isArray(formulas) ? formulas[r] : [formulas];
      for (let c = 0; c < colCount; c++) {
        const val = rowVal ? rowVal[c] : null;
        const form = rowForm ? rowForm[c] : null;
        if (val === null || val === undefined) continue;

        const cellStr = String(val);
        const formStr = form === null || form === undefined ? "" : String(form);
        const testMatch = (text) => matchEntireCell
          ? (matchCase ? text === queryStr : text.toLowerCase() === queryLower)
          : (matchCase ? text.includes(queryStr) : text.toLowerCase().includes(queryLower));

        const hitValue = testMatch(cellStr);
        // 原实现读了 formulas 却只拿 Value2 匹配 —— 按公式片段搜必然 0 命中且返回 success
        // （问题台账 ISS-62）。这里把真正的公式串（以 = 开头）也纳入检索。
        const hitFormula = !hitValue && formStr.startsWith("=") && testMatch(formStr);
        const matched = hitValue || hitFormula;

        if (matched) {
          const cell = range.Cells.Item(r + 1, c + 1);
          const addr = cell.Address;
          let isReplaced = false;

          if (replaceText !== undefined) {
            const source = hitFormula ? formStr : cellStr;
            const newVal = matchEntireCell
              ? replaceText
              : source.replace(new RegExp(queryStr.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'g' : 'gi'), replaceText);
            // 命中在公式里就写回公式位，否则写值位
            if (hitFormula) cell.Formula = newVal; else cell.Value2 = newVal;
            isReplaced = true;
            replacedCount++;
          }

          matches.push({
            address: addr,
            row: r + 1,
            col: c + 1,
            value: val,
            formula: formStr.startsWith("=") ? formStr : null,
            matchedIn: hitFormula ? "formula" : "value",
            replaced: isReplaced
          });

          if (matches.length >= maxResults) break;
        }
      }
      if (matches.length >= maxResults) break;
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      searchQuery,
      replaceText,
      totalFound: matches.length,
      replacedCount,
      results: matches
    };
  }

  // 10.6 工作表复制与模板克隆
  function duplicateSheet(app, params) {
    const { sourceSheetName, newSheetName, position = "after", workbookName } = params || {};
    if (!sourceSheetName || !newSheetName) {
      throw new Error("缺少必要参数: sourceSheetName (源工作表) 和 newSheetName (新工作表名称)");
    }
    const wb = getWorkbook(app, workbookName);
    const srcSheet = getWorksheet(app, sourceSheetName, workbookName);

    try {
      const exist = wb.Worksheets.Item(newSheetName);
      if (exist) throw new Error(`工作表 [${newSheetName}] 已存在，请指定不同的新工作表名称`);
    } catch (e) {
      if (e.message && e.message.includes("已存在")) throw e;
    }

    if (position === "before") {
      srcSheet.Copy(srcSheet);
    } else {
      srcSheet.Copy(null, srcSheet);
    }

    const newSheet = wb.ActiveSheet;
    newSheet.Name = newSheetName;

    // M-3：`position` 原先被**完全忽略**（只落在源表之后），但工具照样报成功。
    // 支持 end / before / after（before|after 需配 targetSheetName），并按请求重排。
    const pos = String(position || "after").toLowerCase();
    let positionApplied = "after";
    const warnings = [];
    try {
      if (pos === "end") {
        newSheet.Move(undefined, wb.Worksheets.Item(wb.Worksheets.Count));
        positionApplied = "end";
      } else if (pos === "before" || pos === "after") {
        const targetName = (params && (params.targetSheetName || params.relativeTo)) || sourceSheetName;
        const anchor = wb.Worksheets.Item(String(targetName));
        if (pos === "before") newSheet.Move(anchor);
        else newSheet.Move(undefined, anchor);
        positionApplied = `${pos}:${targetName}`;
      } else {
        warnings.push(`不认识的 position: ${position}（可用 end / before / after），已按 after 处理`);
      }
    } catch (e) {
      warnings.push(`按 position=${position} 重排失败: ${e.message}（新表已创建，位置可能不是请求的位置）`);
    }
    // 读回真实位置核对
    let actualIndex = null;
    try { actualIndex = Number(newSheet.Index); } catch (e) {}
    const totalSheets = wb.Worksheets.Count;
    if (pos === "end" && actualIndex !== null && actualIndex !== totalSheets) {
      warnings.push(`position='end' 未生效：新表在第 ${actualIndex} 位，共 ${totalSheets} 张`);
    }

    return {
      success: warnings.length === 0,
      workbookName: wb.Name,
      sourceSheetName: srcSheet.Name,
      newSheetName: newSheet.Name,
      positionRequested: pos,
      positionApplied,
      actualIndex,
      totalSheets,
      warnings,
      message: `已将工作表 [${sourceSheetName}] 克隆为 [${newSheetName}]（位置 ${positionApplied}，第 ${actualIndex} 位 / 共 ${totalSheets} 张）`
    };
  }

  function saveWorkbook(app, params) {
    const { workbookName } = params || {};
    const wb = getWorkbook(app, workbookName);
    const warnings = [];
    // 从未保存过的新工作簿（Path 为空）调 Save() 会失败或弹"另存为"——如实告知而非报成功（excel-tester L-4）
    let pathBefore = "";
    try { pathBefore = String(wb.Path || ""); } catch (e) {}
    try { wb.Save(); }
    catch (e) {
      return { success: false, workbookName: wb.Name, hostError: e.message, warnings: [`保存失败: ${e.message}`],
        message: `工作簿 [${wb.Name}] 保存失败${pathBefore === "" ? "（该工作簿从未保存过，请先用 save_as 指定路径）" : ""}` };
    }
    // 读回是否真的落盘（Saved 标志）
    let savedFlag = null;
    try { savedFlag = Boolean(wb.Saved); } catch (e) {}
    if (savedFlag === false) {
      warnings.push(pathBefore === "" ? "该工作簿从未保存过，Save() 未落盘；请用 save_as 指定路径" : "宿主报告保存后仍为未保存状态，请人工确认");
    }
    return {
      success: savedFlag !== false,
      workbookName: wb.Name,
      fullName: wb.FullName,
      saved: savedFlag,
      warnings,
      message: savedFlag === false ? `工作簿 [${wb.Name}] 未确认落盘` : `工作簿 [${wb.Name}] 已成功保存到磁盘`
    };
  }

  // 11. 智能列宽排版 (AutoFit + 宽度上下限与自动换行)
  function autoFitColumns(app, params) {
    const { sheetName, columnRules, workbookName } = params;
    const sheet = getWorksheet(app, sheetName, workbookName);

    const results = [];

    if (!Array.isArray(columnRules) || columnRules.length === 0) {
      // 说明承诺"不传 columnRules 则自适应全表已用区域"，但原实现直接返回 success + results: []，
      // 列宽纹丝不动（静默 no-op，问题台账 ISS-39）。这里按已用区域的列范围逐列自适应。
      const used = sheet.UsedRange;
      const firstColumn = used && typeof used.Column === "number" ? used.Column : 1;
      const columnCount = used && used.Columns && used.Columns.Count ? used.Columns.Count : 0;
      for (let index = firstColumn; index < firstColumn + columnCount; index++) {
        const col = sheet.Columns.Item(index);
        col.AutoFit();
        results.push({ colIndex: index, finalWidth: col.ColumnWidth, wrapped: !!col.WrapText });
      }
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        mode: "usedRange",
        results,
        message: columnCount > 0
          ? `已按已用区域自适应 ${results.length} 列`
          : "已用区域为空，未调整任何列"
      };
    }

    for (const rule of columnRules) {
      const { colIndex, minWidth = 12, maxWidth = 30, wrapText = true } = rule;
      const col = sheet.Columns.Item(colIndex);

      // 先执行自适应计算
      col.AutoFit();
      let currentWidth = col.ColumnWidth;

      // 应用下限
      if (currentWidth < minWidth) {
        currentWidth = minWidth;
      }
      // 应用上限
      if (maxWidth && currentWidth > maxWidth) {
        currentWidth = maxWidth;
        if (wrapText) {
          col.WrapText = true;
        }
      }
      col.ColumnWidth = currentWidth;

      results.push({
        colIndex: colIndex,
        finalWidth: currentWidth,
        wrapped: col.WrapText
      });
    }

    return { success: true, workbookName: sheet.Parent.Name, results };
  }

  // 12. 原子回滚
  function rollbackCells(app, params) {
    const { sheetName, address, snapshot, workbookName } = params;
    if (!address || !snapshot) throw new Error("缺少必要回滚参数");

    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);

    if (snapshot.formulas) {
      range.Formula = snapshot.formulas;
    } else if (snapshot.values) {
      range.Value2 = snapshot.values;
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      address: range.Address(),
      restoredRows: range.Rows.Count,
      restoredCols: range.Columns.Count
    };
  }

  // 13. 插入行或列
  function insertDimension(app, params) {
    const { sheetName, type, index = 1, count = 1, workbookName } = params;
    const sheet = getWorksheet(app, sheetName, workbookName);

    if (type === "column") {
      const colRange = sheet.Columns.Item(index);
      for (let i = 0; i < count; i++) colRange.Insert();
      return { success: true, message: `在列 ${index} 插入 ${count} 列` };
    } else if (type === "row") {
      const rowRange = sheet.Rows.Item(index);
      for (let i = 0; i < count; i++) rowRange.Insert();
      return { success: true, message: `在行 ${index} 插入 ${count} 行` };
    }
    throw new Error("type 必须是 'row' 或 'column'");
  }

  function normalize2DArray(val, expectedRows, expectedCols) {
    if (val === undefined || val === null) {
      return Array(expectedRows).fill(null).map(() => Array(expectedCols).fill(null));
    }
    if (!Array.isArray(val)) return [[val]];
    if (!Array.isArray(val[0])) return [val];
    return val;
  }

  function hookWpsEvents() {
    if (eventsHooked) return;
    try {
      if (typeof wps !== "undefined" && wps.ApiEvent && wps.ApiEvent.AddApiEventListener) {
        wps.ApiEvent.AddApiEventListener("SheetSelectionChange", onSelectionChange);
        eventsHooked = true;
        log("已挂载 WPS 选区事件监听");
      }
    } catch (e) {
      log("挂载 ApiEventListener 提示: " + e.message);
    }
  }

  function onSelectionChange(sh, targetRange) {
    if (!isConnected) return;
    try {
      const addr = targetRange ? targetRange.Address() : "";
      if (addr && addr !== lastSelectionAddress) {
        lastSelectionAddress = addr;
        sendPacket({
          type: "event",
          event: "selection_change",
          data: {
            sheetName: sh ? sh.Name : "",
            address: addr,
            rowCount: targetRange.Rows.Count,
            columnCount: targetRange.Columns.Count
          }
        });
      }
    } catch (e) {}
  }

  // ── CAP-53 新建工作簿
  //
  // 之前**没有这个工具**（Word 有 word_create_document、PPT 有 new_presentation，Excel 漏了）。
  // 但宿主一直支持：真机实测 `app.Workbooks.Add()` + `wb.SaveAs(路径)` 均可用。
  // 「工具缺失」不等于「宿主不能做」——这是补工具，不是补能力。
  function createWorkbook(app, params) {
    const { savePath, sheetName } = params || {};
    const before = Number(app.Workbooks.Count);
    const wb = app.Workbooks.Add();
    const created = wb && wb.Name ? String(wb.Name) : null;

    let saved = null, saveError = null;
    if (savePath) {
      // ⚠️ **必须关掉宿主对话框**：目标路径已存在时 `SaveAs` 会弹"是否覆盖"的**模态框**，
      // 它会阻塞整个 WPS——所有后续调用全部无响应，而且没人点它就一直卡着。
      // （真机踩到：create_workbook 超时，WPS 被模态框阻塞，四组件全部无响应。）
      const prevAlerts = (() => { try { return app.DisplayAlerts; } catch (e) { return null; } })();
      try { app.DisplayAlerts = false; } catch (e) {}
      try {
        wb.SaveAs(String(savePath), 51);
        saved = String(savePath);
        // ⚠️ **不能以"没抛异常"判定保存成功**：DisplayAlerts=false 时，
        // 宿主对**已存在的路径**既不抛异常也不落盘
        // （excel-tester H-4：返回"已保存"但磁盘 mtime 完全没变）。
        // 读回 `wb.Path` 与 `wb.Saved` 才是判据。
        const savedPathBack = (() => { try { return String(wb.Path || ""); } catch (e) { return ""; } })();
        const savedFlag = (() => { try { return Boolean(wb.Saved); } catch (e) { return null; } })();
        if (savedPathBack === "" || savedFlag === false) {
          saveError = `宿主未真正落盘（Path="${savedPathBack}" Saved=${savedFlag}）`;
          saved = null;
        }
      } catch (e) { saveError = String(e.message || e); }
      finally {
        // 无论成功失败都要恢复，否则后续所有操作都静默吞掉提示
        if (prevAlerts !== null) { try { app.DisplayAlerts = prevAlerts; } catch (e) {} }
        else { try { app.DisplayAlerts = true; } catch (e) {} }
      }
    }
    // 首张工作表改名（可选）
    let firstSheet = null;
    try {
      const ws = wb.Worksheets.Item(1);
      if (sheetName) { ws.Name = String(sheetName); }
      firstSheet = String(ws.Name);
    } catch (e) {}

    // 读回核对：工作簿确实存在、页数、首表名、是否已落盘
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
    const after = Number(app.Workbooks.Count);
    const warnings = [];
    if (after !== before + 1) warnings.push(`新建后工作簿数 ${after}，期望 ${before + 1}`);
    if (savePath && !saved) warnings.push(`保存失败：${saveError}`);

    return {
      success: after === before + 1 && (!savePath || !!saved),
      workbookName: g(() => String(wb.Name), created),
      sheetCount: g(() => Number(wb.Worksheets.Count), null),
      firstSheetName: firstSheet,
      savedPath: saved,
      warnings,
      message: `已新建工作簿 [${g(() => String(wb.Name), created)}]（${g(() => Number(wb.Worksheets.Count), "?")} 张表，首表 ${firstSheet}）${saved ? `并保存到 ${saved}` : ""}`
    };
  }

  // ── CAP-22 图表导图（clear_range 早已存在，此处不重复实现）
  //
  // 真机探测确认（WPS 12.1.28496）：
  //   `Chart.Export(路径, "PNG")` 可用且**确实落盘**（实测导出 8040 字节 PNG）。
  // 注意：outputPath 必须落在 **WPS 可写目录**内，写到 /tmp 之类读不到的位置
  // 会"调用成功但不落盘"，所以桥接侧会再用 existsSync 核对一次。

  /** 把图表导出为图片文件。 */
  function exportChartImage(app, params) {
    const { sheetName, workbookName, chartName, chartIndex, outputPath, format = "PNG" } = params || {};
    if (!outputPath) throw new Error("缺少必要参数: outputPath（必须落在 WPS 可写目录内）");
    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}

    const charts = [];
    const shapes = sheet.Shapes;
    for (let i = 1; i <= shapes.Count; i++) {
      const sh = shapes.Item(i);
      let isChart = false;
      try { isChart = Boolean(sh.HasChart); } catch (e) {}
      if (isChart) charts.push(sh);
    }
    if (!charts.length) throw new Error(`工作表 [${sheet.Name}] 上没有图表`);

    let target = null;
    if (chartIndex !== undefined && chartIndex !== null && chartIndex !== "") {
      const idx = Number(chartIndex);
      if (!(idx >= 1) || idx > charts.length) throw new Error(`图表序号 ${idx} 越界（共 ${charts.length} 个）`);
      target = charts[idx - 1];
    } else if (chartName) {
      for (const sh of charts) { try { if (String(sh.Name) === String(chartName)) { target = sh; break; } } catch (e) {} }
      if (!target) {
        const names = charts.map(s => { try { return String(s.Name); } catch (e) { return "?"; } });
        throw new Error(`找不到名为 "${chartName}" 的图表。现有图表：${names.join(" / ")}`);
      }
    } else if (charts.length === 1) {
      target = charts[0];
    } else {
      const names = charts.map(s => { try { return String(s.Name); } catch (e) { return "?"; } });
      throw new Error(`工作表上有 ${charts.length} 个图表，请用 chartName 或 chartIndex 指定。现有：${names.join(" / ")}`);
    }

    const fmt = String(format).toUpperCase();
    const ALLOWED = ["PNG", "JPG", "JPEG", "GIF", "BMP"];
    if (ALLOWED.indexOf(fmt) < 0) throw new Error(`不支持的格式: ${format}（可用 ${ALLOWED.join(" / ")}）`);
    const exportFmt = fmt === "JPEG" ? "JPG" : fmt;

    target.Chart.Export(String(outputPath), exportFmt);

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      chartName: (() => { try { return String(target.Name); } catch (e) { return null; } })(),
      outputPath: String(outputPath),
      format: exportFmt,
      // 宿主侧无法探测文件系统，落盘由**桥接侧**用 existsSync 核对后回填 fileWritten/fileSizeBytes
      fileWritten: null,
      warnings: [],
      message: `已调用图表导出 → ${outputPath}（${exportFmt}）`
    };
  }

  // ── CAP-21 图表更新（wps_update_chart）
  //
  // 此前 `wps_update_chart` 只在路由表里声明，WPS 加载项**没有 RPC 分支**，
  // 调用会被明确拒绝（宿主未实现）。真机确认宿主 API 可用：
  // `shape.Chart.ChartTitle.Text` 可读可写、`ChartType` 可改、系列可读。
  /** 就地更新已有图表：标题、图表类型、图例、数据标签。写后逐项读回核对。 */
  function updateChart(app, params) {
    const { sheetName, workbookName, chartName, chartIndex, title, chartType, hasLegend, showDataLabels,
            fontName, legendPosition, dataLabelColorMatchesSeries, position } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}

    // 定位图表形状（HasChart 为真）
    let target = null;
    const charts = [];
    const shapes = sheet.Shapes;
    for (let i = 1; i <= shapes.Count; i++) {
      const sh = shapes.Item(i);
      let isChart = false;
      try { isChart = Boolean(sh.HasChart); } catch (e) {}
      if (!isChart) continue;
      charts.push(sh);
    }
    if (!charts.length) throw new Error(`工作表 [${sheet.Name}] 上没有图表`);
    if (chartName || chartIndex !== undefined) {
      if (chartIndex !== undefined) {
        const idx = Number(chartIndex);
        if (idx < 1 || idx > charts.length) throw new Error(`图表序号 ${idx} 越界（共 ${charts.length} 个）`);
        target = charts[idx - 1];
      } else {
        for (const sh of charts) { try { if (String(sh.Name) === String(chartName)) { target = sh; break; } } catch (e) {} }
        if (!target) throw new Error(`找不到名为 "${chartName}" 的图表。现有图表：${charts.map(s => { try { return String(s.Name); } catch (e) { return "?"; } }).join(" / ")}`);
      }
    } else if (charts.length === 1) {
      target = charts[0];
    } else {
      throw new Error(`工作表上有 ${charts.length} 个图表，请用 chartName 或 chartIndex 指定。现有：${charts.map(s => { try { return String(s.Name); } catch (e) { return "?"; } }).join(" / ")}`);
    }

    const ch = target.Chart;
    const warnings = [];
    const applied = {};

    if (title !== undefined) {
      try { ch.HasTitle = true; ch.ChartTitle.Text = String(title); applied.title = String(title); }
      catch (e) { warnings.push(`设置标题失败: ${e.message}`); }
    }
    if (chartType !== undefined) {
      // 常用 xlChartType：柱状簇状=51 / 折线=4 / 饼图=5 / 条形簇状=57 / 散点=75 / 面积=1
      const TYPES = { column: 51, column_clustered: 51, bar: 57, bar_clustered: 57, line: 4, pie: 5, scatter: 75, area: 1 };
      const code = typeof chartType === "number" ? chartType : TYPES[String(chartType).toLowerCase()];
      if (code === undefined) { warnings.push(`不认识的 chartType: ${chartType}`); }
      else { try { ch.ChartType = code; applied.chartType = code; } catch (e) { warnings.push(`设置图表类型失败: ${e.message}`); } }
    }
    if (hasLegend !== undefined) {
      try { ch.HasLegend = Boolean(hasLegend); applied.hasLegend = Boolean(hasLegend); } catch (e) { warnings.push(`设置图例失败: ${e.message}`); }
    }
    // 图表字体：标题/图例/坐标轴/数据标签统一换成指定字体
    // （宿主默认走宋体，中英文混排观感差）
    if (fontName) {
      const f = String(fontName);
      const setFont = (obj, label) => { try { obj.Font.Name = f; } catch (e) { warnings.push(`设置${label}字体失败: ${e.message}`); } };
      try { setFont(ch.ChartArea, "图表区"); } catch (e) {}
      try { if (ch.HasTitle) setFont(ch.ChartTitle, "标题"); } catch (e) {}
      try { if (ch.HasLegend) setFont(ch.Legend, "图例"); } catch (e) {}
      // 坐标轴字体要走 **TickLabels**：WPS 里  是 undefined（设不上），
      // 刻度标签才挂得住 Font。真机探明：Axes().Item(n).TickLabels.Font 可读可写。
      const setTickFont = (idx) => {
        try {
          const ax = ch.Axes().Item(idx);
          ax.TickLabels.Font.Name = f;
          return ax.TickLabels.Font.Name === f;
        } catch (e) { return false; }
      };
      if (!setTickFont(1)) warnings.push("设置分类轴刻度字体失败");
      if (!setTickFont(2)) warnings.push("设置数值轴刻度字体失败");
      applied.fontName = f;
    }

    // 图例位置：xlLegendPositionRight=-4152 / Left=-4131 / Top=-4160 / Bottom=-4107
    if (legendPosition !== undefined && legendPosition !== null && legendPosition !== "") {
      const LP = { right: -4152, left: -4131, top: -4160, bottom: -4107 };
      const key = String(legendPosition).toLowerCase();
      const code = LP[key] !== undefined ? LP[key] : (Number.isFinite(Number(legendPosition)) ? Number(legendPosition) : undefined);
      if (code === undefined) warnings.push(`不认识的 legendPosition: ${legendPosition}（可用 right/left/top/bottom）`);
      else {
        try { ch.HasLegend = true; ch.Legend.Position = code; applied.legendPosition = key; }
        catch (e) { warnings.push(`设置图例位置失败: ${e.message}`); }
      }
    }

    // 数据标签文字颜色与所属柱/点颜色一致
    if (dataLabelColorMatchesSeries) {
      let matched = 0;
      try {
        const sc = ch.SeriesCollection();
        for (let i = 1; i <= Number(sc.Count); i++) {
          const s = sc.Item(i);
          let rgb = null;
          try { rgb = Number(s.Format.Fill.ForeColor.RGB); } catch (e) {}
          if (!Number.isFinite(rgb) || rgb < 0) { try { rgb = Number(s.Border.Color); } catch (e) {} }
          if (!Number.isFinite(rgb) || rgb < 0) continue;
          try {
            s.HasDataLabels = true;
            s.DataLabels().Font.Color = rgb;
            matched++;
          } catch (e) { warnings.push(`设置系列 ${i} 数据标签颜色失败: ${e.message}`); }
        }
      } catch (e) { warnings.push(`设置数据标签颜色失败: ${e.message}`); }
      applied.dataLabelColorMatchesSeries = matched;
      if (!matched) warnings.push("没有系列成功匹配到颜色，数据标签颜色未改");
    }

    // 移动/改尺寸（单位：磅）。支持绝对定位与相对位移——
    // 「把图表往下挪一点」这种调整本该一条指令完成，不必重建图表或跑脚本。
    if (position && typeof position === "object") {
      const moved = {};
      const setIf = (key, cur, next) => { if (Number.isFinite(Number(next))) { try { target[key] = Number(next); moved[key] = Number(next); } catch (e) { warnings.push(`设置 ${key} 失败: ${e.message}`); } } };
      setIf("Left", null, position.left);
      setIf("Top", null, position.top);
      setIf("Width", null, position.width);
      setIf("Height", null, position.height);
      // 相对位移：在当前位置基础上加减
      if (Number.isFinite(Number(position.leftDelta))) {
        try { target.Left = Number(target.Left) + Number(position.leftDelta); moved.left = Number(target.Left); } catch (e) { warnings.push(`左移失败: ${e.message}`); }
      }
      if (Number.isFinite(Number(position.topDelta))) {
        try { target.Top = Number(target.Top) + Number(position.topDelta); moved.top = Number(target.Top); } catch (e) { warnings.push(`下移失败: ${e.message}`); }
      }
      // 单元格锚点：左上角对齐到指定单元格
      if (position.leftCell) {
        try { const c = sheet.Range(String(position.leftCell)); target.Left = Number(c.Left) + 2; target.Top = Number(c.Top) + 2; moved.leftCell = String(position.leftCell); } catch (e) { warnings.push(`按单元格定位失败: ${e.message}`); }
      }
      applied.position = moved;
    }

    if (showDataLabels !== undefined) {
      // ⚠️ 不能用 `chart.ApplyDataLabels(0)` 关标签：真机实测它**关不掉**，
      // 但调用不报错——于是工具返回 success 而图表上标签仍在，属"假成功"。
      // 改为**逐个系列设置 HasDataLabels**，再读回核对，报**真实状态**而不是请求值。
      const want = Boolean(showDataLabels);
      let touched = 0;
      try {
        const sc = ch.SeriesCollection();
        const n = Number(sc.Count);
        for (let i = 1; i <= n; i++) { try { sc.Item(i).HasDataLabels = want; touched++; } catch (e) {} }
      } catch (e) { warnings.push(`设置数据标签失败: ${e.message}`); }
      if (!touched) {
        try { ch.ApplyDataLabels(want ? 2 : 0); touched = 1; } catch (e) { warnings.push(`设置数据标签失败: ${e.message}`); }
      }
      // 读回每个系列的真实状态
      const states = [];
      try {
        const sc = ch.SeriesCollection();
        for (let i = 1; i <= Number(sc.Count); i++) { try { states.push(Boolean(sc.Item(i).HasDataLabels)); } catch (e) {} }
      } catch (e) {}
      const actualAll = states.length ? states.every(Boolean) : null;
      const actualNone = states.length ? states.every(v => !v) : null;
      applied.showDataLabels = want;
      // 名字直说含义："实际状态与请求一致"，避免被读成"标签是开着的"
      applied.dataLabelsMatchRequest = want ? actualAll : actualNone;
      if (states.length && (want ? !actualAll : !actualNone)) {
        warnings.push(`数据标签**未按请求生效**：请求 ${want}，各系列实际为 [${states.join(", ")}]`);
      }
    }

    // 读回核对
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
    const actual = {
      name: g(() => String(target.Name), null),
      chartType: g(() => Number(ch.ChartType), null),
      title: g(() => String(ch.ChartTitle.Text), null),
      hasTitle: g(() => Boolean(ch.HasTitle), null),
      hasLegend: g(() => Boolean(ch.HasLegend), null),
      seriesCount: g(() => Number(ch.SeriesCollection().Count), null),
      // 数据标签的真实状态（逐系列读），让"写没写进去"可核对
      dataLabels: (() => {
        try {
          const sc = ch.SeriesCollection(); const out = [];
          for (let i = 1; i <= Number(sc.Count); i++) { try { out.push(Boolean(sc.Item(i).HasDataLabels)); } catch (e) {} }
          return out.length ? (out.every(Boolean) ? true : out.every(v => !v) ? false : out) : null;
        } catch (e) { return null; }
      })()
    };
    if (title !== undefined && actual.title !== String(title)) warnings.push(`标题读回为「${actual.title}」，与请求不一致`);

    return {
      success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name,
      chartName: actual.name, applied, chart: actual, warnings,
      message: `已更新图表「${actual.name}」${Object.keys(applied).length} 项${warnings.length ? `（${warnings.length} 条告警）` : ""}`
    };
  }

  // ── CAP-15~20 第二类：表格常用能力（区域复制 / 超链接 / 命名区域 / 文档属性 / 结构化表格 / 图片）
  //
  // 依据真机探测（WPS 12.1.28496）确认可用：
  //   Range.Copy(dest) ✔ · Hyperlinks.Add/Count ✔ · wb.Names.Add ✔ ·
  //   wb.BuiltinDocumentProperties.Item(name).Value ✔ · CustomDocumentProperties.Add ✔ ·
  //   ws.ListObjects.Add(1, range, null, 1) ✔ · ws.Shapes.AddPicture ✔

  /** 把一处区域的值/公式/格式复制到另一处。 */
  function copyRange(app, params) {
    const { sheetName, workbookName, sourceRange, destRange, destSheetName,
            copyType = "all", skipBlanks = false, transpose = false } = params || {};
    if (!sourceRange) throw new Error("缺少必要参数: sourceRange");
    if (!destRange) throw new Error("缺少必要参数: destRange");
    const sheet = getWorksheet(app, sheetName, workbookName);
    const src = sheet.Range(sourceRange);

    // 目标在不同表时先取目标表，并用全限定地址，避免依赖"活动表"状态
    const destSheet = destSheetName ? getWorksheet(app, destSheetName, workbookName) : sheet;
    const dest = destSheet.Range(destRange);

    // xlPasteAll=-4104 / xlPasteValues=-4163 / xlPasteFormats=-4122 / xlPasteFormulas=-4123
    const PASTE = { all: -4104, values: -4163, formats: -4122, formulas: -4123 };
    const code = PASTE[String(copyType).toLowerCase()];
    if (code === undefined) throw new Error(`不支持的 copyType: ${copyType}（可用 all / values / formats / formulas）`);

    if (transpose) src.Copy();
    if (transpose) { dest.PasteSpecial(-4122, -4142, false, true); }   // xlPasteAll, xlPasteSpecialOperationNone, SkipBlanks=false, Transpose=true
    else src.Copy(dest);

    // 读回核对：目标区域左上角必须拿到源的值（空源不算失败）
    let readBack = null;
    try { readBack = destSheet.Range(destRange.split(":")[0]).Value2; } catch (e) {}
    let srcValue = null;
    try { srcValue = src.Value2; } catch (e) {}
    const srcFirst = Array.isArray(srcValue) ? (Array.isArray(srcValue[0]) ? srcValue[0][0] : srcValue[0]) : srcValue;

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      source: `${sheet.Name}!${sourceRange}`,
      dest: `${destSheet.Name}!${destRange}`,
      copyType,
      readBackFirstCell: readBack === undefined ? null : readBack,
      sourceFirstCell: srcFirst === undefined ? null : srcFirst,
      warnings: [],
      message: `已把 [${sheet.Name}] ${sourceRange} 复制到 [${destSheet.Name}] ${destRange}（${copyType}）`
    };
  }

  /** 超链接增删查。 */
  function manageHyperlink(app, params) {
    const { sheetName, workbookName, action = "list", address, url, displayText, tooltip, emailSubject, targetAddress } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };

    if (action === "list") {
      const items = [];
      const n = g(() => Number(sheet.Hyperlinks.Count), 0);
      for (let i = 1; i <= n; i++) {
        const h = (() => { try { return sheet.Hyperlinks.Item(i); } catch (e) { return null; } })();
        if (!h) continue;
        items.push({
          index: i,
          address: g(() => String(h.Address), null),
          subAddress: g(() => String(h.SubAddress), null),
          text: g(() => String(h.TextToDisplay), null),
          tooltip: g(() => String(h.ScreenTip), null),
          anchor: g(() => String(h.Range.Address()), null),
          type: g(() => Number(h.Type), null)
        });
      }
      return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, count: items.length, hyperlinks: items, warnings: [],
        message: `工作表 [${sheet.Name}] 共 ${items.length} 个超链接` };
    }

    if (action === "delete") {
      if (!address) throw new Error("delete 需要 address（锚点单元格）或省略 address 清空全部");
      let removed = 0;
      for (let i = g(() => Number(sheet.Hyperlinks.Count), 0); i >= 1; i--) {
        const h = (() => { try { return sheet.Hyperlinks.Item(i); } catch (e) { return null; } })();
        if (!h) continue;
        const a = g(() => String(h.Range.Address()), "");
        if (a === sheet.Range(address).Address() || !address) { h.Delete(); removed++; }
      }
      return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, removed,
        remaining: g(() => Number(sheet.Hyperlinks.Count), null), warnings: [], message: `已删除 ${removed} 个超链接` };
    }

    // action = add
    if (!address) throw new Error("add 需要 address（锚点单元格）");
    // L-5：只传 targetAddress（文档内跳转，如 'Sheet2!A1'）也是合法用法，
    // 原先强制要求 url → 单独传 targetAddress 直接被拒，与 schema 暴露的参数不符。
    if (!url && !targetAddress) throw new Error("add 需要 url（外部链接）或 targetAddress（文档内跳转）");
    const anchor = sheet.Range(address);
    const linkAddr = String(url || "");
    const subAddr = String(targetAddress || "");
    const disp = String(displayText || url || targetAddress || "链接");
    // Hyperlinks.Add(Anchor, Address, SubAddress, ScreenTip, TextToDisplay)
    sheet.Hyperlinks.Add(anchor, linkAddr, subAddr, String(tooltip || ""), disp);
    // 读回核对
    let ok = false, readText = null;
    try { const h = anchor.Hyperlinks.Item(1); ok = true; readText = String(h.TextToDisplay); } catch (e) {}
    return {
      success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, address, url: linkAddr, targetAddress: subAddr,
      verified: ok, readBackText: readText,
      warnings: ok ? [] : ["写入后未能从锚点读到超链接，请人工确认"],
      message: `已在 [${sheet.Name}] ${address} 添加超链接 → ${url}`
    };
  }

  /** 命名区域增删查。 */
  function manageNamedRange(app, params) {
    const { workbookName, action = "list", name, refersTo, comment } = params || {};
    const wb = getWorkbook(app, workbookName);
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };

    if (action === "list") {
      const items = [];
      const n = g(() => Number(wb.Names.Count), 0);
      for (let i = 1; i <= n; i++) {
        const nm = (() => { try { return wb.Names.Item(i); } catch (e) { return null; } })();
        if (!nm) continue;
        items.push({ index: i, name: g(() => String(nm.Name), null), refersTo: g(() => String(nm.RefersTo), null),
                     visible: g(() => Boolean(nm.Visible), null), comment: g(() => String(nm.Comment), null) });
      }
      return { success: true, workbookName: wb.Name, count: items.length, names: items, warnings: [],
        message: `工作簿 [${wb.Name}] 共 ${items.length} 个命名区域` };
    }

    if (action === "delete") {
      if (!name) throw new Error("delete 需要 name");
      wb.Names.Item(String(name)).Delete();
      // ⚠️ 不能用 `Names.Item(name)` 探测是否还存在：删除后它返回 **null 而不抛错**，
      // 于是"没抛错＝还在"的判断**恒为 true**（真机踩到：删除明明成功，stillExists 却报 true，
      // 与实际自检结果矛盾）。改为**枚举全部名称**再比对，这是可靠的。
      const still = (() => {
        try {
          const n = Number(wb.Names.Count);
          for (let i = 1; i <= n; i++) { if (String(wb.Names.Item(i).Name) === String(name)) return true; }
        } catch (e) {}
        return false;
      })();
      const remaining = (() => { try { const out = []; const n = Number(wb.Names.Count); for (let i = 1; i <= n; i++) out.push(String(wb.Names.Item(i).Name)); return out; } catch (e) { return []; } })();
      return { success: !still, workbookName: wb.Name, deleted: name, stillExists: still, remainingNames: remaining.length, warnings: still ? [`删除后仍能枚举到 ${name}，删除未生效`] : [],
        message: still ? `命名区域 ${name} 删除失败（仍存在）` : `已删除命名区域 ${name}，剩余 ${remaining.length} 个` };
    }

    if (!name || !refersTo) throw new Error("add 需要 name 和 refersTo（如 'Sheet1!$A$1:$B$10'）");
    wb.Names.Add(String(name), String(refersTo), false, String(comment || ""));
    const back = g(() => String(wb.Names.Item(String(name)).RefersTo), null);
    // comment 是否真的写进去了（M-10：原先被静默忽略还不告警）
    const warnings = [];
    if (!back) warnings.push("写入后读不回该名称");
    if (comment) {
      const backComment = g(() => String(wb.Names.Item(String(name)).Comment), null);
      if (backComment !== String(comment)) {
        warnings.push(`comment 未生效：请求 ${JSON.stringify(String(comment))}，宿主读回 ${JSON.stringify(backComment)}`);
      }
    }
    return { success: !!back, workbookName: wb.Name, name, requested: refersTo, readBack: back,
      readBackComment: comment ? g(() => String(wb.Names.Item(String(name)).Comment), null) : undefined,
      verified: !!back, warnings,
      message: back ? `已添加命名区域 ${name} → ${back}` : `命名区域 ${name} 添加失败（读不回）` };
  }

  /** 文档属性（内置 + 自定义）。 */
  function manageDocumentProperties(app, params) {
    const { workbookName, action = "read", properties } = params || {};
    const wb = getWorkbook(app, workbookName);
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
    const BUILTIN = ["Title", "Subject", "Author", "Keywords", "Comments", "Category", "Company", "Manager"];

    const readAll = () => {
      const builtin = {};
      for (const k of BUILTIN) {
        builtin[k] = g(() => {
          const p = wb.BuiltinDocumentProperties.Item(k);
          const v = p && p.Value !== undefined ? p.Value : p;
          return v === undefined || v === null ? "" : String(v);
        }, "");
      }
      const custom = {};
      const n = g(() => Number(wb.CustomDocumentProperties.Count), 0);
      for (let i = 1; i <= n; i++) {
        const p = (() => { try { return wb.CustomDocumentProperties.Item(i); } catch (e) { return null; } })();
        if (!p) continue;
        custom[g(() => String(p.Name), "prop" + i)] = g(() => String(p.Value), "");
      }
      return { builtin, custom };
    };

    if (action === "read") {
      const cur = readAll();
      return { success: true, workbookName: wb.Name, builtin: cur.builtin, custom: cur.custom, warnings: [],
        message: `工作簿 [${wb.Name}] 属性：标题「${cur.builtin.Title || "（空）"}」作者「${cur.builtin.Author || "（空）"}」，自定义 ${Object.keys(cur.custom).length} 项` };
    }

    // action = delete：按名字删除**自定义**属性（内置属性不能删，只能清空值）
    if (action === "delete") {
      const names = Array.isArray(params?.propertyNames) ? params.propertyNames.filter(Boolean).map(String) : [];
      if (!names.length) throw new Error("delete 需要 propertyNames（要删除的自定义属性名数组）");
      const removed = [], notFound = [];
      for (const want of names) {
        let hit = false;
        for (let i = g(() => Number(wb.CustomDocumentProperties.Count), 0); i >= 1; i--) {
          const pr = (() => { try { return wb.CustomDocumentProperties.Item(i); } catch (e) { return null; } })();
          if (pr && g(() => String(pr.Name), "") === want) { pr.Delete(); removed.push(want); hit = true; break; }
        }
        if (!hit) notFound.push(want);
      }
      const left = readAll();
      return {
        success: true, workbookName: wb.Name, action: "delete",
        removed, notFound,
        remainingCustom: Object.keys(left.custom).length,
        warnings: [],
        message: `已删除自定义属性 ${removed.length} 项${notFound.length ? `，${notFound.length} 项不存在` : ""}；剩余自定义属性 ${Object.keys(left.custom).length} 项`
      };
    }

    // action = apply
    const warnings = [];
    const applied = {};
    if (properties && typeof properties === "object") {
      for (const [k, v] of Object.entries(properties)) {
        if (BUILTIN.indexOf(k) >= 0) {
          try { wb.BuiltinDocumentProperties.Item(k).Value = String(v); applied[k] = String(v); }
          catch (e) { warnings.push(`设置内置属性 ${k} 失败: ${e.message}`); }
        } else {
          // 自定义属性：存在则改，不存在则加（type 4 = msoPropertyTypeString）
          let done = false;
          for (let i = 1; i <= g(() => Number(wb.CustomDocumentProperties.Count), 0); i++) {
            const p = (() => { try { return wb.CustomDocumentProperties.Item(i); } catch (e) { return null; } })();
            if (p && g(() => String(p.Name), "") === k) { try { p.Value = String(v); done = true; applied[k] = String(v); } catch (e) { warnings.push(`改自定义属性 ${k} 失败: ${e.message}`); } break; }
          }
          if (!done) {
            try { wb.CustomDocumentProperties.Add(k, false, 4, String(v)); applied[k] = String(v); }
            catch (e) { warnings.push(`新增自定义属性 ${k} 失败: ${e.message}`); }
          }
        }
      }
    }
    const after = readAll();
    const mismatched = Object.keys(applied).filter(k => {
      const got = BUILTIN.indexOf(k) >= 0 ? after.builtin[k] : after.custom[k];
      return String(got) !== String(applied[k]);
    });
    for (const k of mismatched) warnings.push(`读回与写入不一致: ${k}`);
    return { success: true, workbookName: wb.Name, action: "apply", applied, after, warnings,
      message: `已更新 ${Object.keys(applied).length} 项文档属性${mismatched.length ? `（${mismatched.length} 项读回不一致）` : ""}` };
  }

  /** 结构化表格（ListObject）增删查。 */
  function manageTable(app, params) {
    const { sheetName, workbookName, action = "list", tableName, address, styleName, newName, hasHeaders = true, totalsRow = false } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
    const readOne = (lo) => ({
      name: g(() => String(lo.Name), null),
      range: g(() => String(lo.Range.Address()), null),
      rowCount: g(() => Number(lo.ListRows.Count), null),
      colCount: g(() => Number(lo.ListColumns.Count), null),
      hasHeaders: g(() => Boolean(lo.ShowHeaders), null),
      totalsRow: g(() => Boolean(lo.ShowTotals), null),
      style: g(() => String(lo.TableStyle.Name), null),
      columns: g(() => { const out = []; const n = Number(lo.ListColumns.Count); for (let i = 1; i <= n; i++) out.push(String(lo.ListColumns.Item(i).Name)); return out; }, [])
    });

    if (action === "list") {
      const items = [];
      const n = g(() => Number(sheet.ListObjects.Count), 0);
      for (let i = 1; i <= n; i++) { const lo = (() => { try { return sheet.ListObjects.Item(i); } catch (e) { return null; } })(); if (lo) items.push(readOne(lo)); }
      return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, count: items.length, tables: items, warnings: [],
        message: `工作表 [${sheet.Name}] 共 ${items.length} 个结构化表格` };
    }

    if (action === "delete") {
      if (!tableName) throw new Error("delete 需要 tableName");
      const lo = sheet.ListObjects.Item(String(tableName));
      const rng = g(() => String(lo.Range.Address()), null);
      lo.Delete();
      return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, deleted: tableName, releasedRange: rng,
        remaining: g(() => Number(sheet.ListObjects.Count), null), warnings: [], message: `已删除结构化表格 ${tableName}（数据保留在 ${rng}）` };
    }

    if (action === "apply") {
      if (!address) throw new Error("apply 需要 address");
      const rng = sheet.Range(address);
      // xlSrcRange = 1；Add(SourceType, Source, LinkSource, XlListObjectHasHeaders)
      const lo = sheet.ListObjects.Add(1, rng, null, hasHeaders ? 1 : 2);
      if (tableName) { try { lo.Name = String(tableName); } catch (e) {} }
      if (newName) { try { lo.Name = String(newName); } catch (e) {} }
      if (styleName) { try { lo.TableStyle = String(styleName); } catch (e) {} }
      try { lo.ShowTotals = Boolean(totalsRow); } catch (e) {}
      const back = readOne(lo);
      return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, action: "apply", table: back, warnings: [],
        message: `已在 [${sheet.Name}] ${back.range} 创建结构化表格「${back.name}」（${back.rowCount} 行 × ${back.colCount} 列）` };
    }

    throw new Error(`不支持的 action: ${action}（可用 list / apply / delete）`);
  }

  /** 图片：插入 / 列出 / 删除。 */
  function managePictures(app, params) {
    const { sheetName, workbookName, action = "list", filePath, left, top, width, height, pictureName, pictureIndex } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
    const readOne = (sh, kind) => ({
      name: g(() => String(sh.Name), null), kind,
      left: g(() => Math.round(Number(sh.Left) * 100) / 100, null),
      top: g(() => Math.round(Number(sh.Top) * 100) / 100, null),
      width: g(() => Math.round(Number(sh.Width) * 100) / 100, null),
      height: g(() => Math.round(Number(sh.Height) * 100) / 100, null),
      topLeftCell: g(() => shapeAddressOf(sh.TopLeftCell), null)
    });

    if (action === "list") {
      const items = [];
      const n = g(() => Number(sheet.Shapes.Count), 0);
      for (let i = 1; i <= n; i++) {
        const sh = sheet.Shapes.Item(i);
        const isPic = g(() => Number(sh.Type) === 13, false);   // msoPicture = 13
        if (!isPic) continue;
        items.push(readOne(sh, "picture"));
      }
      return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, count: items.length, pictures: items, warnings: [],
        message: `工作表 [${sheet.Name}] 共 ${items.length} 张图片（形状总数 ${g(() => Number(sheet.Shapes.Count), 0)}）` };
    }

    if (action === "delete") {
      const shapes = sheet.Shapes;
      let removed = 0;
      for (let i = shapes.Count; i >= 1; i--) {
        const sh = shapes.Item(i);
        if (g(() => Number(sh.Type) !== 13, true)) continue;
        const nm = g(() => String(sh.Name), "");
        if (pictureName && nm !== String(pictureName)) continue;
        if (Number.isFinite(Number(pictureIndex)) && i !== Number(pictureIndex)) continue;
        sh.Delete(); removed++;
      }
      return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, removed, warnings: [], message: `已删除 ${removed} 张图片` };
    }

    if (!filePath) throw new Error("insert 需要 filePath（本机图片绝对路径）");
    const args = [String(filePath), false, true];
    if (Number.isFinite(Number(left))) args.push(Number(left));
    if (Number.isFinite(Number(top))) args.push(Number(top));
    if (Number.isFinite(Number(width))) args.push(Number(width));
    if (Number.isFinite(Number(height))) args.push(Number(height));
    const sh = sheet.Shapes.AddPicture.apply(sheet.Shapes, args);
    if (pictureName) { try { sh.Name = String(pictureName); } catch (e) {} }
    return { success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name, action: "insert",
      picture: readOne(sh, "picture"), warnings: [], message: `已插入图片 ${filePath}` };
  }

  // 13.9 工作表视图：网格线 / 行列标题 / 缩放
  //
  // 为什么必须是个独立能力：用矢量形状搭画布页时**必须先隐藏网格线**，
  // 否则形状浮在格线上、观感很乱。此前没有任何工具入口，
  // 只能让 AI 退回 wps_execute_script 手写（真机踩到过）。
  function setSheetView(app, params) {
    const { sheetName, workbookName, showGridlines, showHeadings, zoom } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}
    const win = app.ActiveWindow;
    const before = {
      showGridlines: (() => { try { return Boolean(win.DisplayGridlines); } catch (e) { return null; } })(),
      showHeadings: (() => { try { return Boolean(win.DisplayHeadings); } catch (e) { return null; } })(),
      zoom: (() => { try { return Number(win.Zoom); } catch (e) { return null; } })()
    };
    const warnings = [];
    if (showGridlines !== undefined) {
      try { win.DisplayGridlines = showGridlines ? true : false; } catch (e) { warnings.push(`设置网格线失败: ${e.message}`); }
    }
    if (showHeadings !== undefined) {
      try { win.DisplayHeadings = showHeadings ? true : false; } catch (e) { warnings.push(`设置行列标题失败: ${e.message}`); }
    }
    if (Number.isFinite(Number(zoom))) {
      try { win.Zoom = Number(zoom); } catch (e) { warnings.push(`设置缩放失败: ${e.message}`); }
    }
    // 读回核对：写没写进去必须能验证
    const after = {
      showGridlines: (() => { try { return Boolean(win.DisplayGridlines); } catch (e) { return null; } })(),
      showHeadings: (() => { try { return Boolean(win.DisplayHeadings); } catch (e) { return null; } })(),
      zoom: (() => { try { return Number(win.Zoom); } catch (e) { return null; } })()
    };
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      before, after, warnings,
      message: `工作表 [${sheet.Name}] 视图：网格线 ${after.showGridlines ? "显示" : "隐藏"}、行列标题 ${after.showHeadings ? "显示" : "隐藏"}、缩放 ${after.zoom}%`
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

  // 15. 原生矢量图表生成与管理 (第二梯队 - 健壮事务版)
  function addChart(app, params) {
    const {
      sheetName,
      workbookName,
      chartType = "column_clustered",
      dataRange,
      dataRanges,
      title,
      position,
      hasLegend = true,
      hasDataLabels = false,
      smoothLine = false,
      seriesColors,
      yAxis,
      seriesSettings,
      replaceExisting = true,
      left: explicitLeft,
      top: explicitTop,
      width: explicitWidth,
      height: explicitHeight,
      startCell,
      endCell,
      cellRange
    } = params || {};

    // 1. 强制公式全局重算（彻底根治跨表引用公式未计算导致分类轴塌陷的严重时钟竞争 Bug！）
    try {
      if (app && app.Calculate) {
        app.Calculate();
      }
    } catch (calcErr) {
      log("公式重算提示: " + calcErr.message);
    }

    // 2. 解析数据源（支持非连续区域联合数据源，如 ['A4:A19', 'E4:E19']）
    let targetDataRange = dataRange;
    if (!targetDataRange && Array.isArray(dataRanges) && dataRanges.length > 0) {
      targetDataRange = dataRanges.join(",");
    }
    if (!targetDataRange) {
      throw new Error("缺少必要参数: dataRange (例如 'A4:E19') 或 dataRanges (例如 ['A4:A19', 'E4:E19'])");
    }

    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}

    /** 本次建图的告警集合（类型降级、单系列多色等），随返回体带出。 */
    const warnings = [];

    // 计算图表位置与尺寸
    let left = 360;
    let top = 40;
    let width = 480;
    let height = 280;

    if (position) {
      if (position.leftCell) {
        try {
          const anchorRange = sheet.Range(position.leftCell);
          left = anchorRange.Left;
          top = anchorRange.Top;
        } catch (e) {
          log("图表定位 leftCell 警告: " + e.message);
        }
      }
      if (position.width) width = Number(position.width);
      if (position.height) height = Number(position.height);
    }

    // 原实现只认 `position.leftCell/width/height`，而 startCell / cellRange / endCell 与
    // 顶层 left/top/width/height **全被静默忽略**，多图会叠在默认的 360/40（问题台账 ISS-18）。
    // 这里把它们真正接上，并让"锚点"统一走一套解析顺序：position.leftCell > startCell/cellRange > 顶层像素。
    const anchorRef = (position && position.leftCell)
      ? String(position.leftCell)
      : (startCell ? String(startCell) : (cellRange ? String(cellRange).split(":")[0] : null));
    if (anchorRef && !(position && position.leftCell)) {
      try {
        const anchorRange = sheet.Range(anchorRef.includes(":") ? anchorRef.split(":")[0] : anchorRef);
        left = anchorRange.Left;
        top = anchorRange.Top;
      } catch (e) {
        log("图表定位锚点警告: " + e.message);
      }
    }
    // 顶层像素参数直接生效（显式传入优先于上面的锚点推算）
    if (Number.isFinite(Number(explicitLeft))) left = Number(explicitLeft);
    if (Number.isFinite(Number(explicitTop))) top = Number(explicitTop);
    if (Number.isFinite(Number(explicitWidth))) width = Number(explicitWidth);
    if (Number.isFinite(Number(explicitHeight))) height = Number(explicitHeight);
    // endCell：用"锚点单元格 → endCell"的矩形尺寸作为图表宽高
    if (endCell) {
      try {
        const from = sheet.Range(anchorRef ? anchorRef.split(":")[0] : "A1");
        const to = sheet.Range(String(endCell));
        const w = to.Left + to.Width - from.Left;
        const h = to.Top + to.Height - from.Top;
        if (w > 0) width = w;
        if (h > 0) height = h;
      } catch (e) {
        log("图表 endCell 尺寸推算警告: " + e.message);
      }
    }

    // 3. 覆盖模式（如果开启 replaceExisting，先清理该锚点处的重叠旧图，杜绝废图堆叠）
    if (replaceExisting && position && position.leftCell) {
      try {
        const shapes = sheet.Shapes;
        const total = shapes.Count;
        for (let i = total; i >= 1; i--) {
          const shp = shapes.Item(i);
          if (shp.HasChart && Math.abs(shp.Left - left) < 25 && Math.abs(shp.Top - top) < 25) {
            log(`自动清理锚点 [${position.leftCell}] 处的旧图表: ${shp.Name}`);
            shp.Delete();
          }
        }
      } catch (cleanErr) {
        log("清理旧图表提示: " + cleanErr.message);
      }
    }

    // 映射 Excel 图表类型常量
    // xlLine=4, xlColumnClustered=51, xlBarClustered=57, xlPie=5, xlDoughnut=-4120, xlPareto=122,
    // xlArea=1, xlXYScatter=-4169
    const chartTypeMap = {
      line: 4,
      column: 51,
      column_clustered: 51,
      bar: 57,
      bar_clustered: 57,
      pie: 5,
      doughnut: -4120,
      pareto: 122,
      area: 1,
      scatter: -4169
    };
    // 原来未知类型会**静默退回柱状图**并返回 success（问题台账 ISS-17）：调用方拿到的图不是要的类型却毫无察觉。
    // 现在改为显式报错，并带上可用类型清单。
    if (chartTypeMap[chartType] === undefined) {
      throw new Error(
        `不支持的图表类型: ${chartType}（可用: ${Object.keys(chartTypeMap).join(" / ")}）`
      );
    }
    const xlChartType = chartTypeMap[chartType];

    let shape = null;
    try {
      if (sheet.Shapes.AddChart2) {
        shape = sheet.Shapes.AddChart2(201, xlChartType, left, top, width, height);
      } else {
        shape = sheet.Shapes.AddChart(xlChartType, left, top, width, height);
      }
    } catch (err) {
      // 降级为默认簇状柱状图
      shape = sheet.Shapes.AddChart(51, left, top, width, height);
    }

    if (!shape || !shape.Chart) {
      throw new Error("无法在当前工作表中创建原生图表对象");
    }

    // 4. 事务性配置（一旦发生异常，自动自毁半成品 Shape，绝不在工作表留下空白废图！）
    try {
      const chart = shape.Chart;
      const srcRange = sheet.Range(targetDataRange);
      chart.SetSourceData(srcRange);

      // 安全设置标题（防范 ChartTitle 空指针崩溃）
      if (title) {
        try {
          chart.HasTitle = true;
          if (chart.ChartTitle) {
            chart.ChartTitle.Text = title;
          }
        } catch (tErr) {
          try {
            if (chart.ChartTitle && chart.ChartTitle.Characters) {
              chart.ChartTitle.Characters.Text = title;
            }
          } catch (tErr2) {
            log("设置图表标题警告: " + tErr2.message);
          }
        }
      }

      if (hasLegend !== undefined) {
        chart.HasLegend = !!hasLegend;
      }

      if (hasDataLabels) {
        try {
          chart.ApplyDataLabels();
        } catch (e) {}
      }

      // 4.1 平滑线设置与系列颜色注入
      try {
        const seriesCol = chart.SeriesCollection();
        const seriesCount = seriesCol ? seriesCol.Count : 0;

        // 全局平滑线设置 (主要针对折线图)
        if (smoothLine) {
          for (let sIdx = 1; sIdx <= seriesCount; sIdx++) {
            try {
              seriesCol.Item(sIdx).Smooth = true;
            } catch (e) {}
          }
        }

        // 系列颜色注入 (seriesColors 数组)
        if (Array.isArray(seriesColors)) {
          // 单系列图表上传多个颜色，宿主会按"逐点染色"处理 → 应单色的柱图/条形图变成彩虹柱
          // （问题台账 ISS-24）。这里只应用第一个颜色，并把"其余被忽略"如实告知。
          if (seriesCount <= 1 && seriesColors.length > 1) {
            warnings.push(
              `seriesColors 传了 ${seriesColors.length} 个颜色，但该图只有 ${seriesCount || 1} 个数据系列；` +
              `已只应用第 1 个（宿主对单系列多色的处理是逐点染色，会出现彩虹柱）。`
            );
          }
          const colorLimit = seriesCount <= 1 ? Math.min(1, seriesColors.length) : seriesColors.length;
          seriesColors.slice(0, colorLimit).forEach((colorHex, idx) => {
            const bgr = hexToExcelColor(colorHex);
            const sIdx = idx + 1;
            if (bgr !== null && sIdx <= seriesCount) {
              try {
                const series = seriesCol.Item(sIdx);
                if (series.Format && series.Format.Line) {
                  series.Format.Line.ForeColor.RGB = bgr;
                } else if (series.Border) {
                  series.Border.Color = bgr;
                }
                if (series.Format && series.Format.Fill) {
                  series.Format.Fill.ForeColor.RGB = bgr;
                } else if (series.Interior) {
                  series.Interior.Color = bgr;
                }
              } catch (e) {
                log(`系列 ${sIdx} 颜色注入警告: ${e.message}`);
              }
            }
          });
        }

        // 高阶单系列独立设置 (seriesSettings)
        const seriesTypeApplied = [];
        const seriesTypeWarnings = [];
        const seriesAxisApplied = [];
        if (Array.isArray(seriesSettings)) {
          seriesSettings.forEach((ss) => {
            const sIdx = Number(ss.seriesIndex);
            if (sIdx >= 1 && sIdx <= seriesCount) {
              try {
                const series = seriesCol.Item(sIdx);
                if (ss.smooth !== undefined) {
                  series.Smooth = !!ss.smooth;
                }
                // 单系列图表类型 → 做**组合图**（柱 + 折线同图）。
                // xlChartType：折线=4 / 柱状簇状=51 / 折线带数据点=65 / 面积=1 / 散点=75
                if (ss.type !== undefined && ss.type !== null && ss.type !== "") {
                  const ST = { line: 4, line_markers: 65, column: 51, column_clustered: 51, area: 1, scatter: 75, bar: 57 };
                  const key = String(ss.type).toLowerCase();
                  const code = ST[key] !== undefined ? ST[key] : (Number.isFinite(Number(ss.type)) ? Number(ss.type) : undefined);
                  if (code === undefined) { seriesTypeWarnings.push(`系列 ${sIdx}: 不认识的 type=${ss.type}`); }
                  else { series.ChartType = code; seriesTypeApplied.push({ seriesIndex: sIdx, type: key, code }); }
                }
                // 副坐标轴：组合图里把折线放到次轴（xlSecondary=2）
                if (ss.axisGroup !== undefined) {
                  series.AxisGroup = Number(ss.axisGroup) === 2 ? 2 : 1;
                  seriesAxisApplied.push({ seriesIndex: sIdx, axisGroup: series.AxisGroup });
                }
                if (ss.color) {
                  const bgr = hexToExcelColor(ss.color);
                  if (bgr !== null) {
                    if (series.Format && series.Format.Line) series.Format.Line.ForeColor.RGB = bgr;
                    else if (series.Border) series.Border.Color = bgr;
                    if (series.Format && series.Format.Fill) series.Format.Fill.ForeColor.RGB = bgr;
                    else if (series.Interior) series.Interior.Color = bgr;
                  }
                }
              } catch (e) {}
            }
          });
        }
      } catch (e) {
        log("系列格式化提示: " + e.message);
      }

      // 4.2 Y 轴数值范围、刻度与数字格式设置 (核心排版防呆)
      if (yAxis) {
        try {
          // xlValue = 2 (数值坐标轴), xlPrimary = 1 (主坐标轴)
          const valAxis = chart.Axes(2, 1);
          if (valAxis) {
            if (yAxis.min !== undefined && yAxis.min !== null) {
              valAxis.MinimumScale = Number(yAxis.min);
            }
            if (yAxis.max !== undefined && yAxis.max !== null) {
              valAxis.MaximumScale = Number(yAxis.max);
            }
            if (yAxis.step !== undefined && yAxis.step !== null) {
              valAxis.MajorUnit = Number(yAxis.step);
            }
            if (yAxis.numberFormat) {
              try {
                valAxis.TickLabels.NumberFormat = yAxis.numberFormat;
              } catch (e) {}
            }
            if (yAxis.title) {
              try {
                valAxis.HasTitle = true;
                valAxis.AxisTitle.Text = yAxis.title;
              } catch (e) {}
            }
          }
        } catch (e) {
          log("数值坐标轴设置提示: " + e.message);
        }
      }

      // 建图后读回真实 ChartType：宿主可能（在任何版本）把请求的类型落成别的类型，
      // 原实现对此毫无察觉（ISS-17）。这里把实际类型带回，不一致就给出 warning。
      let actualChartType = null;
      try { actualChartType = Number(shape.Chart.ChartType); } catch (e) {}
      // 组合图核对：逐系列读回真实 ChartType，确认"折线真的是折线"
      let seriesTypes = null;
      try {
        const sc = shape.Chart.SeriesCollection();
        seriesTypes = [];
        for (let i = 1; i <= Number(sc.Count); i++) { try { seriesTypes.push(Number(sc.Item(i).ChartType)); } catch (e) { seriesTypes.push(null); } }
      } catch (e) {}
      if (actualChartType !== null && Number.isFinite(actualChartType) && actualChartType !== xlChartType) {
        warnings.push(
          `请求的 chartType=${chartType}（xlChartType=${xlChartType}）实际落成 ChartType=${actualChartType}；` +
          `请用 wps_get_charts 复核，必要时改用 wps_execute_script 直接指定常量。`
        );
      }

      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        chartType,
        requestedChartType: xlChartType,
        seriesTypes,
        actualChartType,
        warnings,
        dataRange: targetDataRange,
        title: title || "",
        left,
        top,
        width,
        height,
        shapeName: shape.Name,
        chartIndex: getChartOrdinal(sheet, shape.Name),
        message: `已成功在 [${sheet.Name}] 创建 ${chartType} 原生图表，数据源为 ${targetDataRange}`
      };
    } catch (transactionErr) {
      // 事务回滚：立即自毁残缺半成品 Shape，绝不在工作表留下空白废图！
      try {
        if (shape) shape.Delete();
      } catch (delErr) {}
      throw new Error(`图表创建异常并已自动销毁半成品对象: ${transactionErr.message}`);
    }
  }

  function getChartOrdinal(sheet, shapeName) {
    let ordinal = 0;
    for (let i = 1; i <= sheet.Shapes.Count; i++) {
      const shape = sheet.Shapes.Item(i);
      if (!shape.HasChart) continue;
      ordinal++;
      if (shape.Name === shapeName) return ordinal;
    }
    return null;
  }

  function getCharts(app, params) {
    const { sheetName, workbookName, shapeName, chartIndex, chartTitle, detail = false } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const charts = [];
    let ordinal = 0;
    // ChartType 数字 → 可读枚举名（问题台账 ISS-23：detail 原来只给数字，调用方没法判类型）
    const CHART_TYPE_NAMES = {
      1: "xlArea", 4: "xlLine", 5: "xlPie", 51: "xlColumnClustered", 52: "xlColumnStacked",
      57: "xlBarClustered", 58: "xlBarStacked", 65: "xlLineMarkers", 68: "xlBarOfPie",
      122: "xlPareto", "-4120": "xlDoughnut", "-4169": "xlXYScatter"
    };

    for (let i = 1; i <= sheet.Shapes.Count; i++) {
      const shape = sheet.Shapes.Item(i);
      if (!shape.HasChart) continue;
      ordinal++;
      const chart = shape.Chart;
      const title = safeRead(() => chart.HasTitle && chart.ChartTitle ? chart.ChartTitle.Text : "", "");
      if (shapeName && shape.Name !== shapeName) continue;
      if (chartIndex && ordinal !== Number(chartIndex)) continue;
      if (chartTitle && String(title).indexOf(chartTitle) < 0) continue;

      const position = { left: safeRead(() => shape.Left, null), top: safeRead(() => shape.Top, null) };
      const item = {
        chartIndex: ordinal,
        shapeName: shape.Name,
        title,
        chartType: safeRead(() => chart.ChartType, null),
        chartTypeName: CHART_TYPE_NAMES[safeRead(() => chart.ChartType, null)] || null,
        hasLegend: !!safeRead(() => chart.HasLegend, false),
        left: position.left,
        top: position.top,
        width: safeRead(() => shape.Width, null),
        height: safeRead(() => shape.Height, null),
        // leftCell 是"形状左上角所在单元格"，**不是唯一锚点**：多张图叠在同一像素位时会报同一个值
        // （问题台账 ISS-23）。额外给出 isDefaultPosition，避免被当成唯一定位依据。
        leftCell: safeRead(() => shape.TopLeftCell.Address(), null),
        isDefaultPosition: Math.abs(Number(position.left) - 360) < 2 && Math.abs(Number(position.top) - 40) < 2,
        seriesCount: safeRead(() => chart.SeriesCollection().Count, 0)
      };

      if (detail) {
        const series = [];
        const seriesCollection = safeRead(() => chart.SeriesCollection(), null);
        const seriesCount = seriesCollection ? safeRead(() => seriesCollection.Count, 0) : 0;
        for (let s = 1; s <= seriesCount; s++) {
          const current = seriesCollection.Item(s);
          series.push({
            seriesIndex: s,
            name: safeRead(() => current.Name, null),
            formula: safeRead(() => current.Formula, null),
            smooth: safeRead(() => current.Smooth, null),
            lineColor: excelColorToHex(safeRead(() => current.Format.Line.ForeColor.RGB, null)),
            fillColor: excelColorToHex(safeRead(() => current.Format.Fill.ForeColor.RGB, null))
          });
        }
        item.series = series;
        item.yAxis = safeRead(() => {
          const axis = chart.Axes(2, 1);
          return {
            minimumScale: axis.MinimumScale,
            maximumScale: axis.MaximumScale,
            majorUnit: axis.MajorUnit,
            numberFormat: axis.TickLabels.NumberFormat,
            title: axis.HasTitle && axis.AxisTitle ? axis.AxisTitle.Text : ""
          };
        }, null);
      }
      charts.push(item);
    }

    return {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      totalChartCount: ordinal,
      chartCount: charts.length,
      detail: !!detail,
      charts
    };
  }

  // 15.1 原生矢量图表删除与清理引擎
  function deleteChart(app, params) {
    const {
      sheetName,
      workbookName,
      chartTitle,
      shapeName,
      leftCell,
      chartIndex,
      clearAll = false
    } = params || {};

    const sheet = getWorksheet(app, sheetName, workbookName);
    const shapes = sheet.Shapes;
    const count = shapes.Count;
    const chartOrdinals = {};
    let chartOrdinal = 0;
    for (let i = 1; i <= count; i++) {
      if (shapes.Item(i).HasChart) {
        chartOrdinal++;
        chartOrdinals[i] = chartOrdinal;
      }
    }

    let targetLeft = null;
    let targetTop = null;
    if (leftCell) {
      try {
        const anchor = sheet.Range(leftCell);
        targetLeft = anchor.Left;
        targetTop = anchor.Top;
      } catch (e) {}
    }

    // 先收集候选、再决定删不删。
    // `leftCell` 是按**像素邻近（±30px）**匹配的，多张图叠在同一位置时会**全部命中**——
    // 实测传 leftCell:"H2" 一次删掉了 14 张图（问题台账 ISS-19）。破坏性操作必须先设卡。
    const candidates = [];
    for (let i = 1; i <= count; i++) {
      const shp = shapes.Item(i);
      if (!shp.HasChart) continue;
      let reason = null;
      if (clearAll) {
        reason = "clearAll";
      } else if (shapeName && shp.Name === shapeName) {
        reason = "shapeName";
      } else if (chartIndex && chartOrdinals[i] === Number(chartIndex)) {
        reason = "chartIndex";
      } else if (targetLeft !== null && targetTop !== null) {
        if (Math.abs(shp.Left - targetLeft) < 30 && Math.abs(shp.Top - targetTop) < 30) {
          reason = "leftCell(±30px)";
        }
      } else if (chartTitle) {
        try {
          if (shp.Chart.HasTitle && shp.Chart.ChartTitle.Text.includes(chartTitle)) {
            reason = "chartTitle";
          }
        } catch (e) {}
      }
      if (reason) {
        candidates.push({ index: i, name: shp.Name, reason: reason, left: Math.round(shp.Left), top: Math.round(shp.Top) });
      }
    }

    if (candidates.length === 0) {
      throw new Error(
        `未匹配到任何图表：工作表 [${sheet.Name}] 上没有符合条件的图表。` +
        `请先用 wps_get_charts 读回图表清单（shapeName / 序号 / 标题 / 位置），再指定要删除的那一张。`
      );
    }

    if (!clearAll && candidates[0].reason === "leftCell(±30px)" && candidates.length > 1) {
      throw new Error(
        `leftCell 按像素邻近（±30px）匹配，本次命中 ${candidates.length} 张图表，**已拒绝批量删除**以免误伤。` +
        `命中清单：${candidates.map(c => `${c.name}(左${c.left},上${c.top})`).join("；")}。` +
        `请改用 shapeName 或 chartIndex 精确指定；确认要全删请显式传 clearAll: true。`
      );
    }

    const deletedNames = [];
    // 从后往前删，避免索引位移
    for (let k = candidates.length - 1; k >= 0; k--) {
      const shp = shapes.Item(candidates[k].index);
      deletedNames.push(shp.Name);
      shp.Delete();
    }
    const deletedCount = deletedNames.length;

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      deletedCount,
      deletedNames,
      matchedBy: candidates[0].reason,
      message: `已成功在工作表 [${sheet.Name}] 中删除 ${deletedCount} 张图表（匹配方式：${candidates[0].reason}）`
    };
  }

  // 16. 数据透视表一键生成 (第三梯队)
  function createPivotTable(app, params) {
    // CAP-36 透视表读回：此前只能建、不能读，AI 无法回答
    // "这张表上现有哪些透视表、数据源是哪、字段怎么摆、刷新过没有"。
    if ((params || {}).action === "read") {
      const { workbookName, sheetName, pivotTableName } = params || {};
      const wb = getWorkbook(app, workbookName);
      const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
      const readOne = (pt) => {
        const fields = [];
        try {
          const raw = pt.PivotFields();
          const n = Number(raw.Count);
          for (let i = 1; i <= n; i++) {
            const f = raw.Item(i);
            fields.push({
              name: g(() => String(f.Name), null),
              orientation: g(() => Number(f.Orientation), null),   // 1=行 2=列 4=页 0=隐藏
              position: g(() => Number(f.Position), null),
              function: g(() => Number(f.Function), null),
              subtotals: g(() => Boolean(f.Subtotals), null)
            });
          }
        } catch (e) {}
        return {
          name: g(() => String(pt.Name), null),
          sourceData: g(() => String(pt.SourceData), null),
          rowRange: g(() => String(pt.RowRange.Address()), null),
          tableRange1: g(() => String(pt.TableRange1.Address()), null),
          refreshDate: g(() => (pt.RefreshDate ? new Date(pt.RefreshDate).toISOString() : null), null),
          version: g(() => String(pt.Version), null),
          fieldCount: fields.length,
          fields
        };
      };
      const out = [];
      const sheets = sheetName ? [sheetName] : (() => {
        const names = [];
        for (let i = 1; i <= wb.Worksheets.Count; i++) { try { names.push(String(wb.Worksheets.Item(i).Name)); } catch (e) {} }
        return names;
      })();
      for (const sn of sheets) {
        let pts = null;
        try { pts = wb.Worksheets.Item(sn).PivotTables(); } catch (e) { continue; }
        const n = g(() => Number(pts.Count), 0);
        for (let i = 1; i <= n; i++) {
          const pt = (() => { try { return pts.Item(i); } catch (e) { return null; } })();
          if (!pt) continue;
          const one = readOne(pt);
          if (pivotTableName && one.name !== String(pivotTableName)) continue;
          out.push({ sheetName: sn, ...one });
        }
      }
      return {
        success: true, workbookName: wb.Name,
        count: out.length, pivotTables: out, warnings: [],
        message: `工作簿 [${wb.Name}] 共 ${out.length} 个数据透视表`
      };
    }
    const {
      workbookName,
      sourceSheetName,
      sourceRange,
      destSheetName,
      destCell,
      rowFields = [],
      columnFields = [],
      dataFields = []
    } = params || {};

    if (!sourceRange) throw new Error("缺少必要参数: sourceRange (例如 '明细!A1:K5422')");
    if (!destCell) throw new Error("缺少必要参数: destCell (例如 'B4')");

    const wb = getWorkbook(app, workbookName);
    const srcSheet = getWorksheet(app, sourceSheetName, workbookName);
    // 防御性激活。真机探到 Create(1, srcRange) 在活动表上返回对象、非活动表上返回 null，
    // 随后 .CreatePivotTable 报错；同时源区域必须**有数据**，空区域同样建不出透视表
    // （ISS-118 复核结论：双因，不是单一原因）。
    try { srcSheet.Activate(); } catch (e) {}
    const srcRange = srcSheet.Range(sourceRange);

    // 目标表不存在时自动新建（原来必须由调用方先建好，问题台账 ISS-63 的第二个坑）
    let destSheet = null;
    try {
      destSheet = getWorksheet(app, destSheetName, workbookName);
    } catch (e) {
      if (!destSheetName) throw e;
      destSheet = wb.Worksheets.Add();
      try { destSheet.Name = String(destSheetName); } catch (nameErr) {
        throw new Error(`目标工作表 "${destSheetName}" 不存在，自动新建后重命名失败（可能重名或含非法字符）：${nameErr.message}`);
      }
    }
    try { destSheet.Activate(); } catch (e) {}
    const destRange = destSheet.Range(destCell);

    // 1 = xlDatabase
    const pivotCache = wb.PivotCaches().Create(1, srcRange);
    const ptName = "PivotTable_" + Date.now();
    const pivotTable = pivotCache.CreatePivotTable(destRange, ptName);

    // 行字段: xlRowField = 1
    if (Array.isArray(rowFields)) {
      rowFields.forEach((rf) => {
        try {
          const pf = pivotTable.PivotFields(rf);
          if (pf) pf.Orientation = 1;
        } catch (e) {
          log("设置透视表行维度异常: " + rf + ", " + e.message);
        }
      });
    }

    // 列字段: xlColumnField = 2
    if (Array.isArray(columnFields)) {
      columnFields.forEach((cf) => {
        try {
          const pf = pivotTable.PivotFields(cf);
          if (pf) pf.Orientation = 2;
        } catch (e) {
          log("设置透视表列维度异常: " + cf + ", " + e.message);
        }
      });
    }

    // 汇总字段与函数: xlSum = 4, xlCount = 2, xlAverage = 5, xlMax = 6, xlMin = 7
    const summaryFuncMap = {
      sum: 4,
      count: 2,
      average: 5,
      max: 6,
      min: 7
    };

    if (Array.isArray(dataFields)) {
      // 可用的源字段清单（用于字段名校验，M-7：名字不存在时宿主静默丢字段、工具照样报成功）
      const availableFields = [];
      try {
        const rf = pivotTable.PivotFields();
        for (let i = 1; i <= Number(rf.Count); i++) { try { availableFields.push(String(rf.Item(i).Name)); } catch (e) {} }
      } catch (e) {}
      const droppedFields = [];
      dataFields.forEach((df) => {
        try {
          const wanted = String(df.fieldName);
          if (availableFields.length && availableFields.indexOf(wanted) < 0) {
            droppedFields.push({ fieldName: wanted, reason: "源区域没有该字段" });
            return;
          }
          const pf = pivotTable.PivotFields(wanted);
          const caption = df.caption || (`${df.summaryFunction || "求和"}:${wanted}`);
          const func = summaryFuncMap[df.summaryFunction] || 4;
          pivotTable.AddDataField(pf, caption, func);
        } catch (e) {
          droppedFields.push({ fieldName: String(df.fieldName), reason: String(e.message) });
        }
      });
      if (droppedFields.length) {
        warnings.push(`有 ${droppedFields.length} 个数据字段未加上：${JSON.stringify(droppedFields)}；可用字段：${availableFields.join(" / ") || "（读不到）"}`);
      }
    }

    // 新建后是"空骨架"，字段配好也必须显式刷新才会真正取数（问题台账 ISS-63：
    // 调用方拿到 success 却看到一张空表，得自己去手动 Refresh）。
    let refreshedBy = null;
    try {
      pivotTable.RefreshTable();
      refreshedBy = "pivotTable.RefreshTable";
    } catch (e) {
      try {
        pivotCache.Refresh();
        refreshedBy = "pivotCache.Refresh";
      } catch (e2) {
        log("透视表刷新失败: " + e.message + " / " + e2.message);
      }
    }

    // 读回实际结果：记录数与表区域是"真的取到数"的唯一证据
    let recordCount = null;
    let tableRange = null;
    try { recordCount = Number(pivotTable.RecordCount); } catch (e) {}
    try { tableRange = pivotTable.TableRange1 ? pivotTable.TableRange1.Address() : null; } catch (e) {}
    const warnings = [];
    if (!refreshedBy) warnings.push("透视表刷新调用失败，可能是空骨架，请在 WPS 里右键手动刷新后再读回确认。");
    if (recordCount === 0) warnings.push("刷新后 RecordCount 仍为 0：源区域可能没有可用数据，或字段未正确落位。");

    return {
      success: true,
      workbookName: wb.Name,
      pivotTableName: ptName,
      destSheetName: destSheet.Name,
      destCell,
      rowCount: rowFields.length,
      colCount: columnFields.length,
      dataCount: dataFields.length,
      refreshedBy,
      recordCount,
      tableRange,
      warnings,
      message: `已成功在 [${destSheet.Name}] ${destCell} 生成数据透视表${refreshedBy ? "并完成刷新" : "（刷新未成功，见 warnings）"}`
    };
  }

  // 17. 自动筛选与数据排序 (第三梯队)
  function setFilterAndSort(app, params) {
    const { sheetName, workbookName, range, enableAutoFilter, sortRules, action = "apply" } = params || {};

    const sheet = getWorksheet(app, sheetName, workbookName);

    // CAP-33 筛选状态读回：不传 range 也能读（读的是"这张表当前有没有筛选、覆盖哪一块"）。
    if (action === "read") {
      const active = (() => { try { return Boolean(sheet.AutoFilterMode); } catch (e) { return null; } })();
      const addr = (() => {
        try { return sheet.AutoFilterMode && sheet.AutoFilter && sheet.AutoFilter.Range ? sheet.AutoFilter.Range.Address() : null; } catch (e) { return null; }
      })();
      // 逐个字段的筛选条件（Criteria1/Criteria2/Operator）
      const filters = [];
      try {
        if (active && addr) {
          const f = sheet.AutoFilter;
          const rng = f.Range;
          const cols = rng.Columns.Count;
          for (let i = 1; i <= cols; i++) {
            const flt = f.Filters.Item(i);
            let on = false;
            try { on = Boolean(flt.On); } catch (e) {}
            if (!on) continue;
            const one = (() => { try { return flt.Criteria1 === undefined ? null : flt.Criteria1; } catch (e) { return null; } })();
            const two = (() => { try { return flt.Criteria2 === undefined ? null : flt.Criteria2; } catch (e) { return null; } })();
            const op = (() => { try { return Number(flt.Operator); } catch (e) { return null; } })();
            const header = (() => { try { return String(rng.Cells(1, i).Value2); } catch (e) { return null; } })();
            filters.push({ columnIndex: i, header, criteria1: one, criteria2: two, operator: op });
          }
        }
      } catch (e) {}
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        autoFilterOn: active,
        filterRange: addr,
        columnCount: addr ? sheet.Range(addr).Columns.Count : 0,
        filters,
        warnings: [],
        message: active
          ? `工作表 [${sheet.Name}] 筛选范围 ${addr}，其中 ${filters.length} 列设有条件`
          : `工作表 [${sheet.Name}] 当前没有启用筛选`
      };
    }

    if (!range) throw new Error("缺少必要参数: range (例如 'A4:E20')");

    // 防御性激活：WPS 上部分 Range 级操作（AutoFilter / Sort）对非活动表可能静默 no-op。
    // 说明：曾把「筛选不生效」记成产品缺陷（ISS-117），复核后确认是**验证脚本没检查写入结果**——
    // 数据根本没写进表，空区域上 AutoFilter() 自然无效。激活保留作防御，但不是该问题的根因。
    try { sheet.Activate(); } catch (e) {}

    const targetRange = sheet.Range(range);
    const warnings = [];

    // 自动筛选控制
    let appliedFilterRange = null;
    if (enableAutoFilter !== undefined) {
      if (enableAutoFilter) {
        if (!sheet.AutoFilterMode) {
          targetRange.AutoFilter();
        }
      } else {
        if (sheet.AutoFilterMode) {
          sheet.AutoFilterMode = false;
        }
      }
      // 回读筛选实际覆盖范围：宿主会把筛选自动扩展到相邻的整块数据区，
      // 传入的 range 只是锚点而不是约束（问题台账 ISS-40）。这里把真实范围报出来，不再让调用方以为是自己传的那个。
      try {
        appliedFilterRange = sheet.AutoFilterMode && sheet.AutoFilter && sheet.AutoFilter.Range
          ? sheet.AutoFilter.Range.Address()
          : null;
      } catch (e) {
        appliedFilterRange = null;
      }
      if (appliedFilterRange && appliedFilterRange !== targetRange.Address()) {
        warnings.push(`筛选实际覆盖 ${appliedFilterRange}，与传入的 ${range} 不一致：宿主会把筛选扩展到相邻数据块。若需精确范围，请在目标区与其它数据之间留一个空行。`);
      }
      // **读回核对**：写入"不报错"≠"真的生效"。
      // 真机踩到：目标区域若已是**结构化表格**（ListObject），它自带筛选器，
      // 此时 `Range.AutoFilter()` 不生效、`sheet.AutoFilterMode` 仍为 false，
      // 但工具此前返回 success —— 调用方会以为筛选已开。
      if (enableAutoFilter && !appliedFilterRange) {
        let tableNames = [];
        try {
          const los = sheet.ListObjects;
          for (let i = 1; i <= Number(los.Count); i++) { try { tableNames.push(String(los.Item(i).Name)); } catch (e) {} }
        } catch (e) {}
        if (tableNames.length) {
          warnings.push(`筛选**未生效**：工作表上存在结构化表格 ${tableNames.join(" / ")}，它自带筛选器，表级 AutoFilter 不会另外打开。请直接在表格上操作，或先把表格转为普通区域。`);
        } else {
          warnings.push("筛选**未生效**：调用后 sheet.AutoFilterMode 仍为 false，请人工确认。");
        }
      }
    }

    // 数据排序：旧式 Range.Sort(...) 在本机 WPS 上会静默 no-op（问题台账 ISS-38），
    // 因此先走 SortFields，再读回校验，都无效时**报错而不是返回假成功**。
    let sortApplied = null;
    if (Array.isArray(sortRules) && sortRules.length > 0) {
      const rowCount = targetRange.Rows.Count;
      const colCount = targetRange.Columns.Count;
      // colIndex 越界时宿主静默 no-op（excel-tester M-6）——先做边界校验
      for (const rule of sortRules) {
        const ci = Number(rule && rule.colIndex);
        if (!Number.isFinite(ci) || ci < 1 || ci > colCount) {
          throw new Error(`sortRules.colIndex=${rule && rule.colIndex} 越界：目标区域 ${range} 只有 ${colCount} 列（1..${colCount}）`);
        }
      }
      const before = normalize2DArray(targetRange.Value2, rowCount, colCount);
      const attempts = [];
      let after = before;
      let sorted = false;

      const readCurrent = () => normalize2DArray(targetRange.Value2, rowCount, colCount);
      const sortKeyOf = (rule) => targetRange.Columns.Item(Number(rule.colIndex));
      const orderOf = (rule) => (rule.order === "desc" ? 2 : 1);

      // 本机 WPS 实测：`Range.Sort` 是**方法**而不是对象（访问 .SortFields 抛
      // "Cannot read properties of undefined (reading 'Clear')"），而 `sheet.Sort` 才是可用的 Sort 对象。
      // 因此按下面顺序逐条尝试，每一条都**读回校验**，全都不生效就报错（问题台账 ISS-38）。
      const sortPaths = [
        {
          name: "sheet.Sort.SortFields",
          run: () => {
            const s = sheet.Sort;
            s.SortFields.Clear();
            for (let i = 0; i < sortRules.length; i++) s.SortFields.Add(sortKeyOf(sortRules[i]), 0, orderOf(sortRules[i]));
            s.SetRange(targetRange);
            s.Header = 1; // 包含表头
            s.Apply();
          }
        },
        {
          name: "range.Sort.SortFields",
          run: () => {
            const s = targetRange.Sort;
            s.SortFields.Clear();
            for (let i = 0; i < sortRules.length; i++) s.SortFields.Add(sortKeyOf(sortRules[i]), 0, orderOf(sortRules[i]));
            s.SetRange(targetRange);
            s.Header = 1;
            s.Apply();
          }
        },
        {
          name: "range.Sort(旧式)",
          run: () => {
            const key1 = sortKeyOf(sortRules[0]), order1 = orderOf(sortRules[0]);
            let key2, order2;
            if (sortRules[1]) {
              key2 = sortKeyOf(sortRules[1]);
              order2 = orderOf(sortRules[1]);
            }
            targetRange.Sort(key1, order1, key2, order2, undefined, undefined, 1);
          }
        }
      ];

      for (let p = 0; p < sortPaths.length; p++) {
        const path = sortPaths[p];
        let error = null;
        try {
          path.run();
        } catch (e) {
          error = e.message;
        }
        // 即使抛错也读回：某条路径可能已经部分生效
        after = readCurrent();
        sorted = isSortedByRules(after, sortRules);
        attempts.push(path.name + (error ? ":异常(" + error + ")" : ":ok") + (sorted ? ":已生效" : ":未生效"));
        if (sorted) break;
      }

      if (!sorted) {
        throw new Error(
          `排序未生效：区域 ${range} 读回后的顺序不满足请求的排序规则。已尝试：${attempts.join("；")}。` +
          `请核对 colIndex 是否为**区域内相对列号**（1 = 区域第一列），或改用 wps_execute_script 的 sheet.Sort.SortFields 路径。`
        );
      }
      sortApplied = { changed: JSON.stringify(before) !== JSON.stringify(after), attempts };
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      range,
      enableAutoFilter: enableAutoFilter !== undefined ? !!enableAutoFilter : sheet.AutoFilterMode,
      appliedFilterRange,
      sortedRuleCount: Array.isArray(sortRules) ? sortRules.length : 0,
      sortApplied,
      warnings,
      message: `已成功在 [${sheet.Name}] ${range} 应用筛选与排序`
    };
  }

  // 18. 单元格下拉验证菜单 (第三梯队)
  function setDataValidation(app, params) {
    const {
      sheetName,
      workbookName,
      address,
      validationType = "list",
      listItems = [],
      operator = "between",
      minVal,
      maxVal,
      promptTitle,
      promptMessage,
      errorTitle,
      errorMessage
    } = params || {};

    const sheet = getWorksheet(app, sheetName, workbookName);

    // CAP-32 数据有效性读回：读整张表**所有**带校验的区域，或指定 address 的那一块。
    // 只写不读时 AI 无法回答"这列现在允许哪些值"，复核只能靠人看。
    if ((params || {}).action === "read") {
      const VALTYPE = { 1: "whole_number", 2: "decimal", 3: "list", 4: "date", 5: "time", 6: "text_length", 7: "custom" };
      const OP = { 1: "between", 2: "not_between", 3: "equal", 4: "not_equal", 5: "greater_than", 6: "less_than", 7: "greater_equal", 8: "less_equal" };
      const readOne = (rng) => {
        const v = rng.Validation;
        let type = null;
        try { type = Number(v.Type); } catch (e) { return null; }
        // Type = -4142 (xlValidateInputOnly) 表示没设校验，跳过
        if (type === -4142 || !Number.isFinite(type)) return null;
        const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
        return {
          address: g(() => rng.Address(), null),
          type: VALTYPE[type] || `unknown(${type})`,
          typeCode: type,
          operator: OP[g(() => Number(v.Operator), null)] || null,
          formula1: g(() => v.Formula1, null),
          formula2: g(() => v.Formula2, null),
          inCellDropdown: g(() => Boolean(v.InCellDropdown), null),
          ignoreBlank: g(() => Boolean(v.IgnoreBlank), null),
          showError: g(() => Boolean(v.ShowError), null),
          errorTitle: g(() => v.ErrorTitle, null),
          errorMessage: g(() => v.ErrorMessage, null),
          promptTitle: g(() => v.InputTitle, null),
          promptMessage: g(() => v.InputMessage, null)
        };
      };
      let items = [];
      if (address) {
        const one = readOne(sheet.Range(address));
        items = one ? [one] : [];
      } else {
        // 逐行扫描已用区域（宿主没有"枚举所有校验区域"的 API）
        const ur = sheet.UsedRange;
        const r0 = ur.Row, c0 = ur.Column;
        const rows = Math.min(ur.Rows.Count, 2000), cols = Math.min(ur.Columns.Count, 100);
        const colName = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
        for (let r = r0; r < r0 + rows; r++) {
          for (let c = c0; c < c0 + cols; c++) {
            const one = readOne(sheet.Range(`${colName(c)}${r}`));
            if (one) items.push(one);
          }
        }
      }
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        count: items.length,
        validations: items,
        warnings: address ? [] : [],
        message: `工作表 [${sheet.Name}] 上共 ${items.length} 处数据有效性设置`
      };
    }

    if (!address) throw new Error("缺少必要参数: address (例如 'E5:E20')");

    const targetRange = sheet.Range(address);

    try {
      targetRange.Validation.Delete();
    } catch (e) {}

    if (validationType === "list") {
      const listStr = Array.isArray(listItems) ? listItems.join(",") : String(listItems || "");
      // 缺候选项时宿主会静默 no-op（excel-tester M-5）——直接报错而不是假装设上了
      if (listStr.trim() === "") throw new Error("validationType='list' 必须提供 listItems（候选项数组），否则宿主会静默不设任何校验");
      // Type: 3 (xlValidateList), AlertStyle: 1 (xlValidAlertStop), Operator: 1 (xlBetween)
      targetRange.Validation.Add(3, 1, 1, listStr);
      targetRange.Validation.InCellDropdown = true;
    } else if (validationType === "number_range") {
      let op = 1; // xlBetween
      if (operator === "greater_than") op = 5;
      else if (operator === "less_than") op = 6;
      else if (operator === "equal") op = 3;

      const f1 = minVal !== undefined ? String(minVal) : "0";
      const f2 = maxVal !== undefined ? String(maxVal) : undefined;
      // Type: 2 (xlValidateDecimal)
      targetRange.Validation.Add(2, 1, op, f1, f2);
    } else {
      throw new Error(`未知的 validationType: ${validationType} (支持 list, number_range)`);
    }

    if (promptMessage) {
      targetRange.Validation.InputTitle = promptTitle || "选择提示";
      targetRange.Validation.InputMessage = promptMessage;
      targetRange.Validation.ShowInput = true;
    }

    if (errorMessage) {
      targetRange.Validation.ErrorTitle = errorTitle || "输入无效";
      targetRange.Validation.ErrorMessage = errorMessage;
      targetRange.Validation.ShowError = true;
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address,
      validationType,
      message: `已成功在 [${sheet.Name}] ${address} 配置数据有效性验证`
    };
  }

  // 19. 工作表综合管理 (第四梯队)
  function manageSheet(app, params) {
    const { sheetName, workbookName, action, newName, targetIndex, color, password } = params || {};
    if (!sheetName) throw new Error("缺少必要参数: sheetName");
    if (!action) throw new Error("缺少必要参数: action (rename, move, tab_color, protect, unprotect, read)");

    const wb = getWorkbook(app, workbookName);
    const sheet = getWorksheet(app, sheetName, workbookName);

    // CAP-34 工作表状态读回：保护状态与标签色此前只能写不能读，
    // 于是"这张表是不是被保护了""标签什么颜色"只能靠人看。
    if (action === "read") {
      const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
      const prot = g(() => Boolean(sheet.ProtectContents), null);
      const colorRaw = g(() => Number(sheet.Tab.Color), null);
      // Tab.Color 是 BGR 整数；-4142/16777215 之类表示"无颜色"
      const hasColor = colorRaw !== null && colorRaw >= 0 && colorRaw !== 16777215;
      const toHex = (v) => "#" + [v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff].map(n => n.toString(16).padStart(2, "0")).join("").toUpperCase();
      return {
        success: true,
        workbookName: wb.Name,
        sheetName: sheet.Name,
        index: g(() => Number(sheet.Index), null),
        visible: g(() => Number(sheet.Visible), null),
        protection: {
          protectContents: prot,
          protectDrawingObjects: g(() => Boolean(sheet.ProtectDrawingObjects), null),
          protectionMode: g(() => Boolean(sheet.ProtectionMode), null)
        },
        tabColor: hasColor ? toHex(colorRaw) : null,
        tabColorRaw: colorRaw,
        warnings: [],
        message: `工作表 [${sheet.Name}]：${prot ? "已保护" : "未保护"}，标签色 ${hasColor ? toHex(colorRaw) : "未设置（默认）"}`
      };
    }

    switch (action) {
      case "rename": {
        if (!newName) throw new Error("rename 操作必须提供 newName");
        const oldName = sheet.Name;
        // 目标名已存在时宿主会静默失败或改名成别的——先查存在性（excel-tester M-1）
        const exists = (() => {
          try {
            for (let i = 1; i <= Number(wb.Worksheets.Count); i++) {
              if (String(wb.Worksheets.Item(i).Name) === String(newName)) return true;
            }
          } catch (e) {}
          return false;
        })();
        if (exists) throw new Error(`工作表名 "${newName}" 已被占用，改名未执行（原表仍是 "${oldName}"）`);
        sheet.Name = newName;
        // 读回核对：宿主改名可能静默不生效（excel-tester M-1）
        const actual = (() => { try { return String(sheet.Name); } catch (e) { return null; } })();
        if (actual !== String(newName)) {
          return { success: false, workbookName: wb.Name, oldName, newName, actualName: actual, warnings: [`改名未生效：请求 "${newName}"，宿主读回 "${actual}"`], message: `工作表改名未生效` };
        }
        return { success: true, workbookName: wb.Name, oldName, newName, actualName: actual, warnings: [], message: `工作表 [${oldName}] 已重命名为 [${newName}]` };
      }
      case "move": {
        if (!targetIndex) throw new Error("move 操作必须提供 targetIndex (1-indexed)");
        const total = wb.Worksheets.Count;
        const validIndex = Math.max(1, Math.min(Number(targetIndex), total));
        const refSheet = wb.Worksheets.Item(validIndex);
        sheet.Move(refSheet);
        return { success: true, workbookName: wb.Name, sheetName: sheet.Name, targetIndex: validIndex, message: `工作表 [${sheet.Name}] 已成功移动至位置 ${validIndex}` };
      }
      case "tab_color": {
        if (!color) throw new Error("tab_color 操作必须提供十六进制颜色 (例如 '#0F172A')");
        const bgr = hexToExcelColor(color);
        if (bgr !== null) {
          sheet.Tab.Color = bgr;
        }
        return { success: true, workbookName: wb.Name, sheetName: sheet.Name, tabColor: color, message: `工作表 [${sheet.Name}] 标签底色已设置为 ${color}` };
      }
      case "protect": {
        sheet.Protect(password || undefined);
        return { success: true, workbookName: wb.Name, sheetName: sheet.Name, protected: true, message: `工作表 [${sheet.Name}] 已锁定保护` };
      }
      case "unprotect": {
        // ⚠️ 密码错误时宿主**不抛异常也不解锁**，原实现照样返回 protected:false —— 谎报成功
        // （excel-tester M-2）。改为读回 ProtectContents 核对。
        try { sheet.Unprotect(password === undefined ? undefined : String(password)); }
        catch (e) { return { success: false, workbookName: wb.Name, sheetName: sheet.Name, hostError: e.message, warnings: [`解除保护失败: ${e.message}`], message: `工作表 [${sheet.Name}] 解除保护失败` }; }
        let stillProtected = null;
        try { stillProtected = Boolean(sheet.ProtectContents); } catch (e) {}
        if (stillProtected === true) {
          return { success: false, workbookName: wb.Name, sheetName: sheet.Name, protected: true,
            warnings: [password === undefined ? "工作表仍处于保护状态：可能设有密码，请传入 password" : "工作表仍处于保护状态：密码可能不正确"],
            message: `工作表 [${sheet.Name}] 仍处于保护状态（未解除）` };
        }
        return { success: true, workbookName: wb.Name, sheetName: sheet.Name, protected: false, warnings: [], message: `工作表 [${sheet.Name}] 已解除锁定保护` };
      }
      default:
        throw new Error(`未知的 Sheet 操作: ${action} (支持 rename, move, tab_color, protect, unprotect)`);
    }
  }
