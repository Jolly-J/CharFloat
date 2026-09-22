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
      headerPreview: headerPreview
    };
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
    const allowed = ["fontName", "fontSize", "bold", "fontColor", "backgroundColor", "numberFormat", "horizontalAlignment", "verticalAlignment", "wrapText", "rowHeight", "columnWidth", "merged", "mergeArea", "borders"];
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
    if (backgroundColor) {
      const bgr = hexToExcelColor(backgroundColor);
      if (bgr !== null) range.Interior.Color = bgr;
    }
    if (fontColor) {
      const bgr = hexToExcelColor(fontColor);
      if (bgr !== null) range.Font.Color = bgr;
    }

    // 3. 对齐
    if (horizontalAlignment) {
      if (horizontalAlignment === "center") range.HorizontalAlignment = -4108; // xlCenter
      else if (horizontalAlignment === "left") range.HorizontalAlignment = -4131; // xlLeft
      else if (horizontalAlignment === "right") range.HorizontalAlignment = -4152; // xlRight
    }
    if (verticalAlignment === "center") {
      range.VerticalAlignment = -4108; // xlCenter
    }

    // 4. 换行与行高
    if (wrapText !== undefined) {
      range.WrapText = !!wrapText;
    }
    if (rowHeight) {
      range.RowHeight = rowHeight;
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
      success: true,
      address: range.Address(),
      appliedStyles: params
    };
  }

  // 10.1 条件格式与数据条/色阶
  function addConditionalFormatting(app, params) {
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
      clearExisting = false
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
    } else {
      throw new Error(`未知的条件格式类型: ${ruleType} (支持 cell_value, data_bar, color_scale)`);
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      ruleType: ruleType,
      message: `已成功在 ${range.Address()} 应用 ${ruleType} 条件格式`
    };
  }

  // 10.2 冻结窗格吸顶
  function freezePanes(app, params) {
    const { sheetName, workbookName, freezeRowIndex, freezeColumnIndex, unfreeze } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}

    const win = app.ActiveWindow;
    if (!win) throw new Error("无法获取 WPS 活动窗口句柄");

    if (unfreeze) {
      win.FreezePanes = false;
      win.SplitRow = 0;
      win.SplitColumn = 0;
      return { success: true, message: "已成功解除当前工作表的窗格冻结" };
    }

    win.FreezePanes = false;
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

    if (action === "add") {
      if (!text) throw new Error("添加批注时必须提供 text 参数");
      const cell = targetRange.Cells.Item(1, 1);
      try {
        if (cell.Comment) cell.Comment.Delete();
      } catch (e) {}
      const commentContent = author ? `${author}:\n${text}` : text;
      const comment = cell.AddComment(commentContent);
      comment.Visible = false;
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        address: cell.Address,
        action: "add",
        commentText: text,
        author: author || "AI 智能审核",
        message: `已在单元格 ${cell.Address} 成功添加批注`
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
                address: cell.Address,
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

    return {
      success: true,
      workbookName: wb.Name,
      sourceSheetName: srcSheet.Name,
      newSheetName: newSheet.Name,
      totalSheets: wb.Worksheets.Count,
      message: `已成功将工作表 [${sourceSheetName}] 完整克隆为 [${newSheetName}]（包含所有格式、公式与图表）`
    };
  }

  function saveWorkbook(app, params) {
    const { workbookName } = params || {};
    const wb = getWorkbook(app, workbookName);
    wb.Save();
    return {
      success: true,
      workbookName: wb.Name,
      fullName: wb.FullName,
      message: `工作簿 [${wb.Name}] 已成功保存到磁盘`
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
      targetRange = sheet.UsedRange;
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
      replaceExisting = true
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
          seriesColors.forEach((colorHex, idx) => {
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
        if (Array.isArray(seriesSettings)) {
          seriesSettings.forEach((ss) => {
            const sIdx = Number(ss.seriesIndex);
            if (sIdx >= 1 && sIdx <= seriesCount) {
              try {
                const series = seriesCol.Item(sIdx);
                if (ss.smooth !== undefined) {
                  series.Smooth = !!ss.smooth;
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
      const typeWarnings = [];
      if (actualChartType !== null && Number.isFinite(actualChartType) && actualChartType !== xlChartType) {
        typeWarnings.push(
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
        actualChartType,
        warnings: typeWarnings,
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

    for (let i = 1; i <= sheet.Shapes.Count; i++) {
      const shape = sheet.Shapes.Item(i);
      if (!shape.HasChart) continue;
      ordinal++;
      const chart = shape.Chart;
      const title = safeRead(() => chart.HasTitle && chart.ChartTitle ? chart.ChartTitle.Text : "", "");
      if (shapeName && shape.Name !== shapeName) continue;
      if (chartIndex && ordinal !== Number(chartIndex)) continue;
      if (chartTitle && String(title).indexOf(chartTitle) < 0) continue;

      const item = {
        chartIndex: ordinal,
        shapeName: shape.Name,
        title,
        chartType: safeRead(() => chart.ChartType, null),
        hasLegend: !!safeRead(() => chart.HasLegend, false),
        left: safeRead(() => shape.Left, null),
        top: safeRead(() => shape.Top, null),
        width: safeRead(() => shape.Width, null),
        height: safeRead(() => shape.Height, null),
        leftCell: safeRead(() => shape.TopLeftCell.Address(), null),
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
    const srcRange = srcSheet.Range(sourceRange);

    const destSheet = getWorksheet(app, destSheetName, workbookName);
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
      dataFields.forEach((df) => {
        try {
          const pf = pivotTable.PivotFields(df.fieldName);
          const caption = df.caption || (`${df.summaryFunction || "求和"}:${df.fieldName}`);
          const func = summaryFuncMap[df.summaryFunction] || 4;
          pivotTable.AddDataField(pf, caption, func);
        } catch (e) {
          log("设置透视表数据字段异常: " + JSON.stringify(df) + ", " + e.message);
        }
      });
    }

    return {
      success: true,
      workbookName: wb.Name,
      pivotTableName: ptName,
      destSheetName: destSheet.Name,
      destCell,
      rowCount: rowFields.length,
      colCount: columnFields.length,
      dataCount: dataFields.length,
      message: `已成功在 [${destSheet.Name}] ${destCell} 生成数据透视表`
    };
  }

  // 17. 自动筛选与数据排序 (第三梯队)
  function setFilterAndSort(app, params) {
    const { sheetName, workbookName, range, enableAutoFilter, sortRules } = params || {};
    if (!range) throw new Error("缺少必要参数: range (例如 'A4:E20')");

    const sheet = getWorksheet(app, sheetName, workbookName);
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
    }

    // 数据排序：旧式 Range.Sort(...) 在本机 WPS 上会静默 no-op（问题台账 ISS-38），
    // 因此先走 SortFields，再读回校验，都无效时**报错而不是返回假成功**。
    let sortApplied = null;
    if (Array.isArray(sortRules) && sortRules.length > 0) {
      const rowCount = targetRange.Rows.Count;
      const colCount = targetRange.Columns.Count;
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

    if (!address) throw new Error("缺少必要参数: address (例如 'E5:E20')");

    const sheet = getWorksheet(app, sheetName, workbookName);
    const targetRange = sheet.Range(address);

    try {
      targetRange.Validation.Delete();
    } catch (e) {}

    if (validationType === "list") {
      const listStr = Array.isArray(listItems) ? listItems.join(",") : String(listItems || "");
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
    if (!action) throw new Error("缺少必要参数: action (rename, move, tab_color, protect, unprotect)");

    const wb = getWorkbook(app, workbookName);
    const sheet = getWorksheet(app, sheetName, workbookName);

    switch (action) {
      case "rename": {
        if (!newName) throw new Error("rename 操作必须提供 newName");
        const oldName = sheet.Name;
        sheet.Name = newName;
        return { success: true, workbookName: wb.Name, oldName, newName, message: `工作表 [${oldName}] 已成功重命名为 [${newName}]` };
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
        sheet.Unprotect(password || undefined);
        return { success: true, workbookName: wb.Name, sheetName: sheet.Name, protected: false, message: `工作表 [${sheet.Name}] 已解除锁定保护` };
      }
      default:
        throw new Error(`未知的 Sheet 操作: ${action} (支持 rename, move, tab_color, protect, unprotect)`);
    }
  }
