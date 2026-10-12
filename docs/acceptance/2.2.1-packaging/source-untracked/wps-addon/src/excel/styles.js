  // styles.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


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
    const allowed = ["fontName", "fontSize", "bold", "fontColor", "backgroundColor", "numberFormat", "horizontalAlignment", "verticalAlignment", "wrapText", "rowHeight", "columnWidth", "merged", "mergeArea", "borders", "validation", "validationViolations"];
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
      const result = Object.assign(base, { styles, mixedOrUnavailableFields });
      // CAP-10：把 include 里的 validationViolations 当作**区域级附加读取**——
      // 逐格比对规则，列出越界单元格（值 + 命中的规则 + 期望）。
      if (include.indexOf("validationViolations") >= 0) {
        result.validationCheck = collectValidationViolations(sheet, range, { maxCells: config.maxCells });
      }
      return result;
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
    const result = Object.assign(base, { totalCells, returnedCells: cells.length, truncated: totalCells > cells.length, cells });
    // CAP-10：cells 模式同样可以顺带做违规定位（区域级结果，不按格重复）
    if (include.indexOf("validationViolations") >= 0) {
      result.validationCheck = collectValidationViolations(sheet, range, { maxCells: config.maxCells });
    }
    return result;
  }
