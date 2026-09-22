// ── 模块: src/excel/format.js — 格式排版、冻结窗格与条件格式 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 937-1136 行（原样搬迁，未改写）。
  // 5. 格式与排版
  async function handleFormatRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const targetAddress = params.address || params.range || params.cellRange;
      if (!targetAddress) throw new Error("缺少必要参数: address 或 range");
      const range = sheet.getRange(targetAddress);

      // 0. 企业级命名样式模板 (Named Style Templates)
      const styleTemplates = {
        "OPPO_Banner": { fillColor: "#046A38", fontColor: "#FFFFFF", bold: true, fontSize: 14, horizontalAlignment: "Center" },
        "OPPO_Subtitle": { fillColor: "#046A38", fontColor: "#E2ECE6", fontSize: 10, horizontalAlignment: "Center" },
        "OPPO_Card_Header": { fillColor: "#046A38", fontColor: "#FFFFFF", bold: true, fontSize: 11 },
        "OPPO_Table_Header": { fillColor: "#EAF6EE", fontColor: "#046A38", bold: true, fontSize: 10 },
        "OPPO_Total_Row": { fillColor: "#EAF6EE", fontColor: "#046A38", bold: true },
        "OPPO_KPI_Value": { fillColor: "#FFFFFF", fontColor: "#046A38", bold: true, fontSize: 18, horizontalAlignment: "Center" },
        "OPPO_KPI_Sub": { fillColor: "#FFFFFF", fontColor: "#10B981", fontSize: 9, horizontalAlignment: "Center" },
        "OPPO_Canvas_Bg": { fillColor: "#F4F6F8" },
        "OPPO_Card_Body": { fillColor: "#FFFFFF" }
      };

      const preset = params.style ? styleTemplates[params.style] : null;
      if (preset) {
        if (preset.fillColor) range.format.fill.color = preset.fillColor;
        if (preset.fontColor) range.format.font.color = preset.fontColor;
        if (preset.bold !== undefined) range.format.font.bold = preset.bold;
        if (preset.fontSize) range.format.font.size = preset.fontSize;
        if (preset.horizontalAlignment) range.format.horizontalAlignment = preset.horizontalAlignment;
      }

      // 1. 字体样式 (兼容扁平入参与嵌套入参)
      const fontName = params.fontName || params.font?.name;
      if (fontName) range.format.font.name = fontName;

      const fontSize = params.fontSize || params.size || params.font?.size;
      if (fontSize) range.format.font.size = fontSize;

      const fontColor = params.fontColor || params.color || params.textColor || params.font?.color;
      if (fontColor) range.format.font.color = fontColor;

      const bold = params.bold !== undefined ? params.bold : params.font?.bold;
      if (bold !== undefined) range.format.font.bold = bold;

      const italic = params.italic !== undefined ? params.italic : params.font?.italic;
      if (italic !== undefined) range.format.font.italic = italic;

      // 2. 单元格背景填充色 (兼容 backgroundColor, fillColor, bg, fill.color)
      const bgColor = params.backgroundColor || params.fillColor || params.bg || params.fill?.color;
      if (bgColor) {
        range.format.fill.color = bgColor;
      }

      // 3. 对齐方式
      const hAlign = params.horizontalAlignment || params.alignment?.horizontal;
      if (hAlign) {
        const hMap = {
          left: "Left", center: "Center", right: "Right",
          justify: "Justify", general: "General"
        };
        range.format.horizontalAlignment = hMap[String(hAlign).toLowerCase()] || hAlign;
      }

      const vAlign = params.verticalAlignment || params.alignment?.vertical;
      if (vAlign) {
        const vMap = {
          top: "Top", center: "Center", bottom: "Bottom",
          distributed: "Distributed", justify: "Justify"
        };
        range.format.verticalAlignment = vMap[String(vAlign).toLowerCase()] || vAlign;
      }

      const wrapText = params.wrapText !== undefined ? params.wrapText : params.alignment?.wrapText;
      if (wrapText !== undefined) {
        range.format.wrapText = wrapText;
      }

      // 4. 数字/百分比/千分位/货币/日期格式
      if (params.numberFormat) {
        range.numberFormat = [[params.numberFormat]];
      }

      // 5. 边框 (borders)
      if (params.borders) {
        const borderColor = params.borderColor || "#D9D9D9";
        const borderStyle = typeof params.borders === 'string' && ['Continuous', 'Dash', 'Dot', 'Double'].includes(params.borders)
          ? params.borders : 'Continuous';
        const edges = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideHorizontal', 'InsideVertical'];
        for (const edge of edges) {
          try {
            const b = range.format.borders.getItem(edge);
            b.style = borderStyle;
            b.color = borderColor;
          } catch {}
        }
      }

      // 6. 行高与列宽
      if (params.rowHeight) range.rowHeight = params.rowHeight;
      if (params.columnWidth) range.columnWidth = params.columnWidth;

      // 7. 合并 / 拆分
      if (params.merge) range.merge(params.across || false);
      if (params.unmerge) range.unmerge();

      range.load("address");
      sheet.load("name");
      await context.sync();
      return { success: true, address: range.address || params.address, sheetName: sheet.name };
    });
  }

  async function handleAutofitColumns(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params?.sheetName);
      const range = params?.address ? sheet.getRange(params.address) : sheet.getUsedRange();
      range.format.autofitColumns();
      if (params?.autofitRows) range.format.autofitRows();
      await context.sync();
      return { success: true };
    });
  }

  async function handleFreezePanes(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params?.sheetName);
      if (params?.unfreeze) {
        sheet.freezePanes.unfreeze();
        await context.sync();
        return { success: true, message: "已解除冻结窗格" };
      }

      // 避免 Office.js 重复冻结报 InternalError: An internal error has occurred，先重置解冻
      sheet.freezePanes.unfreeze();

      const r = Number(params?.freezeRowIndex || 0);
      const c = Number(params?.freezeColumnIndex || 0);
      if (r > 0 && c > 0) {
        sheet.freezePanes.freezeAt(sheet.getCell(r, c));
      } else if (r > 0) {
        sheet.freezePanes.freezeRows(r);
      } else if (c > 0) {
        sheet.freezePanes.freezeColumns(c);
      } else {
        sheet.freezePanes.freezeRows(1);
      }
      await context.sync();
      return { success: true, message: `已冻结窗格 (行: ${r}, 列: ${c})` };
    });
  }

  async function handleAddConditionalFormatting(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);

      if (params.clearExisting) {
        range.conditionalFormats.clearAll();
      }

      const ruleType = params.ruleType || params.type || "cell_value";
      if (ruleType === "cell_value" || ruleType === "cellValue") {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        const opMap = {
          less_than: Excel.ConditionalCellValueOperator.lessThan,
          greater_than: Excel.ConditionalCellValueOperator.greaterThan,
          equal: Excel.ConditionalCellValueOperator.equalTo,
          not_equal: Excel.ConditionalCellValueOperator.notEqualTo,
          between: Excel.ConditionalCellValueOperator.between
        };
        const op = opMap[params.operator] || Excel.ConditionalCellValueOperator.lessThan;
        cf.cellValue.operator = op;
        if (params.formula1 !== undefined) cf.cellValue.formula1 = String(params.formula1);
        if (params.formula2 !== undefined) cf.cellValue.formula2 = String(params.formula2);
        if (params.backgroundColor) cf.cellValue.format.fill.color = params.backgroundColor;
        if (params.fontColor) cf.cellValue.format.font.color = params.fontColor;
      } else if (ruleType === "color_scale" || ruleType === "colorScale") {
        range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
      } else if (ruleType === "data_bar" || ruleType === "dataBar") {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
        const color = params.barColor || params.color || "#00C05E";
        try {
          if (cf.dataBar.positiveFormat) {
            cf.dataBar.positiveFormat.fillColor = color;
            if (params.borderColor) {
              cf.dataBar.positiveFormat.borderColor = params.borderColor;
            }
            if (params.gradientFill !== undefined) {
              cf.dataBar.positiveFormat.gradientFill = Boolean(params.gradientFill);
            }
          }
        } catch (dbErr) {
          console.warn("设置数据条颜色警告:", dbErr);
        }
      }

      await context.sync();
      return { success: true, message: "已应用条件格式" };
    });
  }

