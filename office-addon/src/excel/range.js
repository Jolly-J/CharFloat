// ── 模块: src/excel/range.js — 区域读写、结构与查找替换 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 614-918 行（原样搬迁，未改写）。
  // 3. 区域核心读写
  async function handleReadRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      let range;
      if (params.address) {
        range = sheet.getRange(params.address);
      } else {
        range = sheet.getUsedRange();
      }

      range.load("address, rowCount, columnCount, values, formulas, text, numberFormat");
      await context.sync();

      if (params.metadata_only) {
        return {
          address: range.address,
          rowCount: range.rowCount,
          columnCount: range.columnCount
        };
      }

      return {
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values,
        formulas: range.formulas,
        text: range.text
      };
    });
  }

  async function handleGetRangeStyles(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);
      range.load("address, format/font/name, format/font/size, format/font/bold, format/font/color, format/fill/color, format/horizontalAlignment, format/verticalAlignment, numberFormat");
      await context.sync();

      return {
        success: true,
        address: range.address,
        font: {
          name: range.format.font.name,
          size: range.format.font.size,
          bold: range.format.font.bold,
          color: range.format.font.color
        },
        fill: {
          color: range.format.fill.color
        },
        alignment: {
          horizontal: range.format.horizontalAlignment,
          vertical: range.format.verticalAlignment
        },
        numberFormat: range.numberFormat ? range.numberFormat[0]?.[0] : null
      };
    });
  }

  async function handleWriteRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      let range;
      if (params.address) {
        range = sheet.getRange(params.address);
      } else {
        throw new Error("write_range 必须指定 address 参数 (如 'A1' 或 'B2:D5')");
      }

      let values = params.values || params.data;
      let formulas = params.formulas;

      // 标准化 values 为规范二维数组
      if (values !== undefined && values !== null) {
        if (!Array.isArray(values)) {
          values = [[values]];
        } else if (!Array.isArray(values[0])) {
          values = [values];
        }
      }

      // 标准化 formulas 为规范二维数组
      if (formulas !== undefined && formulas !== null) {
        if (!Array.isArray(formulas)) {
          formulas = [[formulas]];
        } else if (!Array.isArray(formulas[0])) {
          formulas = [formulas];
        }
      }

      const primaryMatrix = formulas || values;
      if (!primaryMatrix) {
        throw new Error("write_range 必须提供 values 或 formulas 参数");
      }

      const rowCount = primaryMatrix.length;
      const columnCount = primaryMatrix[0].length;

      // 如果给的是单一左上角单元格且提供了二维数组，自动按矩阵尺寸扩展 range
      if (!params.address.includes(":") && (rowCount > 1 || columnCount > 1)) {
        range = range.getResizedRange(rowCount - 1, columnCount - 1);
      }

      // 如果仅有 values，但 values 中包含以 '=' 开头的字符串，自动提取为公式矩阵
      if (values && !formulas) {
        const hasFormulasInValues = values.some(row => Array.isArray(row) && row.some(cell => typeof cell === 'string' && cell.startsWith('=')));
        if (hasFormulasInValues) {
          formulas = values.map(row => (Array.isArray(row) ? row.map(cell => (typeof cell === 'string' && cell.startsWith('=')) ? cell : null) : [null]));
        }
      }

      // 赋值 values (先写入所有常规数值与文字)
      if (values) {
        range.values = values;
      }

      // 赋值 formulas (仅针对实际包含公式的单元格精准注入，杜绝空字符串抹除数据)
      if (formulas) {
        for (let r = 0; r < rowCount; r++) {
          for (let c = 0; c < columnCount; c++) {
            const f = formulas[r]?.[c];
            if (f && typeof f === 'string' && f.startsWith('=')) {
              range.getCell(r, c).formulas = [[f]];
            }
          }
        }
      }

      // 关键：在 sync 前必须显式 load 需要返回的属性，绝不裸读！
      range.load("address");
      sheet.load("name");
      await context.sync();

      return {
        success: true,
        address: range.address || params.address,
        sheetName: sheet.name,
        rowCount: rowCount,
        columnCount: columnCount,
        modifiedCount: rowCount * columnCount
      };
    });
  }

  async function handleUpdateRangeStructure(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const action = params.action || "insert";
      const dimension = params.dimension || "rows";
      let range;

      if (params.address) {
        range = sheet.getRange(params.address);
      } else if (params.index !== undefined) {
        const idx = Math.max(1, Number(params.index));
        const count = Math.max(1, Number(params.count || 1));
        if (dimension === "columns") {
          const colLetter = (colIdx) => {
            let temp, letter = '';
            while (colIdx > 0) {
              temp = (colIdx - 1) % 26;
              letter = String.fromCharCode(temp + 65) + letter;
              colIdx = Math.floor((colIdx - temp - 1) / 26);
            }
            return letter;
          };
          const startCol = colLetter(idx);
          const endCol = colLetter(idx + count - 1);
          range = sheet.getRange(`${startCol}:${endCol}`);
        } else {
          range = sheet.getRange(`${idx}:${idx + count - 1}`);
        }
      } else {
        throw new Error("update_range_structure 必须提供 address 或 (index, dimension)");
      }

      const shift = params.shift || (dimension === "columns" ? "Right" : "Down");

      if (action === "insert") {
        range.insert(shift);
      } else if (action === "delete") {
        range.delete(dimension === "columns" ? "Left" : "Up");
      }

      await context.sync();
      return { success: true, action, dimension, address: params.address };
    });
  }

  async function handleClearRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const targetRange = params.address || params.range;
      const range = sheet.getRange(targetRange);
      const applyTo = params.applyTo || "All";
      range.clear(applyTo);
      await context.sync();
      return { success: true, address: targetRange, cleared: applyTo };
    });
  }

  async function handleCopyRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const sourceRange = sheet.getRange(params.sourceAddress);
      const destRange = sheet.getRange(params.destinationAddress);
      destRange.copyFrom(sourceRange, params.copyType || "All");
      await context.sync();
      return { success: true };
    });
  }

  async function handleSetHyperlink(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);
      range.hyperlink = {
        address: params.url,
        textToDisplay: params.textToDisplay || params.url,
        screenTip: params.screenTip || ""
      };
      await context.sync();
      return { success: true, address: params.address };
    });
  }

  async function handleSetDataValidation(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);
      range.dataValidation.clear();
      if (params.rule) {
        range.dataValidation.rule = params.rule;
      }
      await context.sync();
      return { success: true, address: params.address };
    });
  }

  async function handleFindReplace(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = params.address ? sheet.getRange(params.address) : sheet.getUsedRange();
      // 字段名兼容：网关（WPS 语义）传的是 searchQuery，Office.js 侧原来只认 text/query/findText。
      // 名字对不上 → text 落成空串 → `replaceAll("", …)` 会命中整片区域并可能破坏内容（问题台账 ISS-93）。
      const text = params.text || params.query || params.findText || params.searchQuery || "";
      if (!text) {
        throw new Error(
          "find_and_replace 缺少搜索文本（Office.js 通道识别 text / query / findText / searchQuery，当前都为空）。" +
          "空搜索串会命中整个区域并可能破坏内容，已阻断；请显式提供要查找的文本。"
        );
      }
      const replaceText = params.replaceText;
      const matchCase = !!params.matchCase;
      const matchEntireCell = !!params.matchEntireCell;

      if (replaceText !== undefined) {
        range.replaceAll(text, replaceText, { completeMatch: matchEntireCell, matchCase: matchCase });
        await context.sync();
        return { success: true, replacedWith: replaceText };
      }

      // 跨平台安全搜索：macOS Office.js 缺少 range.findAll，采用纯 JS 内存矩阵检索
      range.load("values, text, rowIndex, columnIndex, address");
      await context.sync();

      function colToLetters(c) {
        let temp = c + 1;
        let letter = "";
        while (temp > 0) {
          const mod = (temp - 1) % 26;
          letter = String.fromCharCode(65 + mod) + letter;
          temp = Math.floor((temp - mod) / 26);
        }
        return letter;
      }

      const baseRow = range.rowIndex || 0;
      const baseCol = range.columnIndex || 0;
      const matches = [];
      const queryStr = matchCase ? String(text) : String(text).toLowerCase();
      const vals = range.values || [];

      for (let r = 0; r < vals.length; r++) {
        const row = vals[r] || [];
        for (let c = 0; c < row.length; c++) {
          const cellVal = row[c];
          if (cellVal === null || cellVal === undefined) continue;
          const str = String(cellVal);
          const targetStr = matchCase ? str : str.toLowerCase();
          const matched = matchEntireCell ? targetStr === queryStr : targetStr.includes(queryStr);
          if (matched) {
            const cellAddr = `${colToLetters(baseCol + c)}${baseRow + r + 1}`;
            matches.push({
              address: cellAddr,
              text: str,
              row: baseRow + r + 1,
              column: baseCol + c + 1
            });
          }
        }
      }

      return {
        success: true,
        count: matches.length,
        matches: matches
      };
    });
  }

