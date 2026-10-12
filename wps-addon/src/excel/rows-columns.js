  // rows-columns.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


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
