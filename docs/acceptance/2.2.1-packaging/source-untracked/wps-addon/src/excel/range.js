  // range.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  // 5. 清理指定区域
  function clearRange(app, params) {
    const { sheetName, address, workbookName } = params;
    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);
    range.Clear();
    return { success: true, workbookName: sheet.Parent.Name, clearedAddress: range.Address() };
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

  // 10.5 全局查找与定位替换
  function findAndReplace(app, params) {
    const { sheetName, workbookName, searchQuery, replaceText, matchCase = false, matchEntireCell = false, searchRange, maxResults = 50 } = params || {};
    if (searchQuery === undefined || searchQuery === null) {
      throw new Error("缺少必要参数: searchQuery (要查找的文本或数值)");
    }
    // ISS-98 同源护栏（**宿主侧最后一道，不依赖上游**）：空串在宿主侧是"恒真匹配"——
    // `String(val).includes("")` 对每个非空单元格都成立，配合 replaceText 还会按**空正则**逐格替换
    // （`"abc".replace(new RegExp("","gi"), "X")` → 每个字符之间都插入 X），**整片内容被改坏**。
    // 桥接层已在工具入口拦截（src/bridge/gateway/excel.ts），但加载项可被 WebSocket RPC 直接调用，
    // 护栏只留一层就等于没有。拦截发生在**任何读取/写入之前**，被拒时文档零改动。
    if (String(searchQuery) === "") {
      throw new Error(
        "searchQuery 不能为空串：空串在宿主侧会命中区域内每个非空单元格（includes('') 恒真），" +
        "配合 replaceText 会把整片内容改坏；请显式提供要查找的文本或数值。本次未做任何修改。"
      );
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

  /** 文档属性（内置 + 自定义）。 */
  // ── CAP-14 自定义视图 + 切片器
  //
  // 真机探测（WPS 12.1.28496）：
  //   `wb.CustomViews` **可用**（Add/Item/Show 都在，实测能建成并读回 Name）
  //   `wb.SlicerCaches` 可用（Count/Add/Add2）；`ws.Slicers` **不存在**
  //   `wb.LinkedDataTypes` **不存在** → **链接数据类型（富值）本机做不到**，如实拒绝
  // ── CAP-11 变更感知（基线比对版）
  //
  // 真机探测结论：WPS 的**事件机制不可用** ——
  //   `wps.ApiEvent.AddApiEventListener("SheetChange"/"WorkbookBeforeSave", fn)` **注册不报错**，
  //   但真机写入单元格后**事件从不触发**（命中恒为 0），与线程批注同属"API 存在但功能是空壳"。
  //   因此"保存前拦一手"这类**必须依赖事件**的能力本机做不到。
  //
  // 能做到的是 CAP-11 的核心价值：**回答"用户刚改了哪个单元格"** ——
  // 用基线快照 + 差异比对实现，不依赖事件。
  const CHANGE_BASELINES = {};   // key: 工作簿!工作表 → { cells: {addr: value}, at, range }

  function captureOrDiffChanges(app, params) {
    const { workbookName, sheetName, action = "capture", maxCells = 2000 } = params || {};
    const wb = getWorkbook(app, workbookName);
    const sheet = getWorksheet(app, sheetName, workbookName);
    const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
    const key = `${wb.Name}!${sheet.Name}`;

    // 读当前已用区域的值（有上限，避免大表卡死）
    const readNow = () => {
      const used = sheet.UsedRange;
      const base = String(g(() => used.Address(), ""));
      const rows = Math.min(Number(g(() => used.Rows.Count, 0)), 500);
      const cols = Math.min(Number(g(() => used.Columns.Count, 0)), 60);
      if (!rows || !cols) return { cells: {}, range: base, truncated: false, count: 0 };
      const m = base.split(":")[0].replace(/\$/g, "").match(/^([A-Za-z]+)(\d+)$/);
      if (!m) return { cells: {}, range: base, truncated: false, count: 0 };
      const colName = (n) => { let s = ""; while (n > 0) { const x = (n - 1) % 26; s = String.fromCharCode(65 + x) + s; n = Math.floor((n - 1) / 26); } return s; };
      const col0 = m[1].toUpperCase().split("").reduce((a, c) => a * 26 + (c.charCodeAt(0) - 64), 0);
      const row0 = Number(m[2]);
      const addr = `${colName(col0)}${row0}:${colName(col0 + cols - 1)}${row0 + rows - 1}`;
      const vals = g(() => sheet.Range(addr).Value2, null);
      const cells = {};
      let count = 0;
      if (Array.isArray(vals)) {
        for (let r = 0; r < vals.length && count < maxCells; r++) {
          const row = vals[r];
          if (!Array.isArray(row)) { cells[`${colName(col0)}${row0 + r}`] = row; count++; continue; }
          for (let c = 0; c < row.length && count < maxCells; c++) {
            cells[`${colName(col0 + c)}${row0 + r}`] = row[c];
            count++;
          }
        }
      } else if (vals !== null && vals !== undefined) {
        cells[`${colName(col0)}${row0}`] = vals; count = 1;
      }
      const total = rows * cols;
      return { cells, range: addr, truncated: total > maxCells, count, total };
    };

    if (action === "capture") {
      const snap = readNow();
      CHANGE_BASELINES[key] = { cells: snap.cells, at: Date.now(), range: snap.range };
      return {
        success: true, workbookName: wb.Name, sheetName: sheet.Name, action,
        range: snap.range, cellCount: snap.count, truncated: snap.truncated,
        warnings: snap.truncated ? [`只记录了前 ${maxCells} 个单元格，超出部分不会参与后续比对`] : [],
        message: `已记录 [${sheet.Name}] ${snap.range} 的基线（${snap.count} 个单元格），之后可用 action='diff' 查看谁改了什么`
      };
    }

    if (action === "diff") {
      const base = CHANGE_BASELINES[key];
      if (!base) throw new Error(`工作表 [${sheet.Name}] 还没有基线；先调 action='capture' 记一次`);
      const now = readNow();
      const norm = (v) => (v === null || v === undefined) ? "" : String(v);
      const changed = [], added = [], removed = [];
      const seen = new Set();
      for (const [addr, v] of Object.entries(now.cells)) {
        seen.add(addr);
        if (!(addr in base.cells)) { added.push({ address: addr, after: v }); continue; }
        if (norm(base.cells[addr]) !== norm(v)) changed.push({ address: addr, before: base.cells[addr], after: v });
      }
      for (const [addr, v] of Object.entries(base.cells)) {
        if (!seen.has(addr)) removed.push({ address: addr, before: v });
      }
      const all = [...changed, ...added, ...removed];
      return {
        success: true, workbookName: wb.Name, sheetName: sheet.Name, action,
        baselineAt: new Date(base.at).toISOString(),
        baselineRange: base.range, currentRange: now.range,
        changedCount: changed.length, addedCount: added.length, removedCount: removed.length,
        totalChanged: all.length,
        changed: changed.slice(0, 200), added: added.slice(0, 200), removed: removed.slice(0, 200),
        truncated: all.length > 200,
        warnings: now.truncated ? ["当前读取被 maxCells 截断，差异可能不完整"] : [],
        message: all.length === 0
          ? `[${sheet.Name}] 自基线以来没有任何变化`
          : `[${sheet.Name}] 自基线以来有 ${all.length} 处变化（改 ${changed.length} / 增 ${added.length} / 删 ${removed.length}）`
      };
    }

    if (action === "clear") {
      const had = !!CHANGE_BASELINES[key];
      delete CHANGE_BASELINES[key];
      return { success: true, workbookName: wb.Name, sheetName: sheet.Name, action, hadBaseline: had, warnings: [], message: had ? `已清除 [${sheet.Name}] 的基线` : `[${sheet.Name}] 本来就没有基线` };
    }

    throw new Error(`未知的变更感知操作 action: ${action}（支持 capture / diff / clear）`);
  }
