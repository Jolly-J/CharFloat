  // table-pivot.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


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
    // ── CAP-13 字段编排：改**已有**透视表的字段布局并刷新
    //
    // 此前只能建新表（且建完字段就固定了），改布局只能整张删了重建；
    // 读回（CAP-36）已经能看到字段怎么摆，缺的是"让它变成我要的摆法"。
    if ((params || {}).action === "configure") {
      const { workbookName, sheetName, pivotTableName, rowFields = [], columnFields = [],
              filterFields = [], dataFields = [], clearFields = false, refresh = true } = params || {};
      const wb = getWorkbook(app, workbookName);
      const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };

      // 定位透视表：给了名字按名字，否则用第一张
      const sheets = sheetName ? [sheetName] : (() => {
        const out = []; for (let i = 1; i <= wb.Worksheets.Count; i++) { try { out.push(String(wb.Worksheets.Item(i).Name)); } catch (e) {} } return out;
      })();
      let pt = null, ptSheet = null;
      for (const sn of sheets) {
        let pts = null;
        try { pts = wb.Worksheets.Item(sn).PivotTables(); } catch (e) { continue; }
        const n = g(() => Number(pts.Count), 0);
        for (let i = 1; i <= n; i++) {
          const cand = (() => { try { return pts.Item(i); } catch (e) { return null; } })();
          if (!cand) continue;
          if (!pivotTableName || g(() => String(cand.Name), "") === String(pivotTableName)) { pt = cand; ptSheet = sn; break; }
        }
        if (pt) break;
      }
      if (!pt) {
        const names = [];
        for (const sn of sheets) {
          try { const pts = wb.Worksheets.Item(sn).PivotTables(); for (let i = 1; i <= Number(pts.Count); i++) names.push(`${sn}/${g(() => String(pts.Item(i).Name), "?")}`); } catch (e) {}
        }
        throw new Error(`找不到数据透视表${pivotTableName ? ` "${pivotTableName}"` : ""}。现有：${names.join(" / ") || "无"}`);
      }

      const warnings = [];
      const applied = {};
      const availableFields = [];
      try {
        const rf = pt.PivotFields();
        for (let i = 1; i <= Number(rf.Count); i++) { try { availableFields.push(String(rf.Item(i).Name)); } catch (e) {} }
      } catch (e) {}

      const setOrientation = (nameList, orientation, label) => {
        const done = [], missing = [];
        for (const raw of (Array.isArray(nameList) ? nameList : [])) {
          const fieldName = String(raw);
          if (availableFields.length && availableFields.indexOf(fieldName) < 0) { missing.push(fieldName); continue; }
          try { const pf = pt.PivotFields(fieldName); pf.Orientation = orientation; done.push(fieldName); }
          catch (e) { missing.push(fieldName); }
        }
        applied[label] = done;
        if (missing.length) warnings.push(`${label} 有 ${missing.length} 个字段没设上：${missing.join(" / ")}（可用字段：${availableFields.join(" / ") || "读不到"}）`);
      };

      // 先清空（可选）：把除数据字段外的所有字段隐藏，避免旧布局残留
      if (clearFields) {
        let cleared = 0;
        for (const fn of availableFields) {
          try { const pf = pt.PivotFields(fn); if (Number(pf.Orientation) !== 0) { pf.Orientation = 0; cleared++; } } catch (e) {}
        }
        applied.clearedFields = cleared;
      }

      setOrientation(rowFields, 1, "rowFields");
      setOrientation(columnFields, 2, "columnFields");
      setOrientation(filterFields, 3, "filterFields");   // xlPageField = 3

      // 数据字段：名字已存在就跳过（避免重复 AddDataField 产生"求和项:金额2"）
      if (Array.isArray(dataFields) && dataFields.length) {
        const existingCaptions = new Set();
        try { const df = pt.DataFields(); for (let i = 1; i <= Number(df.Count); i++) { try { existingCaptions.add(String(df.Item(i).Name)); } catch (e) {} } } catch (e) {}
        const added = [], skipped = [];
        for (const d of dataFields) {
          const fieldName = String(d && d.fieldName);
          const caption = (d && d.caption) || fieldName;
          if (existingCaptions.has(caption)) { skipped.push(caption); continue; }
          if (availableFields.length && availableFields.indexOf(fieldName) < 0) { warnings.push(`数据字段 "${fieldName}" 不在源字段里，已跳过`); continue; }
          try {
            const pf = pt.PivotFields(fieldName);
            const func = d && d.summaryFunction === "count" ? -4112 : (d && d.summaryFunction === "average" ? -4106 : -4157);
            pt.AddDataField(pf, caption, func);   // -4157=xlSum -4112=xlCount -4106=xlAverage
            added.push(caption);
          } catch (e) { warnings.push(`添加数据字段 "${fieldName}" 失败：${e.message}`); }
        }
        applied.dataFieldsAdded = added;
        if (skipped.length) applied.dataFieldsSkipped = skipped;
      }

      // 刷新
      let refreshedBy = null;
      if (refresh) {
        try { pt.RefreshTable(); refreshedBy = "pivotTable.RefreshTable"; }
        catch (e) { warnings.push(`刷新失败：${e.message}`); }
      }

      // 读回核对：字段的真实 Orientation 与记录数
      const readFields = [];
      try {
        const rf = pt.PivotFields();
        for (let i = 1; i <= Number(rf.Count); i++) {
          const f = rf.Item(i);
          readFields.push({ name: g(() => String(f.Name), null), orientation: g(() => Number(f.Orientation), null) });
        }
      } catch (e) {}
      const checkOrientation = (nameList, want, label) => {
        const asked = (Array.isArray(nameList) ? nameList : []).map(String);
        const bad = asked.filter(n => { const f = readFields.find(x => x.name === n); return f && f.orientation !== want; });
        if (bad.length) warnings.push(`${label} 读回与请求不一致：${bad.join(" / ")}`);
      };
      checkOrientation(rowFields, 1, "rowFields");
      checkOrientation(columnFields, 2, "columnFields");
      checkOrientation(filterFields, 3, "filterFields");

      const recordCount = g(() => { const rr = pt.RowRange; return Number(rr.Rows.Count); }, null);
      return {
        success: warnings.length === 0,
        workbookName: wb.Name, sheetName: ptSheet,
        action: "configure",
        pivotTableName: g(() => String(pt.Name), null),
        availableFields, applied, refreshedBy, recordCount,
        fields: readFields.filter(f => f.orientation !== 0),
        warnings,
        message: `已重排透视表「${g(() => String(pt.Name), "")}」的字段布局（行 ${(applied.rowFields || []).length} / 列 ${(applied.columnFields || []).length} / 筛选 ${(applied.filterFields || []).length}）${warnings.length ? `，${warnings.length} 条告警` : ""}`
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
