// ── 模块: src/excel/pivot.js — 数据透视表 PivotTable ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1663-1687 行（原样搬迁，未改写）。
  // 9. 数据透视表
  async function handleCreatePivotTable(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const source = sheet.getRange(params.sourceAddress);
      const dest = sheet.getRange(params.destinationAddress || "R1C1");
      const name = params.name || "PivotTable1";

      const pt = sheet.pivotTables.add(name, source, dest);
      pt.load("name");
      await context.sync();
      return { success: true, name: pt.name };
    });
  }

  async function handleUpdatePivotTable(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const pt = sheet.pivotTables.getItem(params.name || params.pivotTableName);
      if (params.refresh) pt.refresh();
      await context.sync();
      return { success: true };
    });
  }

