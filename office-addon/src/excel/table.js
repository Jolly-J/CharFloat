// ── 模块: src/excel/table.js — 结构化表格 Table ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1169-1194 行（原样搬迁，未改写）。
  // 7. 结构化表格
  async function handleCreateTable(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);
      const table = sheet.tables.add(range, params.hasHeaders !== false);
      if (params.name) table.name = params.name;
      if (params.style) table.style = params.style;
      table.load("name, id");
      await context.sync();
      return { success: true, name: table.name, id: table.id };
    });
  }

  async function handleUpdateTable(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const table = sheet.tables.getItem(params.name || params.tableName);
      if (params.style) table.style = params.style;
      if (params.showTotals !== undefined) table.showTotals = params.showTotals;
      if (params.resizeRange) table.resize(sheet.getRange(params.resizeRange));
      await context.sync();
      return { success: true };
    });
  }

