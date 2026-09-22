// ── 模块: src/excel/sheets.js — 工作表生命周期 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 532-613 行（原样搬迁，未改写）。
  // 2. 工作表操作
  async function handleListSheets() {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets.load("items/name, items/visibility, items/tabColor");
      await context.sync();
      return {
        sheets: sheets.items.map(s => ({ name: s.name, visibility: s.visibility, tabColor: s.tabColor }))
      };
    });
  }

  async function handleAddSheet(params) {
    return await Excel.run(async (context) => {
      const name = params.name || params.sheetName;
      const sheet = context.workbook.worksheets.add(name);
      sheet.load("name, position");
      await context.sync();
      return { success: true, name: sheet.name, position: sheet.position };
    });
  }

  async function handleUpdateSheet(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName || params.oldName);
      if (params.name || params.newName) sheet.name = params.name || params.newName;
      if (params.visibility) sheet.visibility = params.visibility;
      if (params.tabColor) sheet.tabColor = params.tabColor;
      if (params.activate) sheet.activate();
      await context.sync();
      return { success: true };
    });
  }

  async function handleDeleteSheet(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName || params.name);
      sheet.delete();
      await context.sync();
      return { success: true };
    });
  }

  async function handleCopySheet(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sourceSheet);
      const copied = sheet.copy(params.positionType || "After", sheet);
      if (params.newName) copied.name = params.newName;
      copied.load("name");
      await context.sync();
      return { success: true, name: copied.name };
    });
  }

  async function handleManageSheet(params) {
    return await Excel.run(async (context) => {
      const action = params.action || (params.newSheetName ? "copy" : "rename");
      const sheet = getTargetSheet(context, params.sheetName || params.oldName || params.sourceSheet);

      if (action === "copy" || action === "duplicate") {
        const copied = sheet.copy(Excel.WorksheetPositionType.after, sheet);
        if (params.newSheetName || params.newName) {
          copied.name = params.newSheetName || params.newName;
        }
        copied.load("name, position");
        await context.sync();
        return { success: true, sheetName: copied.name, position: copied.position };
      } else if (action === "rename") {
        sheet.name = params.newName || params.newSheetName;
      } else if (action === "hide") {
        sheet.visibility = Excel.SheetVisibility.hidden;
      } else if (action === "show") {
        sheet.visibility = Excel.SheetVisibility.visible;
      } else if (action === "color" && params.tabColor) {
        sheet.tabColor = params.tabColor;
      }

      sheet.load("name");
      await context.sync();
      return { success: true, sheetName: sheet.name, action };
    });
  }

