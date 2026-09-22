// ── 模块: src/excel/workbook.js — 工作簿元数据、保存、计算、命名项与文档属性 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 400-531 行（原样搬迁，未改写）。
  // 1. 工作簿操作
  async function handleGetWorkbookInfo(params) {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets.load("items/name, items/visibility, items/position");
      const names = context.workbook.names.load("items/name");
      const activeSheet = context.workbook.worksheets.getActiveWorksheet().load("name");
      try {
        context.workbook.load("name");
      } catch {}
      await context.sync();

      let wbName = (params && params.workbookName) || "工作簿1.xlsx";
      try {
        if (context.workbook.name) wbName = context.workbook.name;
      } catch {}

      return {
        hasOpenWorkbook: true,
        workbookName: wbName,
        fullName: wbName,
        activeSheet: activeSheet.name,
        activeSheetName: activeSheet.name,
        sheetCount: sheets.items.length,
        sheets: sheets.items.map(s => ({ name: s.name, visibility: s.visibility, position: s.position })),
        namedItemCount: names.items.length,
        namedItems: names.items.map(n => n.name)
      };
    });
  }

  async function handleGetSheetOutline(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params?.sheetName);
      sheet.load("name");
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("address, rowCount, columnCount, rowIndex, columnIndex, values");
      await context.sync();

      if (usedRange.isNullObject) {
        return {
          sheetName: sheet.name,
          isEmpty: true,
          usedRangeAddress: "A1",
          startRow: 1,
          startColumn: 1,
          rowCount: 0,
          columnCount: 0,
          headerPreview: []
        };
      }

      const headers = (usedRange.values && usedRange.values.length > 0) ? [usedRange.values[0]] : [];
      return {
        sheetName: sheet.name,
        isEmpty: false,
        usedRangeAddress: usedRange.address,
        startRow: usedRange.rowIndex + 1,
        startColumn: usedRange.columnIndex + 1,
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount,
        headerPreview: headers
      };
    });
  }

  async function handleSave(params) {
    return await Excel.run(async (context) => {
      context.workbook.save();
      try {
        context.workbook.load("name");
      } catch {}
      await context.sync();
      const wbName = context.workbook.name || (params && params.workbookName) || "工作簿1.xlsx";
      return {
        success: true,
        saved: true,
        workbookName: wbName,
        fullName: wbName,
        message: `工作簿 [${wbName}] 已成功保存到磁盘`
      };
    });
  }

  async function handleCalculate(params) {
    return await Excel.run(async (context) => {
      const type = params.type || "FullRebuild";
      context.application.calculate(type);
      await context.sync();
      return { success: true, calculationType: type };
    });
  }

  async function handleListNamedItems() {
    return await Excel.run(async (context) => {
      const items = context.workbook.names.load("items/name, items/value, items/type, items/visible");
      await context.sync();
      return {
        items: items.items.map(i => ({ name: i.name, value: i.value, type: i.type, visible: i.visible }))
      };
    });
  }

  async function handleUpdateNamedItem(params) {
    return await Excel.run(async (context) => {
      const { name, reference, comment, action } = params;
      if (action === "delete") {
        context.workbook.names.getItem(name).delete();
      } else {
        context.workbook.names.add(name, reference, comment);
      }
      await context.sync();
      return { success: true, name };
    });
  }

  async function handleGetProperties() {
    return await Excel.run(async (context) => {
      const p = context.workbook.properties.load("title, subject, author, keywords, comments, category, manager, company");
      await context.sync();
      return { properties: p };
    });
  }

  async function handleUpdateProperties(params) {
    return await Excel.run(async (context) => {
      const props = params.properties || params;
      context.workbook.properties.set(props);
      await context.sync();
      return { success: true, updated: props };
    });
  }

