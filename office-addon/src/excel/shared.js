// ── 模块: src/excel/shared.js — Excel 操作公共辅助（目标工作表解析） ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 389-399 行（原样搬迁，未改写）。
  // ==========================================
  // 具体 API 业务实现 (借助 Excel.run)
  // ==========================================

  function getTargetSheet(context, sheetName) {
    // 顺手把 `name` 排进加载队列：调用方普遍会把 `sheet.name` 写进返回体，
    // 而 Office.js 读未 load 的属性会直接抛「属性"name"不可用」。
    // 在这里统一 load 一次，避免每个处理器各写一遍（真机实测踩到过：形状工具全部报该错）。
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    return sheet;
  }

