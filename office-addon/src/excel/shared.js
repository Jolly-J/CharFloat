// ── 模块: src/excel/shared.js — Excel 操作公共辅助（目标工作表解析） ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 389-399 行（原样搬迁，未改写）。
  // ==========================================
  // 具体 API 业务实现 (借助 Excel.run)
  // ==========================================

  function getTargetSheet(context, sheetName) {
    if (sheetName) {
      return context.workbook.worksheets.getItem(sheetName);
    }
    return context.workbook.worksheets.getActiveWorksheet();
  }

