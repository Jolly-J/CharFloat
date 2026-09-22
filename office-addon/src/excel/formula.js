// ── 模块: src/excel/formula.js — 公式写入 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 919-936 行（原样搬迁，未改写）。
  // 4. 公式
  async function handleSetFormula(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);
      let formula = params.formula;

      if (Array.isArray(formula)) {
        range.formulas = formula;
      } else {
        if (!formula.startsWith("=")) formula = "=" + formula;
        range.formulas = [[formula]];
      }
      await context.sync();
      return { success: true, address: params.address, formula };
    });
  }

