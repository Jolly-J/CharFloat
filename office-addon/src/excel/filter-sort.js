// ── 模块: src/excel/filter-sort.js — 排序与筛选 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1137-1168 行（原样搬迁，未改写）。
  // 6. 排序与筛选
  async function handleSetFilterAndSort(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);
      const action = params.action || (params.filterColumn !== undefined ? "filter" : "sort");

      if (action === "filter") {
        const col = params.filterColumn !== undefined ? params.filterColumn : 0;
        const criteria = params.criteria ? (Array.isArray(params.criteria) ? params.criteria : [String(params.criteria)]) : [];
        sheet.autoFilter.apply(range, col, {
          filterOn: Excel.FilterOn.values,
          values: criteria
        });
      } else if (action === "clear") {
        sheet.autoFilter.remove();
        range.sort.clear();
      } else {
        const col = params.sortColumn !== undefined ? params.sortColumn : (params.key || 0);
        const ascending = params.sortOrder === "desc" ? false : (params.ascending !== false);
        range.sort.apply([{
          key: col,
          ascending: ascending,
          sortOn: Excel.SortOn.value
        }], params.hasHeader !== false);
      }

      await context.sync();
      return { success: true, action, address: params.address };
    });
  }

