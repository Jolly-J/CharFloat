  // ---------------------------------------------------------------------------
  // sheet-sort.js — 表格取值比较与"是否已排序"判定的纯函数（不调用任何宿主 API）
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 抽出来的原因：排序写完必须**读回校验**才能避免"返回 success 但顺序没变"（问题台账 ISS-38），
  // 而校验逻辑必须可单元测试——否则"校验恒真"这种缺陷没人发现得了。
  // 例外：文件末尾的 @build-strip 导出块只为 Node 单元测试直接 import 使用，构建拼接时整段移除。
  // ---------------------------------------------------------------------------

  /** 单元格比较：空值最小，数字按数值，其余按字符串。 */
  function compareCellValues(a, b) {
    if (a === b) return 0;
    const emptyA = a === null || a === undefined || a === "";
    const emptyB = b === null || b === undefined || b === "";
    if (emptyA && emptyB) return 0;
    if (emptyA) return -1;
    if (emptyB) return 1;
    const na = Number(a), nb = Number(b);
    if (Number.isFinite(na) && Number.isFinite(nb)) return na === nb ? 0 : (na < nb ? -1 : 1);
    const sa = String(a), sb = String(b);
    return sa === sb ? 0 : (sa < sb ? -1 : 1);
  }

  /**
   * 校验二维数据是否已按 rules 指定的列升/降序。
   *
   * - `rules[].colIndex` 是**区域内的相对列号**（与实现里 `targetRange.Columns.Item` 一致）；
   * - 第 0 行按表头处理（排序时传 `header=1`），因此相邻对从 (1,2) 开始；
   * - 相邻对覆盖到最后一对 `(n-2, n-1)`，**不得漏最后一行**；
   * - 少于 3 行（表头 + 1 行数据）时无需比较，直接视为已排序。
   */
  function isSortedByRules(matrix, rules) {
    if (!Array.isArray(matrix) || !Array.isArray(rules) || rules.length === 0) return true;
    if (matrix.length < 3) return true;
    for (let i = 2; i < matrix.length; i++) {
      for (let r = 0; r < rules.length; r++) {
        const col = Number(rules[r].colIndex) - 1;
        if (!Number.isFinite(col) || col < 0) return false;
        const desc = rules[r].order === "desc";
        const cmp = compareCellValues(matrix[i - 1] ? matrix[i - 1][col] : null, matrix[i] ? matrix[i][col] : null);
        if (cmp === 0) continue;
        if (desc ? cmp < 0 : cmp > 0) return false;
        break;
      }
    }
    return true;
  }

  // @build-strip:start — Node 单元测试导出（构建拼接时整段移除）
  export { compareCellValues, isSortedByRules };
  // @build-strip:end
