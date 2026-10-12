  // pure.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  /** 1 基列号 → 列名（A/B/.../AA）；只用来拼地址字符串，不做宿主调用。 */
  function excelColumnName(index) {
    let n = Math.trunc(Number(index) || 0);
    let name = "";
    while (n > 0) {
      const m = (n - 1) % 26;
      name = String.fromCharCode(65 + m) + name;
      n = Math.floor((n - 1) / 26);
    }
    return name;
  }

  function normalize2DArray(val, expectedRows, expectedCols) {
    if (val === undefined || val === null) {
      return Array(expectedRows).fill(null).map(() => Array(expectedCols).fill(null));
    }
    if (!Array.isArray(val)) return [[val]];
    if (!Array.isArray(val[0])) return [val];
    return val;
  }

  // @build-strip:start — 仅用于 Node 单元测试
  export { normalize2DArray, excelColumnName };
  // @build-strip:end
