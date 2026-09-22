// ── 模块: src/excel/script.js — 动态脚本执行 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1763-1773 行（原样搬迁，未改写）。
  // 12. 动态运行任意脚本
  async function handleRunScript(params) {
    DIAG.runScriptSeen += 1;
    DIAG.lastAt = new Date().toISOString();
    DIAG.lastCodeLen = String((params && (params.code || params.script)) || "").length;
    try {
      return await Excel.run(async (context) => {
        const script = params.code || params.script;
        const fn = new Function("context", "Excel", script);
        const res = await fn(context, Excel);
        await context.sync();
        DIAG.runScriptOk += 1;
        return { success: true, result: res };
      });
    } catch (e) {
      DIAG.runScriptError = String((e && e.message) || e);
      throw e;
    }
  }

