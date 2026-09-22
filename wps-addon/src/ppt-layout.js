  // ---------------------------------------------------------------------------
  // ppt-layout.js — PPT 布局几何/字号/容量计算（纯函数，不调用任何宿主 API）
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进。
  // 例外：文件末尾的 @build-strip 导出块只为 Node 单元测试直接 import 使用，
  //      构建拼接时由 scripts/build-wps-addon.mjs 整段移除，不会进入 addon-core.js。
  // ---------------------------------------------------------------------------

  // Office geometry and font sizes are points (72 pt = 1 inch).
  // 只读 pres.PageSetup，可传入普通对象（{ PageSetup: { SlideWidth, SlideHeight } }），
  // 因此本函数可在无宿主环境下直接单元测试。
  function pptPageSize(pres) {
    const width = Number(pres.PageSetup.SlideWidth);
    const height = Number(pres.PageSetup.SlideHeight);
    if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
      throw new Error("无法读取有效的 PPT 页面尺寸，停止布局；请检查 PageSetup");
    }
    return { pageWidth: width, pageHeight: height, unit: "pt" };
  }

  // 设计基准 720x405 (pt) 到真实页面尺寸的缩放系数；算法与拆分前 fitGeneratedPptShapes 内联写法一致。
  function pptScaleForPage(page) {
    const sx = page.pageWidth / 720, sy = page.pageHeight / 405;
    return { sx: sx, sy: sy, scale: Math.min(sx, sy) };
  }

  // 文本视觉宽度：全角 1、半角 0.6；与拆分前 linesAt 内的换算一致。
  function pptCountTextUnits(text) {
    return Array.from(text).reduce((n, ch) => n + (ch.charCodeAt(0) > 255 ? 1 : 0.6), 0);
  }

  // 指定字号下的预估行数：显式换行逐段累加，每段按内宽向上取整（最少 1 行）。
  function pptEstimateLines(text, size, innerWidth) {
    return text.split(/\r\n|\r|\n/).reduce((sum, line) => {
      const units = pptCountTextUnits(line);
      return sum + Math.max(1, Math.ceil(units * size / innerWidth));
    }, 0);
  }

  // 在 [minimum, preferred] 内按 scale 步长下调字号，直到估算高度不再溢出（或触底）。
  function pptFitFontSize(text, preferred, minimum, scale, innerWidth, innerHeight) {
    let size = preferred;
    while (size > minimum && pptEstimateLines(text, size, innerWidth) * size * 1.25 > innerHeight) {
      size = Math.max(minimum, size - scale);
    }
    return size;
  }

  // @build-strip:start — Node 单元测试导出（构建拼接时整段移除）
  export { pptPageSize, pptScaleForPage, pptCountTextUnits, pptEstimateLines, pptFitFontSize };
  // @build-strip:end
