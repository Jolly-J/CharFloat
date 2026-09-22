// 01-normalize.js — 把 generate_deck 留下的 720x405 设计坐标换算到真实页面尺寸。
// 背景：generate_deck 声称按真实页面尺寸生成，实测所有形状仍是 720x405 设计坐标，
// 只占 960x540 页面的左上 75%。这里等价重做 addon 内部 fitGeneratedPptShapes 的换算。
const page = { w: Number(pres.PageSetup.SlideWidth), h: Number(pres.PageSetup.SlideHeight) };
const sx = page.w / 720, sy = page.h / 405;
const scale = Math.min(sx, sy);
const report = [];
for (let s = 1; s <= pres.Slides.Count; s++) {
  const slide = pres.Slides.Item(s);
  let moved = 0, fontScaled = 0;
  for (let i = 1; i <= slide.Shapes.Count; i++) {
    const sh = slide.Shapes.Item(i);
    sh.Left = Number(sh.Left) * sx;
    sh.Top = Number(sh.Top) * sy;
    sh.Width = Number(sh.Width) * sx;
    sh.Height = Number(sh.Height) * sy;
    moved++;
    if (sh.HasTextFrame && sh.TextFrame.HasText) {
      const frame = sh.TextFrame;
      frame.AutoSize = 0;
      frame.WordWrap = true;
      const range = frame.TextRange;
      const size = Number(range.Font.Size);
      if (size > 0) { range.Font.Size = Math.round(size * scale * 10) / 10; fontScaled++; }
    }
  }
  report.push({ slide: s, shapes: moved, fonts: fontScaled });
}
return { page: page, sx: sx, sy: sy, scale: scale, slides: report };
