const out = {};
const R = {};
const D = doc;
R.sectionCount = D.Sections.Count;
R.sections = [];
for (let i = 1; i <= D.Sections.Count; i++) {
  const s = D.Sections.Item(i);
  const e = { i: i, orientation: s.PageSetup.Orientation, diffFirst: s.PageSetup.DifferentFirstPageHeaderFooter,
    oddEven: s.PageSetup.OddAndEvenPagesHeaderFooter,
    h: String(s.Headers.Item(1).Range.Text).replace(/[\r\a]/g, ""),
    f: String(s.Footers.Item(1).Range.Text).replace(/[\r\a]/g, "") };
  try { e.hFirst = String(s.Headers.Item(2).Range.Text).replace(/[\r\a]/g, ""); } catch (x) { e.hFirst = "ERR"; }
  try { e.fFirst = String(s.Footers.Item(2).Range.Text).replace(/[\r\a]/g, ""); } catch (x) { e.fFirst = "ERR"; }
  // 奇偶页
  try { e.hEven = String(s.Headers.Item(3).Range.Text).replace(/[\r\a]/g, ""); } catch (x) { e.hEven = "ERR"; }
  R.sections.push(e);
}
// 水印数量
try { R.watermarkShapes = D.Shapes.Count; R.shapeNames = (() => { const a = []; for (let i = 1; i <= D.Shapes.Count; i++) a.push(D.Shapes.Item(i).Name + "/" + D.Shapes.Item(i).Type); return a; })(); } catch (e) { R.shapeErr = String(e.message); }
out.r = JSON.stringify(R);
return out;
