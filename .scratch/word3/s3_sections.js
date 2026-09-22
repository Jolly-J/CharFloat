const out = {};
const R = {};
const D = doc;

R.beforeSections = D.Sections.Count;
// 1) 在文末新增分节符 → 新的一节
try {
  const rng = D.Range(D.Content.End - 1, D.Content.End - 1);
  const newSec = D.Sections.Add(rng);
  R.addedSectionIndex = newSec.Index;
  R.afterSections = D.Sections.Count;
  R.newSecType = newSec.PageSetup.SectionStart; // 2 = wdSectionNewPage
} catch (e) { R.addSectionErr = String(e.message); }

// 2) 在第 2 节写入内容
try {
  const s2 = D.Sections.Item(2);
  const r2 = s2.Range;
  r2.InsertAfter("附录 A · 高级能力实测记录\r本附录位于第 2 节，页面方向为横向。\r");
  R.sec2End = s2.Range.End;
} catch (e) { R.appendErr = String(e.message); }

// 3) 第 2 节横向
try {
  const s2 = D.Sections.Item(2);
  s2.PageSetup.Orientation = 1; // wdOrientLandscape
  R.sec2Page = { orientation: s2.PageSetup.Orientation, w: s2.PageSetup.PageWidth, h: s2.PageSetup.PageHeight };
} catch (e) { R.orientErr = String(e.message); }

// 4) 第 2 节页眉页脚独立
try {
  const s2 = D.Sections.Item(2);
  s2.Headers.Item(1).LinkToPrevious = false;
  s2.Footers.Item(1).LinkToPrevious = false;
  R.links = { h: s2.Headers.Item(1).LinkToPrevious, f: s2.Footers.Item(1).LinkToPrevious };
  s2.Headers.Item(1).Range.Text = "第四部分 · 附录（横向页）";
  s2.Footers.Item(1).Range.Text = "附录 A · 高级能力实测原始记录";
  R.sec2Header = String(s2.Headers.Item(1).Range.Text).replace(/[\r\a]/g, "");
  R.sec2Footer = String(s2.Footers.Item(1).Range.Text).replace(/[\r\a]/g, "");
} catch (e) { R.headerErr = String(e.message); }

// 5) 第 1 节是否被污染
try {
  const s1 = D.Sections.Item(1);
  R.sec1Header = String(s1.Headers.Item(1).Range.Text).replace(/[\r\a]/g, "");
  R.sec1Footer = String(s1.Footers.Item(1).Range.Text).replace(/[\r\a]/g, "");
  R.sec1Page = { orientation: s1.PageSetup.Orientation, w: s1.PageSetup.PageWidth, h: s1.PageSetup.PageHeight };
} catch (e) { R.sec1Err = String(e.message); }
out.r = JSON.stringify(R);
return out;
