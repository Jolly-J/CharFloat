const out = {};
const R = {};
const D = doc;
R.paraCount = D.Paragraphs.Count;
R.hits = [];
for (let i = 1; i <= D.Paragraphs.Count; i++) {
  const t = String(D.Paragraphs.Item(i).Range.Text);
  if (t.indexOf("书签定位写入测试") >= 0) R.hits.push({ i: i, text: t.replace(/[\r\a]/g, "").slice(0, 60),
    style: D.Paragraphs.Item(i).Style.NameLocal });
  if (t.indexOf("2.2 模块交付状态") >= 0) R.para22 = { i: i, text: t.replace(/[\r\a]/g, "").slice(0, 30) };
  if (t.indexOf("三、验收测试结果") >= 0 && t.length < 40) R.heading3 = { i: i, text: t.replace(/[\r\a]/g, "").slice(0, 30) };
}
// 书签是否被撑大
try {
  for (let i = 1; i <= D.Bookmarks.Count; i++) {
    const b = D.Bookmarks.Item(i);
    R["bm_" + b.Name] = { start: b.Start, end: b.End, text: String(b.Range.Text).replace(/[\r\a]/g, "").slice(0, 60) };
  }
} catch (e) { R.bmErr = String(e.message); }
out.r = JSON.stringify(R);
return out;
