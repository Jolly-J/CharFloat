const out = {};
const D = doc;
const R = { hits: [] };
for (let i = 1; i <= D.Paragraphs.Count; i++) {
  const t = String(D.Paragraphs.Item(i).Range.Text).replace(/[\r\a\n]/g, "");
  if (/pct|\u0001/.test(t)) {
    const p = D.Paragraphs.Item(i);
    R.hits.push({ i: i, text: t.slice(0, 70), bold: p.Range.Font.Bold, font: p.Range.Font.Name, size: p.Range.Font.Size });
  }
}
// 表格内是否也被替换
R.tableHits = [];
for (let ti = 1; ti <= D.Tables.Count; ti++) {
  const T = D.Tables.Item(ti);
  for (let r = 1; r <= T.Rows.Count; r++) {
    for (let c = 1; c <= T.Columns.Count; c++) {
      try { const t = String(T.Cell(r, c).Range.Text).replace(/[\r\a\n]/g, "");
        if (/pct|\u0001/.test(t)) R.tableHits.push({ t: ti, r: r, c: c, text: t.slice(0, 40) }); } catch (e) {}
    }
  }
}
out.r = JSON.stringify(R);
return out;
