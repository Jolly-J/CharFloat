const out = {};
const D = doc;
const R = { total: D.Paragraphs.Count, tail: [] };
for (let i = Math.max(1, D.Paragraphs.Count - 12); i <= D.Paragraphs.Count; i++) {
  R.tail.push({ i: i, t: String(D.Paragraphs.Item(i).Range.Text).replace(/[\r\a]/g, "").slice(0, 80) });
}
out.r = JSON.stringify(R);
return out;
