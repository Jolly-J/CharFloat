const out = {};
const R = { hits: [] };
const D = doc;
for (let i = 1; i <= D.Paragraphs.Count; i++) {
  const t = String(D.Paragraphs.Item(i).Range.Text).replace(/[\r\a]/g, "");
  if (t.indexOf("【探针】") >= 0 || t.indexOf("loction") >= 0 || t.indexOf("bookmrk") >= 0) {
    R.hits.push({ i: i, text: t });
  }
}
R.paraCount = D.Paragraphs.Count;
out.r = JSON.stringify(R);
return out;
