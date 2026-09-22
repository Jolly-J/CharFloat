const out = {};
const D = doc;
const want = params.want;
const stripped = want.replace(/a/g, "");
const found = { want: [], stripped: [] };
for (let i = 1; i <= D.Paragraphs.Count; i++) {
  const t = String(D.Paragraphs.Item(i).Range.Text).replace(/[\r\a\n]/g, "");
  if (t === want) found.want.push(i);
  if (t === stripped) found.stripped.push(i);
}
out.r = JSON.stringify({ sentByPython: want, exactMatch: found.want, strippedMatch: found.stripped,
  strippedForm: stripped, total: D.Paragraphs.Count });
return out;
