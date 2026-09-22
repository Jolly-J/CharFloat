const out = {};
const D = doc;
const R = { cases: [] };
function rec(label, fn) {
  try { R.cases.push({ label: label, v: fn() }); }
  catch (e) { R.cases.push({ label: label, err: String(e.message) }); }
}
const S = "Xa a ab abc banana A Aa 啊阿";

// 1) Paragraphs.Add(targetRange) —— 与 word.js 工具完全一致的写法
rec("Paragraphs.Add(targetRange)+Range.Text", () => {
  const endPos = D.Content.End > 1 ? D.Content.End - 1 : 0;
  const tr = D.Range(endPos, endPos);
  const p = D.Paragraphs.Add(tr);
  p.Range.Text = S + "\n";
  return String(p.Range.Text).replace(/[\r\a\n]/g, "");
});
// 2) 无参 Add + Range.Text
rec("Paragraphs.Add()+Range.Text", () => {
  const p = D.Paragraphs.Add();
  p.Range.Text = S + "\n";
  return String(p.Range.Text).replace(/[\r\a\n]/g, "");
});
// 3) Paragraphs.Add(targetRange) + InsertAfter
rec("Paragraphs.Add(targetRange)+InsertAfter", () => {
  const endPos = D.Content.End > 1 ? D.Content.End - 1 : 0;
  const p = D.Paragraphs.Add(D.Range(endPos, endPos));
  p.Range.InsertAfter(S + "\n");
  return String(p.Range.Text).replace(/[\r\a\n]/g, "");
});
// 4) Content 末尾 Range.InsertAfter
rec("ContentEnd Range.InsertAfter", () => {
  const endPos = D.Content.End > 1 ? D.Content.End - 1 : 0;
  D.Range(endPos, endPos).InsertAfter(S + "\n");
  return String(D.Paragraphs.Item(D.Paragraphs.Count - 1).Range.Text).replace(/[\r\a\n]/g, "");
});
out.r = JSON.stringify(R);
return out;
