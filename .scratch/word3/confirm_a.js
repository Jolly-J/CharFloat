const out = {};
const D = doc;
const R = { cases: [] };
function one(label, text) {
  const endPos = D.Content.End - 1;
  D.Range(endPos, endPos).InsertAfter(text + "\r");
  const got = String(D.Paragraphs.Item(D.Paragraphs.Count).Range.Text).replace(/[\r\a\n]/g, "");
  R.cases.push({ label: label, sent: text, got: got, same: text === got });
}
one("lower-a", "a ab abc aaa banana");
one("upper-A", "A AB ABC AAA BANANA");
one("chinese-a", "啊阿呵");
one("mixed", "参数 location 与 bookmark");
one("no-a", "bcd efg hijk");
out.r = JSON.stringify(R);
return out;
