const out = {};
const R = {};
const D = doc;
R.paraCount = D.Paragraphs.Count;
R.bmCount = D.Bookmarks.Count;
R.bmNames = (() => { const a = []; for (let i = 1; i <= D.Bookmarks.Count; i++) a.push(D.Bookmarks.Item(i).Name); return a; })();

// 正确写法：Fields.Add 会自动加域名前缀，code 只写参数部分
try {
  const r = D.Paragraphs.Item(D.Paragraphs.Count).Range;
  r.Collapse(0);
  r.InsertAfter("\r");
  const r2 = D.Paragraphs.Item(D.Paragraphs.Count).Range;
  r2.Collapse(1);
  const f = D.Fields.Add(r2, 3, 'OAB_验收结果 \\h ');
  R.refOk = { type: f.Type, code: String(f.Code.Text), result: String(f.Result.Text) };
  f.Update();
  R.refUpdated = String(f.Result.Text);
  R.refParaText = String(D.Paragraphs.Item(D.Paragraphs.Count).Range.Text).replace(/[\r\a]/g, "");
} catch (e) { R.refErr = String(e.message); }

try {
  const r = D.Paragraphs.Item(D.Paragraphs.Count).Range;
  r.Collapse(0); r.InsertAfter("\r");
  const r2 = D.Paragraphs.Item(D.Paragraphs.Count).Range;
  r2.Collapse(1);
  const f = D.Fields.Add(r2, 37, 'OAB_验收结果 \\h ');
  R.pageRefOk = { type: f.Type, code: String(f.Code.Text), result: String(f.Result.Text) };
  f.Update();
  R.pageRefUpdated = String(f.Result.Text);
  R.pageRefParaText = String(D.Paragraphs.Item(D.Paragraphs.Count).Range.Text).replace(/[\r\a]/g, "");
} catch (e) { R.pageRefErr = String(e.message); }

// 书签可被工具写入定位？先确认书签现存
R.bmStillThere = (() => { try { return D.Bookmarks.Exists("OAB_验收结果"); } catch (e) { return "ERR:" + e.message; } })();
out.r = JSON.stringify(R);
return out;
