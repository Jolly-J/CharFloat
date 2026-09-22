const out = {};
const D = doc;
const R = {};
// 重建书签（放到文末空段，避免被目录域吃掉）
const p = D.Paragraphs.Add();
p.Range.Text = "附录 B · 书签与交叉引用复测\r";
const anchor = D.Paragraphs.Item(D.Paragraphs.Count).Range;
D.Bookmarks.Add("OAB_复测锚点", anchor);
R.bmAfterAdd = (() => { const a = []; for (let i = 1; i <= D.Bookmarks.Count; i++) a.push(D.Bookmarks.Item(i).Name); return a; })();
// 交叉引用
const r2 = D.Paragraphs.Item(D.Paragraphs.Count).Range;
r2.Collapse(0); r2.InsertAfter("\r");
const r3 = D.Paragraphs.Item(D.Paragraphs.Count).Range; r3.Collapse(1);
const f = D.Fields.Add(r3, 3, 'OAB_复测锚点 \\h ');
R.refResult = String(f.Result.Text).replace(/[\r\a\n]/g, "");
R.refCode = String(f.Code.Text).trim();
out.r = JSON.stringify(R);
return out;
