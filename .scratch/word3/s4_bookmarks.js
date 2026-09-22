const out = {};
const R = {};
const D = doc;
R.contentControlsApi = (typeof D.ContentControls);

// 1) 建书签：取「三、验收测试结果」标题段
let targetText = "";
try {
  let idx = 0;
  for (let i = 1; i <= D.Paragraphs.Count; i++) {
    const t = String(D.Paragraphs.Item(i).Range.Text);
    if (t.indexOf("三、验收测试结果") >= 0) { idx = i; break; }
  }
  R.headingParaIndex = idx;
  const rng = D.Paragraphs.Item(idx).Range;
  targetText = String(rng.Text).replace(/[\r\a]/g, "");
  D.Bookmarks.Add("OAB_验收结果", rng);
  R.bookmarkAdded = true;
} catch (e) { R.bookmarkAddErr = String(e.message); }
R.headingText = targetText;

// 2) 读回书签
try {
  R.bookmarkCount = D.Bookmarks.Count;
  R.bookmarkNames = [];
  for (let i = 1; i <= D.Bookmarks.Count; i++) R.bookmarkNames.push(D.Bookmarks.Item(i).Name);
  const bm = D.Bookmarks.Item("OAB_验收结果");
  R.bmRange = { start: bm.Start, end: bm.End, text: String(bm.Range.Text).replace(/[\r\a]/g, "") };
} catch (e) { R.bookmarkReadErr = String(e.message); }

// 3) 第二个书签：表格 2 之前的段落
try {
  let idx2 = 0;
  for (let i = 1; i <= D.Paragraphs.Count; i++) {
    if (String(D.Paragraphs.Item(i).Range.Text).indexOf("2.2 模块交付状态") >= 0) { idx2 = i; break; }
  }
  R.paraIndex2 = idx2;
  D.Bookmarks.Add("OAB_交付状态", D.Paragraphs.Item(idx2).Range);
  R.bookmark2Text = String(D.Paragraphs.Item(idx2).Range.Text).replace(/[\r\a]/g, "");
} catch (e) { R.bookmark2Err = String(e.message); }

// 4) 内容控件尝试
try {
  const cc = D.ContentControls.Add(0, D.Paragraphs.Item(3).Range); // wdContentControlRichText = 0
  R.ccAdded = true;
  R.ccCount = D.ContentControls.Count;
  cc.Title = "OAB 实测控件"; cc.Tag = "oab-test";
  R.ccReadback = { count: D.ContentControls.Count, title: D.ContentControls.Item(1).Title,
    tag: D.ContentControls.Item(1).Tag, type: D.ContentControls.Item(1).Type };
} catch (e) { R.ccErr = String(e.message); }

// 5) 交叉引用：REF 域引用书签
try {
  const rng = D.Paragraphs.Item(D.Paragraphs.Count).Range;
  rng.Collapse(0); // wdCollapseEnd
  rng.InsertAfter("（交叉引用测试）\r");
  const r2 = D.Paragraphs.Item(D.Paragraphs.Count).Range;
  r2.Collapse(1); // 折叠到起点
  const f = D.Fields.Add(r2, 3, ' REF OAB_验收结果 \\h '); // wdFieldRef = 3
  R.refField = { type: f.Type, code: String(f.Code.Text), result: String(f.Result.Text) };
  R.refFieldText = String(r2.Text).replace(/[\r\a]/g, "");
} catch (e) { R.refErr = String(e.message); }
// 6) PAGEREF 域引用书签
try {
  const r3 = D.Paragraphs.Item(D.Paragraphs.Count).Range;
  r3.InsertAfter("（页码引用：" );
  const r4 = D.Paragraphs.Item(D.Paragraphs.Count).Range;
  r4.Collapse(1);
  const fp = D.Fields.Add(r4, 37, ' PAGEREF OAB_验收结果 \\h '); // wdFieldPageRef
  R.pageRef = { type: fp.Type, code: String(fp.Code.Text), result: String(fp.Result.Text) };
  R.pageRefText = String(r4.Text).replace(/[\r\a]/g, "");
} catch (e) { R.pageRefErr = String(e.message); }
out.r = JSON.stringify(R);
return out;
