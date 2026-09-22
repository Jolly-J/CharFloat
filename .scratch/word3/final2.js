const out = {};
const D = doc;
const R = {};
R.saved = D.Saved;
R.paraCount = D.Paragraphs.Count;
R.pages = D.ComputeStatistics(2);
R.words = D.ComputeStatistics(0);
R.bms = (() => { const a = []; for (let i = 1; i <= D.Bookmarks.Count; i++) a.push(D.Bookmarks.Item(i).Name); return a; })();
R.cc = D.ContentControls.Count;
R.comments = D.Comments.Count;
R.revisions = D.Revisions.Count;
R.tocCount = D.TablesOfContents.Count;
R.sections = D.Sections.Count;
R.tables = [];
for (let i = 1; i <= D.Tables.Count; i++) { const T = D.Tables.Item(i); R.tables.push(T.Rows.Count + "x" + T.Columns.Count + "/uniform=" + T.Uniform); }
R.customStyles = (() => { const a = []; for (let i = 1; i <= D.Styles.Count; i++) { const s = D.Styles.Item(i); if (!s.BuiltIn && s.InUse) a.push(s.NameLocal); } return a; })();
// 高级表关键格
try { R.tbl2 = { title: String(D.Tables.Item(2).Cell(1,1).Range.Text).replace(/[\r\a\n]/g, ""),
  total: String(D.Tables.Item(2).Rows.Item(7).Cells.Item(2).Range.Text).replace(/[\r\a\n]/g, ""),
  col1w: D.Tables.Item(2).Columns.Item(1).Width, row2h: D.Tables.Item(2).Rows.Item(2).Height,
  row2rule: D.Tables.Item(2).Rows.Item(2).HeightRule, headShade: D.Tables.Item(2).Rows.Item(2).Cells.Item(1).Shading.BackgroundPatternColor }; } catch (e) { R.tblErr = String(e.message); }
out.r = JSON.stringify(R);
return out;
