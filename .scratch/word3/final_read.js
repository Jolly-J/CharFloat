const out = {};
const D = doc;
const R = {};
R.name = D.Name; R.saved = D.Saved; R.path = D.FullName;
R.paraCount = D.Paragraphs.Count;
try { R.pages = D.ComputeStatistics(2); } catch (e) { R.pages = "ERR"; }
try { R.words = D.ComputeStatistics(0); } catch (e) {}
R.props = { Title: String(D.BuiltInDocumentProperties.Item("Title").Value),
  Subject: String(D.BuiltInDocumentProperties.Item("Subject").Value),
  Keywords: String(D.BuiltInDocumentProperties.Item("Keywords").Value),
  Category: String(D.BuiltInDocumentProperties.Item("Category").Value) };
R.sections = [];
for (let i = 1; i <= D.Sections.Count; i++) {
  const s = D.Sections.Item(i);
  R.sections.push({ i: i, orient: s.PageSetup.Orientation, w: Math.round(s.PageSetup.PageWidth), h: Math.round(s.PageSetup.PageHeight),
    hdr: String(s.Headers.Item(1).Range.Text).replace(/[\r\a]/g, ""), ftr: String(s.Footers.Item(1).Range.Text).replace(/[\r\a]/g, "") });
}
R.customStyles = (() => { const a = []; for (let i = 1; i <= D.Styles.Count; i++) { const s = D.Styles.Item(i); if (!s.BuiltIn && s.InUse) a.push(s.NameLocal); } return a; })();
R.bookmarks = (() => { const a = []; for (let i = 1; i <= D.Bookmarks.Count; i++) a.push(D.Bookmarks.Item(i).Name); return a; })();
R.contentControls = D.ContentControls.Count;
R.comments = D.Comments.Count;
R.revisions = D.Revisions.Count;
R.tocCount = D.TablesOfContents.Count;
R.tables = [];
for (let i = 1; i <= D.Tables.Count; i++) { const T = D.Tables.Item(i); R.tables.push({ i: i, rows: T.Rows.Count, cols: T.Columns.Count, uniform: T.Uniform }); }
R.fields = D.Fields.Count;
R.refFields = (() => { const a = []; for (let i = 1; i <= D.Fields.Count; i++) { const t = D.Fields.Item(i).Type; if (t === 3 || t === 37) a.push({ type: t, code: String(D.Fields.Item(i).Code.Text).trim(), result: String(D.Fields.Item(i).Result.Text).slice(0, 24) }); } return a; })();
R.shapes = D.Shapes.Count;
out.r = JSON.stringify(R);
return out;
