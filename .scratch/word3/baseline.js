const out = {};
const D = doc;

out.document = { name: D.Name, fullName: D.FullName, saved: D.Saved };

// 文档属性（内置 + 自定义）
try {
  const bp = D.BuiltInDocumentProperties;
  out.builtinProps = {
    Title: bp.Item("Title").Value,
    Author: bp.Item("Author").Value,
    Subject: bp.Item("Subject").Value,
    Keywords: bp.Item("Keywords").Value,
    Comments: bp.Item("Comments").Value,
    Category: bp.Item("Category").Value
  };
} catch (e) { out.builtinPropsError = String(e.message); }
try {
  out.customProps = [];
  const cp = D.CustomDocumentProperties;
  for (let i = 1; i <= cp.Count; i++) {
    try { out.customProps.push({ name: cp.Item(i).Name, value: String(cp.Item(i).Value) }); } catch (e) {}
  }
} catch (e) { out.customPropsError = String(e.message); }

// 节
try {
  out.sections = [];
  for (let i = 1; i <= D.Sections.Count; i++) {
    const s = D.Sections.Item(i);
    const ps = s.PageSetup;
    out.sections.push({
      index: i, start: s.Range.Start, end: s.Range.End,
      orientation: ps.Orientation, pageWidth: ps.PageWidth, pageHeight: ps.PageHeight,
      headerDistance: ps.HeaderDistance, footerDistance: ps.FooterDistance,
      headerText: s.Headers.Item(1).Range.Text.replace(/[\r\a]/g, ""),
      footerText: s.Footers.Item(1).Range.Text.replace(/[\r\a]/g, "")
    });
  }
} catch (e) { out.sectionsError = String(e.message); }

// 样式清单
try {
  out.styleCount = D.Styles.Count;
  const names = [];
  for (let i = 1; i <= D.Styles.Count; i++) {
    const st = D.Styles.Item(i);
    let info = { i: i, name: st.NameLocal, type: st.Type, builtin: st.BuiltIn, inUse: st.InUse };
    try { info.font = st.Font.Name; info.size = st.Font.Size; info.bold = st.Font.Bold; } catch (e) {}
    names.push(info);
  }
  out.styles = names.filter(s => s.inUse).slice(0, 40);
  out.styleTotal = names.length;
} catch (e) { out.stylesError = String(e.message); }

// 书签
try {
  out.bookmarks = [];
  for (let i = 1; i <= D.Bookmarks.Count; i++) {
    out.bookmarks.push({ name: D.Bookmarks.Item(i).Name, start: D.Bookmarks.Item(i).Start, end: D.Bookmarks.Item(i).End });
  }
} catch (e) { out.bookmarksError = String(e.message); }

// 内容控件（ContentControls）
try {
  const ccs = D.ContentControls;
  out.contentControlCount = ccs.Count;
  out.contentControls = [];
  for (let i = 1; i <= ccs.Count; i++) {
    const c = ccs.Item(i);
    out.contentControls.push({ i: i, title: c.Title, tag: c.Tag, type: c.Type, text: String(c.Range.Text).slice(0, 40) });
  }
} catch (e) { out.contentControlsError = String(e.message); }

// 域（Fields / TableOfContents / Hyperlinks / CrossRef）
try {
  out.fields = [];
  for (let i = 1; i <= D.Fields.Count; i++) {
    out.fields.push({ i: i, type: D.Fields.Item(i).Type, code: String(D.Fields.Item(i).Code.Text).slice(0, 80) });
  }
} catch (e) { out.fieldsError = String(e.message); }
try { out.tocCount = D.TablesOfContents.Count; } catch (e) { out.tocError = String(e.message); }
try { out.hyperlinkCount = D.Hyperlinks.Count; } catch (e) { out.hyperlinksError = String(e.message); }
try { out.indexCount = D.Indexes.Count; } catch (e) { out.indexesError = String(e.message); }
try { out.shapesInline = D.InlineShapes.Count; out.shapesFloat = D.Shapes.Count; } catch (e) { out.shapesError = String(e.message); }
try { out.commentCount = D.Comments.Count; } catch (e) { out.commentsError = String(e.message); }
try { out.revisionCount = D.Revisions.Count; } catch (e) { out.revisionsError = String(e.message); }

// 表格
try {
  out.tables = [];
  for (let i = 1; i <= D.Tables.Count; i++) {
    const t = D.Tables.Item(i);
    out.tables.push({ i: i, rows: t.Rows.Count, cols: t.Columns.Count, uniform: (() => { try { return t.Uniform; } catch (e) { return null; } })(), style: (() => { try { return t.Style.NameLocal; } catch (e) { return null; } })(), rangeStart: t.Range.Start });
  }
} catch (e) { out.tablesError = String(e.message); }

return out;
