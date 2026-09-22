const out = {};
const D = doc;
const r = {};
r.A3_builtin = { Title: String(D.BuiltInDocumentProperties.Item("Title").Value),
  Subject: String(D.BuiltInDocumentProperties.Item("Subject").Value),
  Keywords: String(D.BuiltInDocumentProperties.Item("Keywords").Value),
  Category: String(D.BuiltInDocumentProperties.Item("Category").Value),
  Comments: String(D.BuiltInDocumentProperties.Item("Comments").Value),
  Author: String(D.BuiltInDocumentProperties.Item("Author").Value) };
try { const s = D.Styles.Item("OAB 实测正文");
  r.B4_paraStyle = { type: s.Type, font: s.Font.Name, farEast: s.Font.NameFarEast, size: s.Font.Size,
    bold: s.Font.Bold, color: s.Font.Color, inUse: s.InUse, builtIn: s.BuiltIn,
    lineSpacing: s.ParagraphFormat.LineSpacing, lineRule: s.ParagraphFormat.LineSpacingRule,
    firstLineIndent: s.ParagraphFormat.FirstLineIndent }; } catch (e) { r.B4_paraStyle = "ERR " + e.message; }
try { const s = D.Styles.Item("OAB 强调标记");
  r.B4_charStyle = { type: s.Type, font: s.Font.Name, bold: s.Font.Bold, color: s.Font.Color, inUse: s.InUse }; } catch (e) { r.B4_charStyle = "ERR " + e.message; }
try { const p = D.Paragraphs.Item(23);
  r.B5_applyPara = { text: String(p.Range.Text).slice(0, 24), style: p.Style.NameLocal,
    font: p.Range.Font.Name, size: p.Range.Font.Size, bold: p.Range.Font.Bold }; } catch (e) { r.B5_applyPara = "ERR " + e.message; }
try { const p = D.Paragraphs.Item(24);
  r.B6_applyRange = { text: String(p.Range.Text).slice(0, 24), style: p.Style.NameLocal,
    font: p.Range.Font.Name, size: p.Range.Font.Size }; } catch (e) { r.B6_applyRange = "ERR " + e.message; }
try { const p = D.Paragraphs.Item(102);
  r.B7_heading = { text: String(p.Range.Text).slice(0, 24), style: p.Style.NameLocal, outline: p.OutlineLevel }; } catch (e) { r.B7_heading = "ERR " + e.message; }
try { const b = D.Styles.Item("正文"); r.B8_body = { font: b.Font.Name, size: b.Font.Size }; } catch (e) { r.B8_body = "ERR " + e.message; }
try { r.B9_inUseNow = (() => { const a = []; for (let i = 1; i <= D.Styles.Count; i++) { const s = D.Styles.Item(i); if (s.InUse && /OAB/.test(s.NameLocal)) a.push(s.NameLocal + "/type" + s.Type); } return a; })(); } catch (e) { r.B9_inUseNow = "ERR " + e.message; }
out.r = JSON.stringify(r);
return out;
