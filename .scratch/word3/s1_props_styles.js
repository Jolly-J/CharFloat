const out = { steps: [], failed: null };
const D = doc;
function step(id, fn) {
  try { const v = fn(); out.steps.push({ id: id, ok: true, v: v }); }
  catch (e) { out.steps.push({ id: id, ok: false, err: String(e.message || e) }); }
}

// A. 文档属性：内置属性（通过工具改写前的原生探测）
step("A1_read_title_before", () => String(D.BuiltInDocumentProperties.Item("Title").Value));
step("A2_can_write_builtin", () => {
  D.BuiltInDocumentProperties.Item("Title").Value = "Office Agent Bridge 2.1 版本发布与验收说明";
  D.BuiltInDocumentProperties.Item("Subject").Value = "2.1.0 P0/P1 阶段验收材料";
  D.BuiltInDocumentProperties.Item("Keywords").Value = "WPS;Office Agent Bridge;验收;MCP";
  D.BuiltInDocumentProperties.Item("Category").Value = "阶段验收";
  D.BuiltInDocumentProperties.Item("Comments").Value = "由文档高级能力实测生成";
  return "written";
});
step("A3_readback_builtin", () => ({
  Title: String(D.BuiltInDocumentProperties.Item("Title").Value),
  Subject: String(D.BuiltInDocumentProperties.Item("Subject").Value),
  Keywords: String(D.BuiltInDocumentProperties.Item("Keywords").Value),
  Category: String(D.BuiltInDocumentProperties.Item("Category").Value),
  Author: String(D.BuiltInDocumentProperties.Item("Author").Value)
}));
// 自定义属性
step("A4_add_custom_prop", () => {
  const cp = D.CustomDocumentProperties;
  try { cp.Item("验收批次").Delete(); } catch (e) {}
  cp.Add("验收批次", false, 4, "2.1.0-p0p1-p5");
  return "added";
});
step("A5_readback_custom", () => {
  const cp = D.CustomDocumentProperties; const arr = [];
  for (let i = 1; i <= cp.Count; i++) { try { arr.push(cp.Item(i).Name + "=" + cp.Item(i).Value); } catch (e) {} }
  return arr;
});

// B. 样式体系
step("B1_existing_styles", () => {
  const arr = [];
  for (let i = 1; i <= D.Styles.Count; i++) {
    const s = D.Styles.Item(i);
    if (s.InUse) arr.push(s.NameLocal + "(type=" + s.Type + ")");
  }
  return arr;
});
step("B2_add_paragraph_style", () => {
  try { D.Styles.Item("OAB 实测正文").Delete(); } catch (e) {}
  const st = D.Styles.Add("OAB 实测正文", 1); // wdStyleTypeParagraph
  st.Font.Name = "微软雅黑"; st.Font.NameFarEast = "微软雅黑"; st.Font.Size = 12;
  st.Font.Bold = false; st.Font.Color = 0x404040;
  st.ParagraphFormat.FirstLineIndent = D.Styles.Item("正文").ParagraphFormat.FirstLineIndent || 21;
  st.ParagraphFormat.LineSpacingRule = 4; // wdLineSpaceExactly
  st.ParagraphFormat.LineSpacing = 20;
  st.ParagraphFormat.SpaceAfter = 6;
  return { name: st.NameLocal, existsNow: true };
});
step("B3_add_character_style", () => {
  try { D.Styles.Item("OAB 强调标记").Delete(); } catch (e) {}
  const st = D.Styles.Add("OAB 强调标记", 2); // wdStyleTypeCharacter
  st.Font.Name = "微软雅黑"; st.Font.Bold = true; st.Font.Color = 0xC00000;
  return { name: st.NameLocal, type: st.Type, size: st.Font.Size };
});
step("B4_readback_new_styles", () => {
  const r = {};
  ["OAB 实测正文", "OAB 强调标记"].forEach(n => {
    try {
      const s = D.Styles.Item(n);
      r[n] = { type: s.Type, font: s.Font.Name, size: s.Font.Size, bold: s.Font.Bold,
               inUse: s.InUse, builtIn: s.BuiltIn };
    } catch (e) { r[n] = "ERROR " + e.message; }
  });
  return r;
});
step("B5_apply_para_style", () => {
  // 应用到第 22 段（1.1 编写目的 之后的正文段落，按当前索引取）
  const p = D.Paragraphs.Item(23);
  const before = { style: p.Style.NameLocal, text: String(p.Range.Text).slice(0, 30) };
  p.Style = D.Styles.Item("OAB 实测正文");
  return { before: before, after: p.Style.NameLocal, font: p.Range.Font.Name, size: p.Range.Font.Size };
});
step("B6_apply_style_to_range", () => {
  const p = D.Paragraphs.Item(24);
  p.Range.Style = D.Styles.Item("OAB 实测正文");
  return { idx: 24, style: p.Style.NameLocal, text: String(p.Range.Text).slice(0, 30) };
});
step("B7_apply_heading_style", () => {
  const p = D.Paragraphs.Item(102); // 后期核对
  const before = { style: p.Style.NameLocal, text: String(p.Range.Text).slice(0, 30), outline: p.OutlineLevel };
  p.Style = D.Styles.Item("标题 3");
  return { before: before, after: p.Style.NameLocal, outline: p.OutlineLevel };
});
// 正文样式基样式修改（Body Text 改字号，验证样式继承）
step("B8_body_text_style_change", () => {
  const body = D.Styles.Item("正文");
  const before = { font: body.Font.Name, size: body.Font.Size };
  body.Font.Name = "宋体"; body.Font.NameFarEast = "宋体"; body.Font.Size = 12;
  return { before: before, after: { font: body.Font.Name, size: body.Font.Size } };
});

return out;
