const out = {};
const D = doc;
const R = { tocs: [] };
for (let i = 1; i <= D.TablesOfContents.Count; i++) {
  const t = D.TablesOfContents.Item(i);
  let before = String(t.Range.Text).slice(0, 60).replace(/[\r\a\t]/g, "|");
  let updated = "n/a";
  try { t.Update(); updated = "ok"; } catch (e) { updated = "ERR:" + e.message; }
  R.tocs.push({ i: i, start: t.Range.Start, end: t.Range.End, paras: t.Range.Paragraphs.Count,
    before: before, update: updated, after: String(t.Range.Text).slice(0, 60).replace(/[\r\a\t]/g, "|"),
    upper: (() => { try { return t.UpperHeadingLevel; } catch (e) { return "ERR"; } })(),
    lower: (() => { try { return t.LowerHeadingLevel; } catch (e) { return "ERR"; } })() });
}
out.r = JSON.stringify(R);
return out;
