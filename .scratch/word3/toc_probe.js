const out = {};
const D = doc;
const R = {};
try {
  R.tocCount = D.TablesOfContents.Count;
  for (let i = 1; i <= D.TablesOfContents.Count; i++) {
    const t = D.TablesOfContents.Item(i);
    R["toc" + i] = { start: t.Range.Start, end: t.Range.End,
      entries: t.Range.Paragraphs.Count,
      head: String(t.Range.Text).slice(0, 120).replace(/[\r\a\t]/g, "|") };
  }
} catch (e) { R.tocErr = String(e.message); }
try { R.fields = D.Fields.Count; } catch (e) { R.fErr = String(e.message); }
R.paraCount = D.Paragraphs.Count;
out.r = JSON.stringify(R);
return out;
