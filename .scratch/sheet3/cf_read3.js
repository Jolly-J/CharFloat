const sheet = wb.Worksheets.Item(params.sheetName);
const lines = [];
const addrs = ['F4:F15','D4:D15','E4:E15'];
for (let k = 0; k < addrs.length; k++) {
  const a = addrs[k];
  const rng = sheet.Range(a);
  let n = -1, err = '';
  try {
    n = rng.FormatConditions.Count;
    for (let i = 1; i <= n; i++) {
      const fc = rng.FormatConditions.Item(i);
      const g = {};
      try { g.type = String(fc.Type); } catch (e) { g.type = 'ERR'; }
      try { g.op = String(fc.Operator); } catch (e) { g.op = '-'; }
      try { g.f1 = String(fc.Formula1); } catch (e) { g.f1 = '-'; }
      try { g.f2 = String(fc.Formula2); } catch (e) { g.f2 = '-'; }
      try { g.bar = String(fc.BarColor); } catch (e) { g.bar = '-'; }
      try { g.csN = String(fc.ColorScaleCriteria.Count); } catch (e) { g.csN = '-'; }
      try { g.icon = String(fc.IconSet.ID); } catch (e) { g.icon = '-'; }
      try { g.interior = String(fc.Interior.Color); } catch (e) { g.interior = '-'; }
      lines.push(a + ' #' + i + ' type=' + g.type + ' op=' + g.op + ' f1=' + g.f1 + ' f2=' + g.f2 + ' bar=' + g.bar + ' csN=' + g.csN + ' icon=' + g.icon + ' interior=' + g.interior);
    }
    lines.push(a + ' COUNT=' + n);
  } catch (e) { lines.push(a + ' ERROR=' + e.message); }
}
return { sheet: sheet.Name, total: sheet.Cells.FormatConditions.Count, dump: lines.join(' ;; ') };
