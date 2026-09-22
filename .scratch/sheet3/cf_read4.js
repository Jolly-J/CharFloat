const sheet = wb.Worksheets.Item(params.sheetName);
const lines = [];
const addrs = ['F4:F15','D4:D15','E4:E15'];
for (let k = 0; k < addrs.length; k++) {
  const a = addrs[k];
  const rng = sheet.Range(a);
  let n = -1;
  try {
    n = rng.FormatConditions.Count;
    for (let i = 1; i <= n; i++) {
      const fc = rng.FormatConditions.Item(i);
      const g = [];
      try { g.push('type=' + fc.Type); } catch (e) { g.push('type=?'); }
      try { g.push('op=' + fc.Operator); } catch (e) {}
      try { g.push('f1=' + fc.Formula1); } catch (e) {}
      try { g.push('priority=' + fc.Priority); } catch (e) {}
      try { g.push('iconID=' + fc.IconSet.ID); } catch (e) {}
      try { g.push('iconCount=' + fc.IconCriteria.Count); } catch (e) {}
      try { g.push('csCriteria=' + fc.ColorScaleCriteria.Count); } catch (e) {}
      try { g.push('font=' + fc.Font.Color); } catch (e) {}
      try { g.push('barColor=' + String(fc.BarColor)); } catch (e) {}
      lines.push(a + ' #' + i + ' ' + g.join(' '));
    }
    lines.push(a + ' COUNT=' + n);
  } catch (e) { lines.push(a + ' ERROR=' + e.message); }
}
return { total: sheet.Cells.FormatConditions.Count, dump: lines.join(' ;; ') };
