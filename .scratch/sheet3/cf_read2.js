const sheet = wb.Worksheets.Item(params.sheetName);
const parts = [];
const addrs = ['F4:F15','D4:D15','E4:E15'];
for (const a of addrs) {
  const rng = sheet.Range(a);
  const items = [];
  let n = -1, err = null;
  try {
    n = rng.FormatConditions.Count;
    for (let i = 1; i <= n; i++) {
      const fc = rng.FormatConditions.Item(i);
      const o = {};
      try { o.type = fc.Type; } catch (e) { o.type = 'ERR'; }
      try { o.operator = fc.Operator; } catch (e) {}
      try { o.formula1 = fc.Formula1; } catch (e) {}
      try { o.formula2 = fc.Formula2; } catch (e) {}
      try { o.barColor = fc.BarColor; } catch (e) {}
      try { o.colorScaleCriteria = fc.ColorScaleCriteria.Count; } catch (e) {}
      try { o.iconSet = fc.IconSet.ID; } catch (e) {}
      items.push(o);
    }
  } catch (e) { err = e.message; }
  parts.push(a + ' => count=' + n + ' err=' + err + ' :: ' + JSON.stringify(items));
}
return { sheet: sheet.Name, total: sheet.Cells.FormatConditions.Count, dump: parts.join(' | ') };
