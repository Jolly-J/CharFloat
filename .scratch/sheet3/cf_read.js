const sheet = wb.Worksheets.Item(params.sheetName);
const out = [];
const addrs = ['F4:F15','D4:D15','E4:E15'];
for (const a of addrs) {
  const rng = sheet.Range(a);
  let n = null, err = null, types = [], details = [];
  try {
    n = rng.FormatConditions.Count;
    for (let i = 1; i <= n; i++) {
      const fc = rng.FormatConditions.Item(i);
      let t = null, op = null, f1 = null, f2 = null, bar = null, cs = null, icon = null;
      try { t = fc.Type; } catch (e) { t = 'ERR:' + e.message; }
      try { op = fc.Operator; } catch (e) {}
      try { f1 = fc.Formula1; } catch (e) {}
      try { f2 = fc.Formula2; } catch (e) {}
      try { bar = fc.BarColor ? String(fc.BarColor) : null; } catch (e) {}
      try { cs = fc.ColorScaleCriteria ? fc.ColorScaleCriteria.Count : null; } catch (e) {}
      try { icon = fc.IconSet ? String(fc.IconSet.ID) : null; } catch (e) {}
      details.push({ idx: i, type: t, operator: op, formula1: f1, formula2: f2, barColor: bar, colorScaleCriteria: cs, iconSet: icon });
    }
  } catch (e) { err = e.message; }
  out.push({ address: a, count: n, error: err, details: details });
}
return { sheetName: sheet.Name, totalFormatConditions: sheet.Cells.FormatConditions.Count, ranges: out };
