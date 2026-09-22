const sheet = wb.Worksheets.Item(params.sheetName);
const out = [];
// 清掉探测残留
sheet.Range('N17:O19').Clear();
if (sheet.AutoFilterMode) { sheet.AutoFilterMode = false; }
const rng = sheet.Range(params.range); // A4:G15
const before = String(rng.Value2);
try {
  const s = sheet.Sort;
  s.SortFields.Clear();
  // params.rules: [{col:1,order:1},{col:6,order:2}]  col 为区域内相对列号, order 1=升序 2=降序
  for (let i = 0; i < params.rules.length; i++) {
    s.SortFields.Add(rng.Columns.Item(params.rules[i].col), 0, params.rules[i].order);
  }
  s.SetRange(rng);
  s.Header = 2; // 1=有表头(首行不参与排序) 2=无表头
  s.Apply();
  out.push('applied');
} catch (e) { out.push('FAIL ' + e.message); }
const after = String(rng.Value2);
return { before: before, after: after, changed: before !== after, dump: out.join(' ;; ') };
