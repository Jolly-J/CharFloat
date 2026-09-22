const sheet = wb.Worksheets.Item(params.sheetName);
const out = [];
out.push('sheet=' + sheet.Name + ' pivotCount=' + sheet.PivotTables.Count);
try {
  const pt = sheet.PivotTables.Item(1);
  out.push('name=' + pt.Name + ' tableRange=' + pt.TableRange1.Address());
  out.push('rowFields=' + pt.RowFields.Count + ' colFields=' + pt.ColumnFields.Count + ' dataFields=' + pt.DataFields.Count);
  for (let i = 1; i <= pt.RowFields.Count; i++) out.push('row' + i + '=' + pt.RowFields.Item(i).Name);
  for (let i = 1; i <= pt.ColumnFields.Count; i++) out.push('col' + i + '=' + pt.ColumnFields.Item(i).Name);
  for (let i = 1; i <= pt.DataFields.Count; i++) {
    const df = pt.DataFields.Item(i);
    out.push('data' + i + '=' + df.Name + '|func=' + df.Function + '|src=' + df.SourceName);
  }
  try { out.push('srcData=' + pt.SourceData); } catch (e) { out.push('srcData ERR'); }
} catch (e) { out.push('PIVOT ERR ' + e.message); }
// 独立读回 B4:J14 值
const v = sheet.Range('B4:J14').Value2;
const rows = [];
for (let i = 0; i < v.length; i++) rows.push(String(v[i]).replace(/,/g, '|'));
return { dump: out.join(' ;; '), grid: rows.join(' // ') };
