const sheet = wb.Worksheets.Item(params.sheetName);
const pt = sheet.PivotTables(1);
const out = [];
out.push('name=' + pt.Name + ' range=' + pt.TableRange1.Address());
out.push('rowFields=' + pt.RowFields().Count + ' colFields=' + pt.ColumnFields().Count + ' dataFields=' + pt.DataFields().Count);
for (let i = 1; i <= pt.RowFields().Count; i++) out.push('row' + i + '=' + pt.RowFields().Item(i).Name);
for (let i = 1; i <= pt.ColumnFields().Count; i++) out.push('col' + i + '=' + pt.ColumnFields().Item(i).Name);
for (let i = 1; i <= pt.DataFields().Count; i++) {
  const d = pt.DataFields().Item(i);
  out.push('data' + i + '=' + d.Name + '|func=' + d.Function + '|srcName=' + d.SourceName + '|numFmt=' + d.NumberFormat);
}
out.push('sourceData=' + String(pt.SourceData));
out.push('pivotFieldNames=' + (function () { const n = []; for (let i = 1; i <= pt.PivotFields().Count; i++) n.push(pt.PivotFields().Item(i).Name); return n.join('/'); })());
return { dump: out.join(' ;; ') };
