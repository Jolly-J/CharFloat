const sheet = wb.Worksheets.Item(params.sheetName);
const lines = [];
const A = params.addrs;
for (let i = 0; i < A.length; i++) {
  const r = sheet.Range(A[i]);
  let s = A[i] + ': rowH=' + r.RowHeight + ' colW=' + r.ColumnWidth + ' hidden=' + r.Hidden + ' rowsH=' + r.Rows.RowHeight + ' colsW=' + r.Columns.ColumnWidth;
  lines.push(s);
}
// 整行/整列隐藏状态
lines.push('row4hidden=' + sheet.Rows.Item(4).Hidden + ' row5hidden=' + sheet.Rows.Item(5).Hidden);
lines.push('colHhidden=' + sheet.Columns.Item('H').Hidden + ' colIhidden=' + sheet.Columns.Item('I').Hidden);
lines.push('usedRange=' + sheet.UsedRange.Address());
return { dump: lines.join(' ;; ') };
