// 创建 apimap 探测用 Excel 工作簿（只写自己的新文件）
const path = '/Users/jolin/Downloads/apimap-scratch.xlsx';
const nb = app.Workbooks.Add();
const ws = nb.Worksheets.Item(1);
ws.Name = 'Data';
const rows = [
  ['地区', '产品', '季度', '销量', '单价', '销售额', '日期'],
  ['华东', 'A', 'Q1', 120, 30.5, null, '2024-01-15'],
  ['华东', 'B', 'Q1', 80, 45, null, '2024-02-20'],
  ['华南', 'A', 'Q2', 200, 30.5, null, '2024-04-10'],
  ['华南', 'C', 'Q2', 60, 99.9, null, '2024-05-05'],
  ['华北', 'B', 'Q3', 310, 45, null, '2024-07-18'],
  ['华北', 'A', 'Q3', 150, 30.5, null, '2024-08-22'],
  ['西南', 'C', 'Q4', 90, 99.9, null, '2024-10-30'],
  ['西南', 'B', 'Q4', 240, 45, null, '2024-11-11'],
  ['东北', 'A', 'Q4', 175, 30.5, null, '2024-12-01'],
  ['东北', 'C', 'Q1', 65, 99.9, null, '2024-03-14']
];
for (let i = 0; i < rows.length; i++) {
  for (let j = 0; j < rows[i].length; j++) {
    if (rows[i][j] !== null) ws.Cells.Item(i + 1, j + 1).Value2 = rows[i][j];
  }
}
ws.Range('F2').Formula = '=D2*E2';
ws.Range('F2').AutoFill(ws.Range('F2:F11'));
ws.Columns.Item(1).ColumnWidth = 12;
const ws2 = nb.Worksheets.Add(null, ws);
ws2.Name = 'Probe';
nb.SaveAs(path);
return {
  saved: path,
  sheets: nb.Worksheets.Count,
  firstName: nb.Worksheets.Item(1).Name,
  used: ws.UsedRange.Address(),
  f2: ws.Range('F2').Formula,
  f11: ws.Range('F11').Value2
};
