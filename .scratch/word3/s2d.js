const out = {};
const R = {};
const D = doc;
const T = D.Tables.Item(2);

// 1) 合计行第 2 个单元格（ColumnIndex=2，实际是合并后的第 2 个格子）写公式
try {
  const rc = T.Rows.Item(7).Cells;
  const cell = rc.Item(2);
  R.cell2_colIndex = cell.ColumnIndex;
  R.cell2_before = String(cell.Range.Text).replace(/[\r\a]/g, "");
  cell.Formula("=SUM(ABOVE)", null, 0);
  R.cell2_after = String(rc.Item(2).Range.Text).replace(/[\r\a]/g, "");
  try { R.cell2_fieldCode = String(rc.Item(2).Range.Fields.Item(1).Code.Text); } catch (e) { R.cell2_fieldErr = String(e.message); }
  try { R.cell2_fieldResult = String(rc.Item(2).Range.Fields.Item(1).Result.Text); } catch (e) { R.cell2_resultErr = String(e.message); }
} catch (e) { R.formulaErr = String(e.message); }

// 2) AutoSum 显式测试：对 col5 数值列求和到一个新格
try {
  const asCell = T.Cell(6, 5);
  R.autosum_before = String(asCell.Range.Text).replace(/[\r\a]/g, "");
  const ok = asCell.AutoSum();
  R.autosum_return = String(ok);
  R.autosum_after = String(T.Cell(6, 5).Range.Text).replace(/[\r\a]/g, "");
} catch (e) { R.autosumErr = String(e.message); }

// 3) 行高 Exactly
try {
  T.Rows.Item(2).HeightRule = 2; T.Rows.Item(2).Height = 28;
  R.row2 = { height: T.Rows.Item(2).Height, rule: T.Rows.Item(2).HeightRule };
  T.Rows.Item(1).HeightRule = 1; T.Rows.Item(1).Height = 32;
  R.row1 = { height: T.Rows.Item(1).Height, rule: T.Rows.Item(1).HeightRule };
} catch (e) { R.rowHeightErr = String(e.message); }

// 4) 列宽
try {
  T.Columns.Item(1).Width = 100; T.Columns.Item(2).Width = 150;
  T.Columns.Item(3).Width = 80; T.Columns.Item(4).Width = 46; T.Columns.Item(5).Width = 70;
  R.widths = [1,2,3,4,5].map(i => T.Columns.Item(i).Width);
  R.tableWidth = (() => { try { return T.PreferredWidth; } catch (e) { return "ERR"; } })();
  R.columnsCount = T.Columns.Count;
} catch (e) { R.colWidthErr = String(e.message); }

// 5) 边框
try {
  T.Borders.Enable = true;
  T.Borders.OutsideLineStyle = 1; T.Borders.OutsideLineWidth = 12;
  T.Borders.InsideLineStyle = 1; T.Borders.InsideLineWidth = 4;
  R.borders = { outside: T.Borders.OutsideLineWidth, inside: T.Borders.InsideLineWidth };
} catch (e) { R.borderErr = String(e.message); }

// 6) 底纹：表头行 + 合计行
try {
  for (let c = 1; c <= 5; c++) {
    try { T.Rows.Item(2).Cells.Item(c).Shading.BackgroundPatternColor = 0xE8EEF7; } catch (e) {}
  }
  R.headerShading = T.Rows.Item(2).Cells.Item(1).Shading.BackgroundPatternColor;
  R.totalShading_before = (() => { try { return T.Rows.Item(7).Cells.Item(1).Shading.BackgroundPatternColor; } catch (e) { return "ERR"; } })();
  T.Rows.Item(7).Shading.BackgroundPatternColor = 0xDCE6F1;
  R.totalShading_after = T.Rows.Item(7).Cells.Item(1).Shading.BackgroundPatternColor;
} catch (e) { R.shadingErr = String(e.message); }

// 7) 表头跨页重复
try { T.Rows.Item(2).HeadingFormat = true; R.headingFormat = T.Rows.Item(2).HeadingFormat; } catch (e) { R.headingErr = String(e.message); }

// 8) 表格样式
try {
  R.styleBefore = T.Style.NameLocal;
  T.Style = "网格型";
  R.styleAfter = T.Style.NameLocal;
} catch (e) { R.styleErr = String(e.message); }

out.r = JSON.stringify(R);
return out;
