const R = {};
const D = doc;
const T = D.Tables.Item(2);

// 1) 修复标题行（合并单元格内文字）
try {
  T.Cell(1, 1).Range.Text = "2.2 模块交付状态 · 高级表格实测汇总";
  T.Cell(1, 1).Shading.BackgroundPatternColor = 0xE8EEF7;
  T.Cell(1, 1).Range.Font.Bold = true;
  T.Cell(1, 1).Range.ParagraphFormat.Alignment = 1;
  R.title_cell = T.Cell(1, 1).Range.Text.replace(/[\r\a]/g, "");
} catch (e) { R.titleErr = String(e.message); }

// 2) 新列（列4）表头与说明
try { T.Cell(2, 4).Range.Text = "序号"; R.h_new = T.Cell(2, 4).Range.Text.replace(/[\r\a]/g, ""); } catch (e) { R.hErr = String(e.message); }
try {
  for (let r = 3; r <= 6; r++) { T.Cell(r, 4).Range.Text = String(r - 2); }
  R.seqCol = [1,2,3,4].map(i => T.Cell(i + 2, 4).Range.Text.replace(/[\r\a]/g, ""));
} catch (e) { R.seqErr = String(e.message); }

// 3) 表内公式：SUM(ABOVE) 对 C 列百分比求和会得 0，改用数值列。
//   先写入一列真实数值到 col5（"实测均值"）以便公式有可加对象
try {
  const nums = { 3: 100, 4: 0, 5: 99.4, 6: 96.8 };
  Object.keys(nums).forEach(function (k) { T.Cell(Number(k), 5).Range.Text = String(nums[k]); });
  R.numCol = [3,4,5,6].map(i => T.Cell(i, 5).Range.Text.replace(/[\r\a]/g, ""));
} catch (e) { R.numErr = String(e.message); }

// 4) 合计行 = ROW(7): 合并 cell(7,1..4) 已存在，输出到 cell(7,5)
try {
  const total = T.Cell(7, 5);
  R.totalCellFound = total !== null;
} catch (e) { R.totalCellErr = String(e.message); }
try {
  // 真实公式：对上方数值列求和
  const c = T.Cell(7, 5);
  c.Formula("=SUM(ABOVE)", null, 0);
  R.formulaText = String(T.Cell(7, 5).Range.Text).replace(/[\r\a]/g, "");
  R.formulaCode = (() => { try { return String(T.Cell(7,5).Range.Fields.Item(1).Code.Text); } catch (e) { return "noField:" + e.message; } })();
} catch (e) { R.formulaErr = String(e.message); }

// 5) AutoSum 单独测试
try {
  const c2 = T.Cell(6, 5);
  R.autosumLabel = String(c2.Range.Text).replace(/[\r\a]/g, "");
} catch (e) { R.autosumErr = String(e.message); }

// 6) 行高（Exactly）
try {
  T.Rows.Item(2).HeightRule = 2;   // wdRowHeightExactly
  T.Rows.Item(2).Height = 28;
  R.row2 = { height: T.Rows.Item(2).Height, rule: T.Rows.Item(2).HeightRule };
  T.Rows.Item(1).HeightRule = 1;   // AtLeast
  T.Rows.Item(1).Height = 32;
  R.row1 = { height: T.Rows.Item(1).Height, rule: T.Rows.Item(1).HeightRule };
} catch (e) { R.rowHeightErr = String(e.message); }

// 7) 列宽统一设置
try {
  T.Columns.Item(2).Width = 150;
  T.Columns.Item(3).Width = 120;
  T.Columns.Item(5).Width = 80;
  R.widths = [1,2,3,4,5].map(i => T.Columns.Item(i).Width);
  R.tableWidth = T.PreferredWidth;
  R.preferredType = T.PreferredWidthType;
} catch (e) { R.colWidthErr = String(e.message); }

// 8) 边框：整表外框 + 表头下框
try {
  T.Borders.Enable = true;
  T.Borders.Item(-3).LineWidth = 8;   // 上框
  T.Borders.Item(-1).LineWidth = 8;   // 下框
  T.Rows.Item(2).Borders.Item(-3).LineWidth = 12;
  R.borders = { outer: T.Borders.Item(-1).LineWidth, inner: T.Borders.Item(-5).LineStyle };
} catch (e) { R.borderErr = String(e.message); }

// 9) 表内编号（AutoNum 域）
try {
  const rng = T.Cell(6, 4).Range;
  rng.Text = "";
  D.Fields.Add(rng, -1, "AutoNum", false); // wdFieldAutoNum = -1?
  R.autonum = String(T.Cell(6, 4).Range.Text).replace(/[\r\a]/g, "");
} catch (e) { R.autonumErr = String(e.message); }

out.r = JSON.stringify(R);
return out;
