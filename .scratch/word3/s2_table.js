const out = {};
const D = doc;
const R = {};
const T = D.Tables.Item(2);

// 1) 表格追加前结构
R.before_rows = T.Rows.Count;
R.before_cols = T.Columns.Count;
R.before_uniform = T.Uniform;
R.before_cell11 = { w: T.Cell(1, 1).Width, shading: T.Cell(1, 1).Shading.BackgroundPatternColor };

// 2) 新增一列（列宽 / 行高基础）—— 在末尾插入一列承载"序号"
try {
  const lastCol = T.Columns.Count;
  T.Columns.Add(T.Columns.Item(lastCol));
  R.after_addColumn_cols = T.Columns.Count;
} catch (e) { R.addColumnErr = String(e.message); }

// 3) 追加 2 行
try { T.Rows.Add(); T.Rows.Add(); R.after_add2_rows = T.Rows.Count; } catch (e) { R.addRowErr = String(e.message); }

// 4) 写入数据（含数值列）
try {
  const fill = {
    6: ["连续写入成功率", "99.4%", "99.6%", "提升", ""],
    7: ["合计（公式）", "", "", "", ""]
  };
  Object.keys(fill).forEach(function (rk) {
    const rowArr = fill[rk];
    for (let c = 0; c < rowArr.length; c++) {
      try { T.Cell(Number(rk), c + 1).Range.Text = rowArr[c]; } catch (e) {}
    }
  });
  R.fill_ok = true;
} catch (e) { R.fillErr = String(e.message); }

// 5) 合并：标题行 [1,1]-[1,5]（转换为跨列标题）
try {
  T.Cell(1, 1).Merge(T.Cell(1, T.Columns.Count));
  R.merge_title_ok = true;
} catch (e) { R.merge_titleErr = String(e.message); }
R.after_merge_cols = T.Columns.Count;
R.after_merge_uniform = T.Uniform;
try { R.after_merge_row1_cells = T.Rows.Item(1).Cells.Count; } catch (e) { R.row1cellsErr = String(e.message); }

// 6) 合计行合并 [7,1]-[7,4] 后再放公式
try {
  T.Cell(7, 1).Merge(T.Cell(7, T.Columns.Count - 1));
  R.merge_total_ok = true;
} catch (e) { R.merge_totalErr = String(e.message); }

// 7) 单元格边框与底纹（脚本层）
try {
  const c = T.Cell(2, 1);
  c.Shading.BackgroundPatternColor = 0xF1F5F9; // BGR? 直接测试
  c.Borders.Item(-3).LineStyle = 1;
  c.Borders.Item(-3).LineWidth = 8;
  c.Range.Font.Bold = true;
  R.cell_shading_written = c.Shading.BackgroundPatternColor;
} catch (e) { R.cellShadingErr = String(e.message); }

// 8) 列宽 / 行高
try {
  T.Columns.Item(1).Width = 90;
  T.Rows.Item(2).Height = 30;
  T.Rows.Item(2).HeightRule = 0; // wdRowHeightAuto=0, AtLeast=1, Exactly=2
  R.col1_width = T.Columns.Item(1).Width;
  R.row2_height = T.Rows.Item(2).Height;
  R.row2_rule = T.Rows.Item(2).HeightRule;
} catch (e) { R.sizeErr = String(e.message); }

// 9) 表内公式
try {
  const cell = T.Cell(7, T.Columns.Count);
  cell.Formula("=SUM(ABOVE)", null, 0);
  R.formula_ok = true;
  R.formula_text = String(T.Cell(7, T.Columns.Count).Range.Text).replace(/[\r\a]/g, "");
} catch (e) { R.formulaErr = String(e.message); }
try { R.autosum_ok = T.Cell(6, T.Columns.Count).AutoSum(); } catch (e) { R.autosumErr = String(e.message); }

out.r = JSON.stringify(R);
return out;
