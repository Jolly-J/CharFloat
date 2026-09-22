const out = {};
const R = {};
const D = doc;
const T = D.Tables.Item(2);

// 先探测合并后的行实际单元格数
R.rowCells = {};
for (let r = 1; r <= 7; r++) {
  try { R.rowCells["r" + r] = T.Rows.Item(r).Cells.Count; } catch (e) { R.rowCells["r" + r] = "ERR"; }
}
// 合计行：找出可写的单元格
try {
  const rc = T.Rows.Item(7).Cells;
  R.totalRowCellTexts = [];
  for (let i = 1; i <= rc.Count; i++) {
    R.totalRowCellTexts.push({ i: i, col: rc.Item(i).ColumnIndex, text: String(rc.Item(i).Range.Text).replace(/[\r\a]/g, "") });
  }
} catch (e) { R.totalRowErr = String(e.message); }

out.r = JSON.stringify(R);
return out;
