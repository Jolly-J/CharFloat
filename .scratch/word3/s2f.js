const out = {};
const R = {};
const D = doc;
const T = D.Tables.Item(2);
const rc = T.Rows.Item(7).Cells;
const cell = rc.Item(2);

// 手动写公式域文本
try {
  cell.Range.Text = "=SUM(ABOVE)";
  R.afterWrite = String(cell.Range.Text).replace(/[\r\a]/g, "");
  R.fields = cell.Range.Fields.Count;
  if (R.fields > 0) {
    R.fieldType = cell.Range.Fields.Item(1).Type;
    R.fieldCode = String(cell.Range.Fields.Item(1).Code.Text);
    R.fieldResult = String(cell.Range.Fields.Item(1).Result.Text);
  }
} catch (e) { R.manualErr = String(e.message); }
// 强制更新域
try { cell.Range.Fields.Update(); R.afterUpdate = String(cell.Range.Text).replace(/[\r\a]/g, ""); } catch (e) { R.updateErr = String(e.message); }
// 直接写计算结果（人工公式）
try {
  const vals = [3,4,5,6].map(i => parseFloat(String(T.Cell(i,5).Range.Text).replace(/[^\d.]/g, "")));
  R.rawVals = JSON.stringify(vals);
  const sum = vals.reduce((a, b) => a + (isNaN(b) ? 0 : b), 0);
  R.computedSum = sum;
  cell.Range.Text = "合计 " + sum.toFixed(1);
  R.cellText = String(cell.Range.Text).replace(/[\r\a]/g, "");
  cell.Shading.BackgroundPatternColor = 0xDCE6F1;
  cell.Range.Font.Bold = true;
  R.cellShading = cell.Shading.BackgroundPatternColor;
} catch (e) { R.writeErr = String(e.message); }

// 表格样式（上一次未执行到）
try { R.styleBefore = T.Style.NameLocal; } catch (e) { R.styleErr = String(e.message); }
out.r = JSON.stringify(R);
return out;
