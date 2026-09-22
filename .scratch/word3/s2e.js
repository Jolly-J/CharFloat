const out = {};
const R = {};
const D = doc;
const T = D.Tables.Item(2);
const rc = T.Rows.Item(7).Cells;
const cell = rc.Item(2);

function tryCall(label, fn) {
  try { const v = fn(); R[label] = "OK:" + String(v); }
  catch (e) { R[label] = "ERR:" + String(e.message); }
}
tryCall("f1", () => cell.Formula("=SUM(ABOVE)"));
tryCall("f2", () => cell.Formula("=SUM(ABOVE)", 0));
tryCall("f3", () => cell.Formula("=SUM(ABOVE)", 0, "0"));
tryCall("f4", () => cell.Formula("=SUM(ABOVE)", "\\# 0.0"));
tryCall("f5", () => cell.Formula());
tryCall("f6", () => cell.Formula("=SUM(ABOVE)", true, 0));
R.after = String(rc.Item(2).Range.Text).replace(/[\r\a]/g, "");
try { R.fields = rc.Item(2).Range.Fields.Count; } catch (e) { R.fieldsErr = String(e.message); }
try { R.fieldType = rc.Item(2).Range.Fields.Item(1).Type; R.fieldCode = String(rc.Item(2).Range.Fields.Item(1).Code.Text); } catch (e) { R.fieldErr = String(e.message); }
// 换成真正的数值列求和：把第 5 列设为数值（100 / 0 / 99.4 / 96.8）
try {
  [100, 0, 99.4, 96.8].forEach((v, i) => { T.Cell(i + 3, 5).Range.Text = String(v); });
  R.col5 = [3,4,5,6].map(i => String(T.Cell(i, 5).Range.Text).replace(/[\r\a]/g, ""));
} catch (e) { R.numErr = String(e.message); }
out.r = JSON.stringify(R);
return out;
