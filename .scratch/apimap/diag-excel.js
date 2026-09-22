// 逐个复核静默失败的调用，并立刻读回
const out = [];
function step(label, fn) {
  let v, err = null;
  try { v = fn(); } catch (e) { err = String(e && e.message ? e.message : e); }
  out.push(label + ' || ret=' + (err ? 'THROW(' + err + ')' : (v === null ? 'null' : v === undefined ? 'undefined' : typeof v === 'object' ? 'obj' : String(v))));
  return v;
}
const B = wb.Worksheets.Item('Data');
const P = wb.Worksheets.Item('Probe');
out.push('P.Name=' + P.Name + ' P.Index=' + P.Index + ' B.Name=' + B.Name);

// 形状
var sh = step('P.Shapes.AddShape(1,20,20,160,80)', function () { return P.Shapes.AddShape(1, 20, 20, 160, 80); });
out.push('  -> P.Shapes.Count=' + P.Shapes.Count);
var sh2 = step('P.Shapes.AddShape(msoShapeRectangle=1,...) 常量1 check', function () { return P.Shapes.AddShape(1, 300, 300, 50, 50); });
out.push('  -> P.Shapes.Count=' + P.Shapes.Count);
// 用命名常量
var sh3 = step('P.Shapes.AddShape(5,...)', function () { return P.Shapes.AddShape(5, 20, 200, 100, 50); });
out.push('  -> P.Shapes.Count=' + P.Shapes.Count);

// 图表
var co = step('P.ChartObjects().Add(260,20,380,240)', function () { return P.ChartObjects().Add(260, 20, 380, 240); });
out.push('  -> P.ChartObjects().Count=' + P.ChartObjects().Count);
var co2 = step('P.Shapes.AddChart2(-1,5,260,20,380,240)', function () { return P.Shapes.AddChart2(-1, 5, 260, 20, 380, 240); });
out.push('  -> P.Shapes.Count=' + P.Shapes.Count + ' ChartObjects=' + P.ChartObjects().Count);

// 条件格式
var fc = step('B.Range("D2:D11").FormatConditions.AddColorScale(3)', function () { return B.Range('D2:D11').FormatConditions.AddColorScale(3); });
out.push('  -> FormatConditions.Count=' + B.Range('D2:D11').FormatConditions.Count);
var fc2 = step('B.Range("D2:D11").FormatConditions.Add(1,3,100)', function () { return B.Range('D2:D11').FormatConditions.Add(1, 3, 100); });
out.push('  -> FormatConditions.Count=' + B.Range('D2:D11').FormatConditions.Count);
var fc3 = step('B.Range("D2:D11").FormatConditions.AddDatabar()', function () { return B.Range('D2:D11').FormatConditions.AddDatabar(); });
out.push('  -> FormatConditions.Count=' + B.Range('D2:D11').FormatConditions.Count);
var fc4 = step('B.Range("D2:D11").FormatConditions.AddIconSetCondition()', function () { return B.Range('D2:D11').FormatConditions.AddIconSetCondition(); });
out.push('  -> FormatConditions.Count=' + B.Range('D2:D11').FormatConditions.Count);
out.push('  -> typeof AddColorScale=' + typeof B.Range('D2:D11').FormatConditions.AddColorScale + ' typeof Add=' + typeof B.Range('D2:D11').FormatConditions.Add);

// 透视表
out.push('typeof wb.PivotCaches=' + typeof wb.PivotCaches);
var wp = null; try { wp = wb.PivotCaches; } catch (e) { out.push('wb.PivotCaches ERR ' + e); }
if (wp) { out.push('wb.PivotCaches type=' + typeof wp + ' count=' + (function () { try { return wp.Count; } catch (e) { return 'ERR'; } })()); }
var pcs = step('wb.PivotCaches()', function () { return wb.PivotCaches(); });
if (pcs) {
  out.push('  PivotCaches().Count=' + (function () { try { return pcs.Count; } catch (e) { return 'ERR ' + e; } })());
  var pcc = step('PivotCaches.Create(1,data,6)', function () { return pcs.Create(1, B.Range('A1:G11'), 6); });
  out.push('  -> PivotCaches().Count=' + (function () { try { return pcs.Count; } catch (e) { return 'ERR'; } })());
  if (pcc) { out.push('  pc type=' + typeof pcc); }
}
out.push('typeof B.PivotTables=' + typeof B.PivotTables + ' typeof B.PivotTableWizard=' + typeof B.PivotTableWizard);

// 验证
var va = step('B.Range("B2:B11").Validation.Add(3,1,1,"A,B,C")', function () { return B.Range('B2:B11').Validation.Add(3, 1, 1, 'A,B,C'); });
out.push('  -> Validation.Type=' + (function () { try { return B.Range('B2:B11').Validation.Type; } catch (e) { return 'ERR ' + e; } })());

wb.Save();
return out;
