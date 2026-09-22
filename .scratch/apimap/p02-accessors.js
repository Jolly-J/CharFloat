// 访问形式探测：只读，不调用有副作用的成员
function T(label, fn) {
  try { const v = fn(); return label + ' => OK type=' + typeof v + (v && typeof v === 'object' && v.Name !== undefined ? ' name=' + v.Name : ''); }
  catch (e) { return label + ' => ERR ' + String(e && e.message ? e.message : e); }
}
const r = [];
// Excel 集合访问形式
r.push(T('wb.Worksheets(1)', function () { return wb.Worksheets(1); }));
r.push(T('wb.Worksheets.Item(1)', function () { return wb.Worksheets.Item(1); }));
r.push(T('wb.Worksheets[1]', function () { return wb.Worksheets[1]; }));
r.push(T('wb.Worksheets[0]', function () { return wb.Worksheets[0]; }));
r.push(T('wb.Sheets.Item(1)', function () { return wb.Sheets.Item(1); }));
r.push(T('wb.Worksheets.Count', function () { return wb.Worksheets.Count; }));
r.push(T('wb.Worksheets("行业概览")', function () { return wb.Worksheets('行业概览'); }));
r.push(T('wb.Worksheets.Item("行业概览")', function () { return wb.Worksheets.Item('行业概览'); }));
var ws = null;
try { ws = wb.Worksheets.Item(1); } catch (e) { ws = null; }
if (ws) {
  r.push(T('ws.Range("A1")', function () { return ws.Range('A1'); }));
  r.push(T('ws.Range("A1:B2")', function () { return ws.Range('A1:B2'); }));
  r.push(T('ws.Cells(1,1)', function () { return ws.Cells(1, 1); }));
  r.push(T('ws.Cells.Item(1,1)', function () { return ws.Cells.Item(1, 1); }));
  r.push(T('ws.Cells(1,1).Value2', function () { return ws.Cells(1, 1).Value2; }));
  r.push(T('ws.Cells.Item(1,1).Value2', function () { return ws.Cells.Item(1, 1).Value2; }));
  r.push(T('ws.Rows(1)', function () { return ws.Rows(1); }));
  r.push(T('ws.Rows.Item(1)', function () { return ws.Rows.Item(1); }));
  r.push(T('ws.Columns(1)', function () { return ws.Columns(1); }));
  r.push(T('ws.Columns.Item(1)', function () { return ws.Columns.Item(1); }));
  r.push(T('ws.UsedRange', function () { return ws.UsedRange; }));
  r.push(T('ws.Name', function () { return ws.Name; }));
  r.push(T('ws.Index', function () { return ws.Index; }));
  r.push(T('ws.Visible', function () { return ws.Visible; }));
  var rg = null;
  try { rg = ws.Range('A1:B2'); } catch (e) { rg = null; }
  if (rg) {
    r.push(T('rg.Value2', function () { return rg.Value2; }));
    r.push(T('rg.Formula', function () { return rg.Formula; }));
    r.push(T('rg.Address', function () { return rg.Address; }));
    r.push(T('rg.Row', function () { return rg.Row; }));
    r.push(T('rg.Column', function () { return rg.Column; }));
    r.push(T('rg.Rows.Count', function () { return rg.Rows.Count; }));
    r.push(T('rg.Cells.Item(1,1)', function () { return rg.Cells.Item(1, 1); }));
    r.push(T('rg.Item(1,1)', function () { return rg.Item(1, 1); }));
    r.push(T('rg.Areas.Count', function () { return rg.Areas.Count; }));
  }
}
// 全局变量可用性
r.push(T('typeof wps', function () { return typeof wps; }));
r.push(T('typeof doc', function () { return typeof doc; }));
r.push(T('typeof pres', function () { return typeof pres; }));
r.push(T('app.ActiveWorkbook.Name', function () { return app.ActiveWorkbook.Name; }));
r.push(T('app.ActiveSheet.Name', function () { return app.ActiveSheet.Name; }));
r.push(T('app.Workbooks.Count', function () { return app.Workbooks.Count; }));
r.push(T('app.Worksheets.Count', function () { return app.Worksheets.Count; }));
return r;
