// 在 apimap-scratch.xlsx 中造出各类对象实例（只操作自己的探测文件）
const log = [];
function T(label, fn) { try { const v = fn(); log.push(label + ' => OK ' + (typeof v === 'object' && v !== null ? 'obj' : String(v))); return v; } catch (e) { log.push(label + ' => ERR ' + String(e && e.message ? e.message : e)); return null; } }
const B = wb.Worksheets.Item('Data');
const P = wb.Worksheets.Item('Probe');
const data = B.Range('A1:G11');

// --- ListObject（表格对象）---
T('ListObjects.Add', function () { return B.ListObjects.Add(1, data, null, 1); });

// --- 条件格式 ---
T('FormatConditions.AddColorScale', function () { return B.Range('D2:D11').FormatConditions.AddColorScale(3); });
T('FormatConditions.Add(Type=1)', function () { return B.Range('E2:E11').FormatConditions.Add(1, 3, 100); });
T('FormatConditions.AddDatabar', function () { return B.Range('D2:D11').FormatConditions.AddDatabar(); });

// --- 数据验证 ---
T('Validation.Add', function () { return B.Range('B2:B11').Validation.Add(3, 1, 1, 'A,B,C'); });

// --- 批注（传统 + 线程）---
T('Range.AddComment', function () { return B.Range('A1').AddComment('apimap 传统批注'); });
T('Range.AddCommentThreaded', function () { return B.Range('A2').AddCommentThreaded('apimap 线程批注'); });

// --- 自动筛选 ---
T('Range.AutoFilter', function () { return data.AutoFilter(1, '华东'); });
T('ShowAllData', function () { return B.ShowAllData(); });

// --- 排序对象 ---
T('B.Sort.SortFields', function () { return B.Sort.SortFields; });

// --- 冻结 / 拆分 / 缩放 ---
T('FreezePanes', function () { P.Activate(); app.ActiveWindow.FreezePanes = false; app.ActiveWindow.SplitRow = 2; app.ActiveWindow.FreezePanes = true; return 'frozen'; });
T('Zoom', function () { return app.ActiveWindow.Zoom; });

// --- 形状 + 图表 ---
T('Shapes.AddShape', function () { return P.Shapes.AddShape(1, 20, 20, 160, 80); });
T('Shapes.AddTextbox', function () { return P.Shapes.AddTextbox(1, 20, 130, 220, 60); });
T('ChartObjects.Add', function () { return P.ChartObjects().Add(260, 20, 380, 240); });

// --- 命名区域 ---
T('wb.Names.Add', function () { return wb.Names.Add('apimapSales', '=Data!$D$2:$D$11'); });
T('B.Names.Add', function () { return B.Names.Add('apimapLocal', '=Data!$E$2:$E$11'); });

// --- 透视表 ---
var pc = T('PivotCaches.Create', function () { return wb.PivotCaches().Create(1, data, 6); });
if (pc) T('PivotTableWizard/CreatePivotTable', function () { return pc.CreatePivotTable(P.Range('A1'), 'apimapPivot'); });

// --- 工作表视图 ---
T('NamedSheetViews.Add', function () { return B.NamedSheetViews.Add('apimapView'); });

// --- 超链接 ---
T('Hyperlinks.Add', function () { return P.Hyperlinks.Add(P.Range('J1'), 'https://www.wps.com', null, 'apimapLink', 'WPS'); });

// --- 分页/打印 ---
T('PageSetup.Orientation', function () { return B.PageSetup.Orientation; });
T('PageSetup.PrintArea', function () { B.PageSetup.PrintArea = '$A$1:$G$11'; return B.PageSetup.PrintArea; });

// --- 大纲/分组 ---
T('Outline', function () { return B.Outline.SummaryRow; });
T('Rows.Group', function () { return B.Rows.Item('2:3').Group(); });

// --- 窗口 ---
T('app.Windows.Count', function () { return app.Windows.Count; });
T('app.Windows.Item(1).Caption', function () { return app.Windows.Item(1).Caption; });

// --- 连接 ---
T('wb.Connections', function () { return wb.Connections.Count; });

wb.Save();
return { log: log, listObjects: B.ListObjects.Count, names: wb.Names.Count, sheetNames: B.Names.Count, pivots: P.PivotTables().Count, shapes: P.Shapes.Count, charts: P.ChartObjects().Count, fmtConds: B.Range('D2:D11').FormatConditions.Count, threaded: B.CommentsThreaded.Count, legacy: B.Comments.Count };
