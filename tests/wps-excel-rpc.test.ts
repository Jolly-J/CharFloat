import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

// 可对固定基线与新生成物运行同一套夹具；只操作内存模拟宿主。
const code = fs.readFileSync(process.env.OAB_TEST_WPS_ARTIFACT || 'wps-addon/addon-core.js', 'utf8');
const snapshots: Record<string, unknown> = {};
const clone = (value: any) => JSON.parse(JSON.stringify(value));
function load(failWrite = false) {
  const trace: any[] = [], packets: any[] = [], intervals: any[] = [];
  let socket: any, listener: any;
  const track = (label: string, object: any) => new Proxy(object, {
    set(target, key, value) { trace.push(['set', label, String(key), clone(value)]); target[key] = value; return true; }
  });
  const values: any[] = [0, '旧B'], formulas: any[] = [0, '旧B'];
  const font = track('font', { Name: '微软雅黑', Size: 11, Bold: false, Italic: false, Color: 0 });
  const interior = track('interior', { Color: 16777215, Pattern: 1 });
  const validation: any = track('validation', { Type: -4142, Operator: 1, Formula1: '', Formula2: '', IgnoreBlank: true, InCellDropdown: true,
    Delete() { trace.push(['validation.delete']); this.Type = -4142; },
    Add(type: number, alert: number, op: number, f1: any, f2: any) { trace.push(['validation.add', type, alert, op, f1, f2 ?? null]); this.Type = type; this.Operator = op; this.Formula1 = f1; this.Formula2 = f2 ?? ''; }
  });
  const border = { LineStyle: 1, Weight: 1, Color: 0 };
  const cells: any[] = values.map((_v, i) => {
    const cell: any = { Row: 1, Column: i + 1, Address: () => i ? '$B$1' : '$A$1', Font: font, Interior: interior, NumberFormat: '0', Validation: validation, MergeCells: false, Comment: null,
      AddComment(text: string) { trace.push(['comment.add', i, text]); let content = text; this.Comment = track('comment', { Visible: true, Author: '宿主', Text: (next?: string) => { if (next !== undefined) content = next; return content; } }); return this.Comment; }
    };
    Object.defineProperties(cell, { Value2: { get: () => values[i] }, Formula: { get: () => formulas[i], set: value => { trace.push(['cell.formula', i, value]); values[i] = value === '' || value === null || value === undefined ? null : value; formulas[i] = value === null || value === undefined ? '' : value; } } });
    return cell;
  });
  const rowItems = new Map<number, any>();
  const rows = { Count: 1, Item(index: number) { if (!rowItems.has(index)) rowItems.set(index, track('row.' + index, { RowHeight: 15, Hidden: false })); return rowItems.get(index); } };
  const makeRange = (address: string) => {
    const columns = address.includes(':') ? 2 : 1;
    const cellCollection: any = (_row: number, col: number) => cells[col - 1];
    cellCollection.Item = cellCollection;
    const range: any = track('range', { Row: 1, Column: 1, Rows: rows, Columns: { Count: columns }, Cells: cellCollection,
      Address: () => columns === 2 ? '$A$1:$B$1' : '$A$1', Font: font, Interior: interior, Borders: { Item: () => border }, Validation: validation,
      NumberFormat: '0', HorizontalAlignment: -4131, VerticalAlignment: -4108, WrapText: false, RowHeight: 15, ColumnWidth: 8, MergeCells: false,
      AutoFilter() { trace.push(['filter.apply']); sheet.AutoFilterMode = true; sheet.AutoFilter = { Range: range, Filters: { Item: () => ({ On: false }) } }; }
    });
    Object.defineProperties(range, {
      Value2: { get: () => columns === 2 ? [values.slice()] : values[0], set: matrix => { trace.push(['values.write', clone(matrix)]); if (failWrite) throw new Error('模拟宿主写入失败'); matrix[0].forEach((value: any, i: number) => { values[i] = value; formulas[i] = value; }); } },
      Formula: { get: () => columns === 2 ? [formulas.slice()] : formulas[0], set: matrix => { trace.push(['formulas.write', clone(matrix)]); matrix[0].forEach((value: any, i: number) => { values[i] = value === '' || value === null || value === undefined ? null : value; formulas[i] = value === null || value === undefined ? '' : value; }); } }
    });
    return range;
  };
  const empty = { Count: 0, Item() { throw new Error('集合为空'); } };
  const chart: any = track('chart', { ChartType: 51, HasTitle: true, HasLegend: true, ChartTitle: track('chart.title', { Text: '原标题' }), Legend: { Position: -4107 }, SeriesCollection: () => empty });
  const chartShape: any = { Name: '图表1', HasChart: true, Chart: chart, Type: 3, Left: 0, Top: 0, Width: 200, Height: 100 };
  const shapes: any[] = [chartShape];
  const sheet: any = { Name: 'Test', Visible: -1, AutoFilterMode: false, ProtectContents: false, Rows: rows, ListObjects: empty,
    Shapes: { get Count() { return shapes.length; }, Item: (index: number) => shapes[index - 1] },
    Range(address: string) { trace.push(['range', address]); return makeRange(address); },
    Activate() { trace.push(['sheet.activate']); }, PageSetup: track('page', { Orientation: 1, PaperSize: 9, PrintArea: '', Zoom: 100 }), HPageBreaks: empty, VPageBreaks: empty, PivotTables: () => empty
  };
  sheet.Cells = { Item: (_r: number, col: number) => cells[col - 1] };
  sheet.UsedRange = makeRange('A1:B1');
  const book: any = { Name: 'Exact.xlsx', FullName: '/tmp/Exact.xlsx', ActiveSheet: sheet, CustomViews: empty, BuiltinDocumentProperties: empty, CustomDocumentProperties: empty,
    Worksheets: { Count: 1, Item(index: any) { if (index === 1 || index === 'Test') return sheet; throw new Error('工作表不存在'); } },
    Save() { trace.push(['workbook.save']); }
  }; sheet.Parent = book;
  const app: any = { ActiveWorkbook: book, ActiveSheet: sheet, Selection: sheet.UsedRange, ActiveWindow: track('window', { DisplayGridlines: true, DisplayHeadings: true }),
    Workbooks: { Count: 1, Item(index: any) { if (index === 1 || index === 'Exact.xlsx') return book; throw new Error('工作簿不存在'); } }
  };
  class Socket { static OPEN = 1; static CONNECTING = 0; readyState = 1; onmessage: any; onopen: any; constructor() { socket = this; } send(text: string) { packets.push(JSON.parse(text)); } }
  class FixedDate extends Date { constructor(...args: any[]) { super(args.length ? args[0] : 1700000000000); } static now() { return 1700000000000; } }
  vm.runInNewContext(code, { window: { WPS_BRIDGE_CONFIG: { port: 12345, token: 'isolated' }, addEventListener() {} }, document: { readyState: 'complete', getElementById: () => null },
    wps: { EtApplication: () => app, ApiEvent: { AddApiEventListener(name: string, fn: any) { trace.push(['event.register', name]); listener = fn; } } },
    WebSocket: Socket, Date: FixedDate, console: { log() {} }, setTimeout() {}, clearTimeout() {}, setInterval(fn: any) { intervals.push(fn); }, alert() {}, confirm: () => false });
  socket.onopen();
  return { trace, values, formulas, font, validation, chart, sheet, rows, packets,
    select() { assert.ok(listener, '事件监听必须在连接后注册'); listener(sheet, sheet.UsedRange); listener(sheet, sheet.UsedRange); },
    heartbeat() { intervals[0](); },
    async call(method: string, params: any) {
      socket.onmessage({ data: JSON.stringify({ id: 'case', method, params: { workbookName: 'Exact.xlsx', sheetName: 'Test', ...params } }) });
      await new Promise(resolve => setImmediate(resolve));
      return packets.findLast(packet => packet.type === 'rpc_response');
    }
  };
}

const cases: [string, string, any, (host: ReturnType<typeof load>, result: any) => void][] = [
  ['读区域保留零与公式', 'read_range', { address: 'A1:B1', includeFormulas: true }, (_h, r) => assert.deepEqual(r.result.values, [[0, '旧B']])],
  ['混合值和公式不清空空公式位置', 'patch_cells', { address: 'A1:B1', values: [['新A', '新B']], formulas: [['=1+1', '']] }, (h, r) => { assert.ok(r.result); assert.deepEqual(h.values, ['=1+1', '新B']); }],
  ['非法矩阵执行前拒绝', 'patch_cells', { address: 'A1:B1', values: [[1]] }, (h, r) => { assert.match(r.error, /尺寸/); assert.ok(!h.trace.some(x => x[0] === 'values.write')); }],
  ['样式跨模块写入与读回', 'format_cells', { address: 'A1:B1', fontSize: 18, bold: true, fontColor: '#FF0000', verticalAlignment: 'top', rowHeight: 24 }, (h, r) => { assert.ok(r.result, r.error); assert.equal(h.font.Size, 18); assert.equal(h.font.Color, 255); assert.equal(h.rows.Item(1).RowHeight, 24); }],
  ['设计语言读取不编造缺失字段', 'get_style_token', {}, (_h, r) => { assert.ok(r.result, r.error); assert.ok(JSON.stringify(r.result).includes('微软雅黑')); }],
  ['验证读取跨模块运行', 'get_range_styles', { address: 'A1:B1', include: ['fontSize', 'validation', 'validationViolations'] }, (_h, r) => { assert.ok(r.result.validationCheck, r.error); assert.equal(r.result.styles.validation.type, -4142); }],
  ['验证写后核对列表', 'set_data_validation', { address: 'A1:B1', validationType: 'list', listItems: ['通过', '不通过'] }, (h, r) => { assert.ok(r.result, r.error); assert.equal(h.validation.Type, 3); assert.ok(h.trace.some(x => x[0] === 'validation.add')); }],
  ['工作表视图生效', 'set_sheet_view', { showGridlines: false }, (h, r) => { assert.ok(r.result, r.error); assert.ok(h.trace.some(x => x[1] === 'window' && x[2] === 'DisplayGridlines')); }],
  ['行隐藏逐行核对', 'modify_rows_columns', { targetType: 'row', action: 'hide', index: 1, count: 2 }, (h, r) => { assert.ok(r.result, r.error); assert.equal(h.rows.Item(1).Hidden, true); assert.equal(h.rows.Item(2).Hidden, true); }],
  ['图表清点', 'get_charts', {}, (_h, r) => { assert.ok(r.result, r.error); assert.ok(JSON.stringify(r.result).includes('图表1')); }],
  ['图表更新写后核对', 'update_chart', { chartName: '图表1', title: '新标题', chartType: 'line' }, (h, r) => { assert.ok(r.result, r.error); assert.equal(h.chart.ChartTitle.Text, '新标题'); assert.equal(h.chart.ChartType, 4); }],
  ['形状清点', 'list_shapes', {}, (_h, r) => { assert.ok(r.result, r.error); assert.ok(JSON.stringify(r.result).includes('图表1')); }],
  ['结构化表清点', 'manage_table', { action: 'list' }, (_h, r) => { assert.ok(r.result, r.error); assert.deepEqual(r.result.tables, []); }],
  ['透视表清点', 'create_pivot_table', { action: 'read' }, (_h, r) => assert.ok(r.result, r.error)],
  ['筛选越界在写操作前拒绝', 'set_filter_and_sort', { range: 'A1:B1', enableAutoFilter: true, sortRules: [{ colIndex: 3, order: 'asc' }] }, (h, r) => { assert.match(r.error, /越界/); assert.ok(!h.trace.some(x => x[0] === 'filter.apply')); }],
  ['批注保留作者与目标地址', 'manage_cell_comments', { address: 'A1', action: 'add', text: '批注正文', author: '审阅者' }, (_h, r) => { assert.ok(r.result, r.error); assert.equal(r.result.address, '$A$1'); assert.equal(r.result.authorOnHost, '审阅者'); }],
  ['保存定位目标工作簿', 'save_workbook', {}, (h, r) => { assert.ok(r.result, r.error); assert.ok(h.trace.some(x => x[0] === 'workbook.save')); }],
  ['打印设置写后读取', 'configure_print_layout', { orientation: 'landscape', printArea: 'A1:B1' }, (h, r) => { assert.ok(r.result, r.error); assert.equal(h.sheet.PageSetup.Orientation, 2); assert.equal(h.sheet.PageSetup.PrintArea, 'A1:B1'); }],
  ['patch_cells 规范化字体与字重', 'patch_cells', { address: 'A1:B1', values: [['测试A', '测试B']], fontName: 'Microsoft YaHei', bold: false }, (h, r) => {
    assert.ok(r.result, r.error);
    assert.equal(r.result.appliedFontName, 'Microsoft YaHei');
    assert.equal(r.result.appliedBold, false);
    assert.equal(h.font.Name, 'Microsoft YaHei');
    assert.equal(h.font.Bold, false);
  }],
  ['get_range_styles 读取富文本 runs', 'get_range_styles', { address: 'A1', include: ['fontName', 'bold', 'runs'] }, (_h, r) => {
    assert.ok(r.result, r.error);
    assert.ok(Array.isArray(r.result.styles.runs));
    assert.ok(r.result.styles.runs.length >= 1);
    assert.equal(r.result.styles.runs[0].fontName, '微软雅黑');
  }],
  ['未知方法保持明确失败', 'no_such_method', {}, (_h, r) => assert.ok(r.error)],
];
for (const [label, method, params, verify] of cases) test(label, async () => {
  const host = load(); const response = await host.call(method, params); verify(host, response);
  snapshots[label] = clone({ response, trace: host.trace, values: host.values, formulas: host.formulas });
});
test('宿主写入异常不能伪装成功', async () => {
  const host = load(true); const response = await host.call('patch_cells', { address: 'A1:B1', values: [['新A', '新B']] });
  assert.match(response.error, /模拟宿主写入失败/); assert.deepEqual(host.values, [0, '旧B']);
  snapshots['宿主写入异常'] = clone({ response, trace: host.trace, values: host.values });
});
test('选区事件去重，注册与心跳指纹保持一致', () => {
  const host = load(); host.select(); host.heartbeat();
  const events = host.packets.filter(p => p.type === 'event');
  assert.equal(events.length, 1); assert.equal(events[0].data.address, '$A$1:$B$1');
  const fingerprint = code.slice(0, 4096).match(/ADDON_BUILD_FINGERPRINT: ([a-f0-9]{64})/)![1];
  const registrations = host.packets.filter(p => p.type === 'register');
  assert.equal(registrations.length, 2); registrations.forEach(p => assert.equal(p.buildFingerprint, fingerprint));
  snapshots['事件与心跳'] = clone({ events, trace: host.trace, registrationCount: registrations.length });
});
after(() => {
  if (process.env.CHARFLOAT_EXCEL_SNAPSHOT) fs.writeFileSync(process.env.CHARFLOAT_EXCEL_SNAPSHOT, JSON.stringify(snapshots, null, 2) + '\n');
});
