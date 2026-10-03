import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { DOMParser } from '@xmldom/xmldom';
// @ts-ignore 构建脚本导出的只读 PNG 解码器。
import { decodePng, resizeRgba, encodePng } from '../scripts/build-win-icon.mjs';

const tick = () => new Promise(resolve => setImmediate(resolve));
const wpsCode = fs.readFileSync('wps-addon/addon-core.js', 'utf8');
const officeCode = fs.readFileSync('office-addon/public/taskpane.js', 'utf8');
function node() {
  const handlers: Record<string, any> = {};
  return { innerText: '', className: '', hidden: false, disabled: false, handlers,
    addEventListener: (name: string, handler: any) => { handlers[name] = handler; } };
}
function wpsHost() {
  const storage = new Map<string, string>(), panes = new Map<number, any>();
  let created = 0;
  const sheet = { Name: 'Sheet1', Range() {}, Visible: -1 };
  const book = { Name: 'Exact.xlsx', FullName: '/tmp/Exact.xlsx', ActiveSheet: sheet, Worksheets: { Count: 1, Item: () => sheet } };
  const app = { ActiveWorkbook: book, Workbooks: { Count: 1, Item: () => book } };
  const api = { EtApplication: () => app, PluginStorage: { getItem: (key: string) => storage.get(key), setItem: (key: string, value: string) => storage.set(key, value) },
    GetTaskPane: (id: number) => panes.get(id),
    CreateTaskPane: (url: string, title: string) => {
      const pane = { ID: ++created, Visible: false, url, title, navigationCount: 0,
        Navigate(next: string) { this.url = next; this.navigationCount++; } };
      panes.set(pane.ID, pane); return pane;
    } };
  return { api, storage, panes, get created() { return created; } };
}
function loadWps(host: ReturnType<typeof wpsHost>, panel = false, url = 'file:///addon/index.html') {
  const sockets: any[] = [], packets: any[] = [], intervals: any[] = [], requests: any[] = [];
  class Socket {
    static OPEN = 1; static CONNECTING = 0; readyState = 1; onopen: any; onmessage: any;
    constructor() { sockets.push(this); }
    send(value: string) { packets.push(JSON.parse(value)); }
  }
  const nodes: Record<string, ReturnType<typeof node>> = {};
  for (const id of ['btnClose', 'btnReconnect', 'statusBadge', 'detailInfo', 'docText', 'sheetText', 'selectionText', 'hostText', 'sheetRow', 'selectionRow']) nodes[id] = node();
  const listeners: Record<string, any> = {};
  const window: any = { WPS_BRIDGE_CONFIG: { port: 12345, token: 'isolated' }, CHARFLOAT_PANEL_VIEW: panel,
    location: { href: url }, addEventListener: (name: string, handler: any) => { listeners[name] = handler; } };
  vm.runInNewContext(wpsCode, { window, document: { readyState: 'complete', getElementById: (id: string) => nodes[id] || null },
    wps: host.api, WebSocket: Socket, URL, console: { log() {} }, alert: (message: string) => { throw new Error(message); },
    fetch: async (target: string, options: any) => { requests.push({ target, options }); return { ok: true, json: async () => ({ components: { excel: { connected: true, summary: { workbookName: 'Exact.xlsx', activeSheetName: 'Sheet1', selection: { address: '$A$1' } } }, word: { connected: true, summary: { documentName: 'Document.docx' } }, ppt: { connected: true, summary: { presentationName: 'Deck.pptx' } } } }) }; },
    setTimeout: () => 1, clearTimeout() {}, setInterval: (fn: any) => { intervals.push(fn); } });
  return { window, sockets, packets, nodes, intervals, requests, listeners };
}

test('WPS 侧栏使用单实例，隐藏后后台仍接收 RPC；面板不创建第二条 WebSocket', async () => {
  const host = wpsHost(), core = loadWps(host);
  assert.equal(core.sockets.length, 1);
  core.sockets[0].onopen();
  core.window.OnActionShowTaskPane(); core.window.OnActionShowTaskPane();
  assert.equal(host.created, 1);
  const pane = host.panes.get(1);
  assert.equal(pane.title, '字浮 CharFloat');
  assert.equal(new URL(pane.url).searchParams.get('paneId'), '1');
  assert.match(new URL(pane.url).searchParams.get('build')!, /^[a-f0-9]{64}$/);
  const panel = loadWps(host, true, pane.url);
  await tick();
  assert.equal(panel.sockets.length, 0);
  assert.equal(panel.intervals.length, 0);
  assert.equal(panel.nodes.docText.innerText, 'Exact.xlsx');
  assert.equal(panel.requests[0].options.headers.Authorization, 'Bearer isolated');
  panel.nodes.btnClose.handlers.click();
  assert.equal(pane.Visible, false);
  core.sockets[0].onmessage({ data: JSON.stringify({ id: 'after-hide', method: 'ping', params: {} }) });
  await tick();
  assert.equal(core.packets.find(p => p.id === 'after-hide').result.pong, true);
  core.window.OnActionShowTaskPane();
  assert.equal(pane.Visible, true);
  assert.equal(host.created, 1);
  core.intervals[0]();
  assert.equal(core.packets.at(-1).version, '2.2.0');
});

test('WPS 入口重载后复用旧侧栏，新构建导航已有侧栏而不是重复创建', () => {
  const host = wpsHost();
  loadWps(host).window.OnActionShowTaskPane();
  const pane = host.panes.get(1), previousNavigations = pane.navigationCount;
  host.storage.set('charfloat.taskpane.excel.build', 'prior-build');
  loadWps(host).window.OnActionShowTaskPane();
  assert.equal(host.created, 1);
  assert.equal(pane.navigationCount, previousNavigations + 1);
});

test('WPS 无共享存储接口时，当前入口仍可创建并复用侧栏', () => {
  const host = wpsHost();
  Object.defineProperty(host.api, 'PluginStorage', { get() { throw new Error('unavailable'); } });
  const core = loadWps(host);
  core.window.OnActionShowTaskPane();
  core.window.OnActionShowTaskPane();
  assert.equal(host.created, 1);
  assert.equal(host.panes.get(1).Visible, true);
});

test('WPS 文字/演示侧栏显示对应文档，隐藏工作表和选区信息', async () => {
  for (const [component, name] of [['word', 'Document.docx'], ['ppt', 'Deck.pptx']]) {
    const panel = loadWps(wpsHost(), true, `file:///addon/panel.html?component=${component}&paneId=7`);
    await tick();
    assert.equal(panel.nodes.docText.innerText, name);
    assert.equal(panel.nodes.sheetRow.hidden, true);
    assert.equal(panel.nodes.selectionRow.hidden, true);
    assert.equal(panel.sockets.length, 0);
  }
});

function loadOffice(settings: { behavior: string }, options: { supported?: boolean; failStartup?: boolean; workbookName?: string; documentUrl?: string } = {}) {
  const sockets: any[] = [], packets: any[] = [], nodes: Record<string, ReturnType<typeof node>> = {};
  for (const id of ['backgroundStatus', 'backgroundHint', 'btnHide']) nodes[id] = node();
  let ready: any, visibility: any, startupWrites = 0, hooks = 0, hideCalls = 0;
  const range: any = { address: 'Sheet1!A1', rowCount: 1, columnCount: 1, values: [[0]], formulas: [['=0']], text: [['0']], load() {} };
  range.getCell = () => range;
  const sheet = { name: 'Sheet1', load() {}, getRange: () => range, onSelectionChanged: { add() { hooks++; } } };
  const context = { workbook: { name: options.workbookName || 'Exact.xlsx', load() {}, worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }, getSelectedRange: () => range }, sync: async () => {} };
  class Socket {
    static OPEN = 1; static CONNECTING = 0; readyState = 1; onopen: any; onmessage: any;
    constructor() { sockets.push(this); }
    send(value: string) { packets.push(JSON.parse(value)); }
  }
  const office = { HostType: { Excel: 'Excel' }, StartupBehavior: { load: 'load' }, onReady: (fn: any) => { ready = fn; },
    context: { document: { url: options.documentUrl ?? 'file:///Exact.xlsx' }, requirements: { isSetSupported: () => options.supported !== false } },
    addin: { getStartupBehavior: async () => settings.behavior,
      setStartupBehavior: async (behavior: string) => { if (options.failStartup) throw new Error('startup rejected'); startupWrites++; settings.behavior = behavior; },
      onVisibilityModeChanged: async (fn: any) => { visibility = fn; return async () => {}; },
      hide: async () => { hideCalls++; await visibility?.({ visibilityMode: 'hidden' }); } } };
  vm.runInNewContext(officeCode, { window: { location: { hash: '', href: 'https://localhost:19891/office-addon/taskpane.html' } },
    document: { readyState: 'complete', getElementById: (id: string) => nodes[id] || null }, Office: office,
    Excel: { run: async (fn: any) => fn(context) }, WebSocket: Socket, URL, console: { log() {}, error() {} }, setTimeout() {}, clearTimeout() {} });
  return { ready: () => ready({ host: 'Excel' }), sockets, packets, nodes,
    get startupWrites() { return startupWrites; }, get hooks() { return hooks; }, get hideCalls() { return hideCalls; } };
}

test('Office 首次启用自动启动，收起侧栏后保持同一连接并可处理读请求', async () => {
  const settings = { behavior: 'none' }, host = loadOffice(settings);
  host.ready(); await tick(); await host.sockets[0].onopen();
  assert.equal(settings.behavior, 'load');
  assert.equal(host.startupWrites, 1);
  assert.equal(host.nodes.btnHide.disabled, false);
  await host.nodes.btnHide.handlers.click();
  assert.equal(host.hideCalls, 1);
  assert.equal(host.sockets.length, 1);
  host.sockets[0].onmessage({ data: JSON.stringify({ id: 'hidden-read', method: 'read_range', params: { address: 'A1' } }) });
  await tick();
  assert.equal(host.packets.find(p => p.id === 'hidden-read').result.values[0][0], 0);
  assert.equal(host.packets.find(p => p.type === 'register' && p.summary.backgroundStartup === 'enabled').version, '2.2.0');
  host.ready();
  assert.equal(host.sockets.length, 1);
  assert.equal(host.hooks, 1);
  const reopened = loadOffice(settings);
  reopened.ready(); await tick();
  assert.equal(reopened.startupWrites, 0, '已有自动启动配置不重复写入');
  assert.equal(reopened.sockets.length, 1);
});

test('Office 不支持共享运行时或启用失败时，不阻断现有连接并如实提示', async () => {
  for (const options of [{ supported: false }, { failStartup: true }]) {
    const host = loadOffice({ behavior: 'none' }, options);
    host.ready(); await tick(); await host.sockets[0].onopen();
    assert.equal(host.sockets.length, 1);
    assert.equal(host.startupWrites, 0);
    assert.match(host.nodes.backgroundStatus.innerText, options.supported === false ? /保持侧栏开启/ : /尚未启用/);
    if (options.supported === false) assert.equal(host.nodes.btnHide.disabled, true);
  }
});

test('Office 未保存工作簿使用宿主实际名称，避免多个后台文档都报告相同占位名', async () => {
  for (const name of ['Book1', 'Book2']) {
    const host = loadOffice({ behavior: 'none' }, { workbookName: name, documentUrl: '' });
    host.ready(); await tick(); await host.sockets[0].onopen();
    assert.equal(host.packets.find(p => p.type === 'register').summary.workbookName, name);
  }
});

test('清单将 Office 命令与侧栏放进长生命周期运行时，WPS 只保留侧栏入口', () => {
  const office = new DOMParser().parseFromString(fs.readFileSync('office-addon/excel/manifest.xml', 'utf8'), 'application/xml');
  const runtime = office.getElementsByTagName('Runtime')[0];
  assert.equal(runtime.parentNode?.nodeName, 'Runtimes');
  assert.equal(runtime.parentNode?.parentNode?.nodeName, 'Host');
  assert.equal(runtime.getAttribute('lifetime'), 'long');
  assert.equal(runtime.getAttribute('resid'), office.getElementsByTagName('FunctionFile')[0].getAttribute('resid'));
  assert.equal(office.getElementsByTagName('Action')[0].getElementsByTagName('SourceLocation')[0].getAttribute('resid'), runtime.getAttribute('resid'));
  assert.equal(office.getElementsByTagName('TaskpaneId').length, 0);
  const ribbon = new DOMParser().parseFromString(fs.readFileSync('wps-addon/ribbon.xml', 'utf8'), 'application/xml');
  assert.equal(ribbon.getElementsByTagName('tab')[0].getAttribute('idMso'), 'TabHome');
  assert.equal(ribbon.getElementsByTagName('button').length, 1);
  assert.equal(ribbon.getElementsByTagName('button')[0].getAttribute('onAction'), 'OnActionShowTaskPane');
  assert.equal(ribbon.getElementsByTagName('button')[0].getAttribute('size'), 'large');
  assert.equal(ribbon.getElementsByTagName('button')[0].getAttribute('showImage'), 'true');
  assert.equal(ribbon.getElementsByTagName('button')[0].getAttribute('showLabel'), 'true');
});

test('WPS 侧栏只显示状态与日志，图标回调使用有可见像素的当前 App 图标', () => {
  const html = new DOMParser().parseFromString(fs.readFileSync('wps-addon/panel.html', 'utf8'), 'text/html');
  const buttons = Array.from(html.getElementsByTagName('button')).map(button => button.getAttribute('id')).sort();
  assert.deepEqual(buttons, ['btnClose', 'btnReconnect']);
  const text = html.documentElement?.textContent || '';
  for (const removed of ['工作表操作', '常用 AI 指令', '控制中心', '使用向导']) assert.equal(text.includes(removed), false);
  assert.ok(text.includes('当前工作区')); assert.ok(text.includes('运行日志'));
  const logo = fs.readFileSync('wps-addon/logo.png');
  assert.ok(logo.equals(fs.readFileSync('resources/icon.png')));
  const { rgba } = decodePng(logo);
  let visible = false;
  for (let i = 3; i < rgba.length; i += 4) if (rgba[i] > 0) { visible = true; break; }
  assert.equal(visible, true, '全透明 PNG 不能作为 App 图标');
  const core = loadWps(wpsHost());
  for (const control of ['btnShowCharFloatPanel', { Id: 'btnShowCharFloatPanel' }, { ID: 'btnShowCharFloatPanel' }]) {
    assert.equal(core.window.OnGetImage(control), 'ribbon-icon.png');
  }
  const original = decodePng(logo);
  for (const [name, size] of [['ribbon-icon.png', 32], ['ribbon-icon@2x.png', 64]] as const) {
    const bytes = fs.readFileSync('wps-addon/' + name);
    const icon = decodePng(bytes);
    assert.equal(icon.width, size); assert.equal(icon.height, size);
    assert.ok(bytes.equals(encodePng(resizeRgba(original, size))), '专用图标应从原图直接高质量缩放');
  }
});

test('两种侧栏共享 Office 样式并显示 App 图标，Office 清单图标均指向新 URL', () => {
  const css=fs.readFileSync('office-addon/public/taskpane.css','utf8');
  assert.equal(fs.readFileSync('wps-addon/panel.css','utf8').split('\n').slice(1).join('\n'),css);
  for (const name of ['wps-addon/panel.html','office-addon/public/taskpane.html']) {
    const html=new DOMParser().parseFromString(fs.readFileSync(name,'utf8'),'text/html');
    for (const cls of ['app-container','brand-logo','status-badge','doc-card','doc-item-row','footer']) {
      assert.ok(Array.from(html.getElementsByTagName('*')).some(el=>(el.getAttribute('class')||'').split(' ').includes(cls)),name+': '+cls);
    }
    const brand=Array.from(html.getElementsByTagName('*')).find(el=>el.getAttribute('class')==='brand-logo')!;
    assert.equal(brand.getElementsByTagName('img').length,1);
    assert.equal(brand.getElementsByTagName('svg').length,0);
  }
  const manifest=new DOMParser().parseFromString(fs.readFileSync('office-addon/excel/manifest.xml','utf8'),'application/xml');
  for (const icon of Array.from(manifest.getElementsByTagName('Icon'))) {
    const sizes=Array.from(icon.getElementsByTagName('bt:Image')).map(el=>Number(el.getAttribute('size')));
    assert.deepEqual(sizes,[16,20,24,32,48,64,80]);
  }
  for (const image of Array.from(manifest.getElementsByTagName('bt:Image')).filter(el=>el.hasAttribute('DefaultValue'))) {
    assert.match(image.getAttribute('DefaultValue')!,/^https:\/\/localhost:19891\/assets\/icon-\d+\.png\?v=charfloat-20261003$/);
  }
  assert.match(manifest.getElementsByTagName('IconUrl')[0].getAttribute('DefaultValue')!,/\?v=charfloat-20261003$/);
});
