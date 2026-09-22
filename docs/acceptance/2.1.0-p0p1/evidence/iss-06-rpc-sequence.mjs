// ISS-06-C：桩测试 20 条 RPC 序列快照比对
// 用法: node .scratch/evidence/rpc-sequence.mjs <改造前文件> <改造后文件> [--all]
//
// 原理：把加载项入口放进 vm，用"记录型桩"替换 Office.js / DOM / WebSocket：
//   · Excel.run(cb) 走桩，context 是记录代理（属性访问与调用都按出现顺序记录）
//   · 每次 RPC 前清空记录，RPC 结束后取出该条的调用序列
//   · 对同样 20 条 RPC、同样参数分别跑改造前/改造后文件，逐行比对序列
import fs from 'node:fs';
import vm from 'node:vm';

const SNAPSHOT_METHODS = [
  'get_workbook_info', 'get_workspace_summary', 'get_sheet_outline', 'save_workbook', 'calculate',
  'list_sheets', 'read_range', 'get_range_styles', 'write_range', 'patch_cells',
  'set_formula', 'format_cells', 'auto_fit_columns', 'freeze_panes', 'sort_range',
  'add_chart', 'create_table', 'manage_cell_comments', 'run_script', 'capture_sheet_preview'
];

const COMMON_PARAMS = {
  workbookName: 'Stub.xlsx',
  sheetName: 'Sheet1',
  address: 'A1:C3',
  range: 'A1:C3',
  values: [[1, 2, 3], [4, 5, 6], [7, 8, 9]],
  formulas: [['=1+1', '', '']],
  chartType: 'column_clustered',
  position: { leftCell: 'E2', width: 360, height: 240 },
  columns: [{ column: 'A', width: 120 }],
  rowIndex: 2,
  columnIndex: 2,
  freezeRowIndex: 2,
  freezeColumnIndex: 2,
  cellAddress: 'A1',
  text: '查找文本',
  replaceText: '替换文本',
  script: 'return 1 + 1;',
  commentText: '桩测试批注',
  tableName: 'StubTable',
  shapeId: 1,
  chartName: 'StubChart'
};

function shorten(value) {
  if (value === null) return 'null';
  if (value === undefined) return 'undefined';
  if (typeof value === 'string') return JSON.stringify(value.length > 24 ? value.slice(0, 24) + '…' : value);
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (Array.isArray(value)) return `[${value.length}]`;
  if (typeof value === 'function') return 'fn';
  return '{obj}';
}

/** 记录型桩：属性访问 / 调用 / 赋值全部按出现顺序记进 rec，并返回可继续链式访问的桩。 */
function makeStub(path, memo, rec) {
  if (memo.has(path)) return memo.get(path);
  const target = function stub() {};
  const proxy = new Proxy(target, {
    get(_t, prop) {
      if (prop === 'then') return undefined;                 // 关键：不让 await 把桩当 thenable
      if (prop === Symbol.toPrimitive) return () => '';
      if (prop === 'toString') return () => `<${path}>`;
      if (prop === 'valueOf') return () => 0;
      if (prop === 'toJSON') return () => `<${path}>`;
      if (prop === 'constructor') return Object;
      if (prop === 'inspect' || prop === Symbol.for('nodejs.util.inspect.custom')) return () => `<${path}>`;
      if (typeof prop === 'symbol') return undefined;
      return makeStub(`${path}.${String(prop)}`, memo, rec);
    },
    apply(_t, _thisArg, args) {
      rec.push(`CALL ${path}(${args.map(shorten).join(', ')})`);
      return makeStub(`${path}()`, memo, rec);
    },
    set(_t, prop, value) {
      rec.push(`SET ${path}.${String(prop)} = ${shorten(value)}`);
      return true;
    },
    has() { return true; },
    ownKeys() { return []; },
    getOwnPropertyDescriptor() { return { configurable: true, enumerable: true, value: undefined }; }
  });
  memo.set(path, proxy);
  return proxy;
}

function makeElement(tag) {
  return {
    tagName: tag, className: '', innerHTML: '', innerText: '', textContent: '',
    children: [], lastChild: null, style: {},
    prepend() {}, appendChild() {}, removeChild() {}, remove() {},
    querySelector() { return null; }, addEventListener() {}
  };
}

function loadAddon(file) {
  const rec = [];
  const sent = [];
  let socket = null;
  let readyCallback = null;

  class Socket {
    static OPEN = 1;
    static CONNECTING = 0;
    constructor(url) { this.url = url; this.readyState = Socket.OPEN; socket = this; }
    send(text) { sent.push(JSON.parse(text)); }
    close() { this.readyState = 3; }
  }

  const memo = new Map();
  const Excel = {
    run: async (callback) => {
      rec.push('Excel.run');
      const context = makeStub('context', memo, rec);
      try {
        return await callback(context);
      } finally {
        rec.push('Excel.run:finally');
      }
    }
  };
  const Office = {
    HostType: { Excel: 'Excel' },
    onReady: (cb) => { readyCallback = cb; },
    context: { document: { url: '' } },
    actions: { associate() {} }
  };

  const context = vm.createContext({
    window: { WPS_BRIDGE_CONFIG: {}, addEventListener() {} },
    document: { getElementById: () => makeElement('div'), createElement: makeElement, querySelector: () => null, addEventListener() {} },
    Office, Excel, WebSocket: Socket,
    console: { log() {}, warn() {}, error() {} },
    setTimeout: () => 0, clearTimeout() {}, setInterval: () => 0, clearInterval() {},
    alert() {}, confirm: () => false,
    Math: Object.create(Math, { random: { value: () => 0.5 } }),
    Date
  });
  vm.runInContext(fs.readFileSync(file, 'utf8'), context);
  readyCallback({ host: 'Excel' });
  socket.onopen();

  return {
    rec,
    sent,
    readyCallback,
    async drain() { for (let i = 0; i < 25; i++) await new Promise((r) => setImmediate(r)); },
    async call(method, params) {
      rec.length = 0;
      const before = sent.length;
      socket.onmessage({ data: JSON.stringify({ id: `${method}-1`, method, params }) });
      for (let i = 0; i < 25; i++) await new Promise((r) => setImmediate(r));
      const response = sent.slice(before).findLast((p) => p && p.type === 'rpc_response');
      return { lines: rec.slice(), response: response || null };
    }
  };
}

async function snapshotOf(host, methods) {
  const out = [];
  for (const method of methods) {
    const { lines, response } = await host.call(method, { ...COMMON_PARAMS, method });
    out.push({ method, lines, response });
  }
  return out;
}

function responseSignature(response) {
  if (!response) return '(无 rpc_response)';
  if (response.error) return `error: ${response.error}`;
  let body;
  try { body = JSON.stringify(response.result); } catch (e) { body = `<序列化失败: ${e.message}>`; }
  return `result: ${body}`;
}

const [beforeFile, afterFile] = process.argv.slice(2).filter((a) => !a.startsWith('--'));
const runAll = process.argv.includes('--all');

const beforeHost = loadAddon(beforeFile);
const afterHost = loadAddon(afterFile);
await beforeHost.drain();
await afterHost.drain();
// 自检：两个文件入口都应完成 Office.onReady + WebSocket onopen 的注册动作
await beforeHost.drain();
await afterHost.drain();

// --all：把改造后文件 dispatchExcelTool 里的全部 case 都当作快照条目（扩展覆盖，不属于台账声称的 20 条）
function allCases(file) {
  const src = fs.readFileSync(file, 'utf8');
  const body = src.slice(src.indexOf('async function dispatchExcelTool'));
  const names = [...body.matchAll(/case "([^"]+)"/g)].map((m) => m[1]);
  return [...new Set(names)];
}
const allMethods = runAll ? allCases(afterFile) : SNAPSHOT_METHODS;

const A = await snapshotOf(beforeHost, allMethods);
const B = await snapshotOf(afterHost, allMethods);

console.log('=== ISS-06-C · 桩测试 RPC 序列快照比对 ===');
console.log(`改造前: ${beforeFile}`);
console.log(`改造后: ${afterFile}`);
console.log(`快照条数: ${A.length}${runAll ? '（--all：含扩展集合）' : ''}`);
console.log('');

let identical = 0;
const mismatches = [];
for (let i = 0; i < A.length; i++) {
  const a = A[i].lines.join('\n');
  const b = B[i].lines.join('\n');
  const ra = responseSignature(A[i].response);
  const rb = responseSignature(B[i].response);
  const same = a === b && ra === rb;
  if (same) identical++;
  else mismatches.push({ method: A[i].method, a, b, ra, rb });
  console.log(`[${same ? '一致' : '不一致'}] ${String(i + 1).padStart(2, '0')}. ${A[i].method}`);
  console.log(`        调用序列行数 ${A[i].lines.length} / ${B[i].lines.length} · 响应 ${ra.slice(0, 110)}`);
}

console.log('');
console.log(`逐行一致的 RPC 条数: ${identical}/${A.length}`);
console.log(`不一致条数: ${mismatches.length}`);
for (const m of mismatches) {
  console.log('');
  console.log(`--- 不一致: ${m.method} ---`);
  console.log(`改造前响应: ${m.ra}`);
  console.log(`改造后响应: ${m.rb}`);
  console.log('改造前序列:');
  console.log(m.a);
  console.log('改造后序列:');
  console.log(m.b);
}

console.log('');
console.log('--- 明细：每条 RPC 的调用序列（改造前） ---');
for (const item of A) {
  console.log('');
  console.log(`### ${item.method}  ·  ${item.lines.length} 行`);
  for (const l of item.lines) console.log(`    ${l}`);
  console.log(`    响应: ${responseSignature(item.response)}`);
}
