// CAP-08 离线验证：用 Office.js 批处理语义的 mock 跑真实 taskpane.js 里的形状处理函数。
//
// 目的：在动真机之前，把「读回口径 / 参数解析 / 错误处理」这些纯逻辑错误挡掉。
// 这不是真机验证（真机读数另见文档），也不替代它。
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
let src = fs.readFileSync(path.join(REPO, 'office-addon', 'public', 'taskpane.js'), 'utf8');
const EXPOSE = `
  globalThis.__CAP08 = { handleShapeSelfTest, handleAddShape, handleListShapes, handleUpdateShape, handleGroupShapes, handleUngroupShapes, handleSetShapeZOrder, handleGetActiveShape, handleExportShapeImage, normalizeHexColor, resolveGeometricShapeType, shapeToJson };
`;
src = src.replace(/\n\}\)\(\);\s*$/, EXPOSE + '\n})();');

// ── 极简 Office.js 批处理 mock ─────────────────────────────────────────────
let idSeq = 0;
function makeStore(init) {
  const store = Object.assign({}, init);
  const sid = 'shape-' + (++idSeq);
  Object.defineProperty(store, '__id', { value: sid, enumerable: false });
  // Office.js 的 Shape.id 是只读字符串，getItem(id) 也用它；mock 保持一致
  if (store.id === undefined) Object.defineProperty(store, 'id', { value: sid, enumerable: true, writable: true });
  return store;
}

function makeProxy(store, onWrite) {
  const pendingWrites = {};
  return new Proxy({}, {
    get(_t, prop) {
      if (prop === '__store') return store;
      if (prop === '__pendingWrites') return pendingWrites;
      if (prop === 'load') return () => (store && Array.isArray(store.items) ? { items: store.items } : undefined);
      if (prop === 'type' || prop === 'name' || prop === 'id') return store[prop];
      const v = store[prop];
      if (typeof v === 'function') return v.bind(store);
      if (v && typeof v === 'object') {
        if (!store['__proxy_' + String(prop)]) {
          const child = makeProxy(v, (k, val) => { store[prop][k] = val; });
          Object.defineProperty(store, '__proxy_' + String(prop), { value: child, enumerable: false });
        }
        return store['__proxy_' + String(prop)];
      }
      return v;
    },
    set(_t, prop, value) { pendingWrites[prop] = value; onWrite?.(prop, value); return true; }
  });
}

function makeShape(init) {
  const store = makeStore(Object.assign({
    name: init.name || 'Shape' + idSeq, type: init.type || 'GeometricShape',
    left: 0, top: 0, width: 100, height: 50, rotation: 0, zOrderPosition: 0, visible: true,
    fill: { type: 'Solid', foregroundColor: '#FFFFFF', transparency: 0 },
    lineFormat: { color: '#000000', weight: 1, visible: true },
    textFrame: {
      textRange: { text: '', font: { bold: false, size: 11, color: '#000000' } },
      horizontalAlignment: 'Center', verticalAlignment: 'Middle'
    }
  }, init));
  const proxy = makeProxy(store, () => {});
  store.__proxy = proxy;
  // ShapeFill.setSolidColor 是 Office.js 的真实方法，mock 也要有
  store.fill.setSolidColor = (hex) => { store.fill.foregroundColor = hex; store.fill.type = 'Solid'; };
  store.fill.clear = () => { store.fill.foregroundColor = null; store.fill.type = 'Background'; };
  // 增量/缩放/层级/删除/导出：mock 只记录调用，具体效果由断言检查 store
  store._ops = [];
  store.incrementRotation = (d) => { store.rotation = (store.rotation + d) % 360; store._ops.push(['incrementRotation', d]); };
  store.setZOrder = (z) => {
    const list = store.__sheetShapes || [];
    const others = list.filter(s => s !== store.__proxy);
    if (String(z).includes('bringToFront')) { store.zOrderPosition = others.length; others.forEach((s, i) => { s.__store.zOrderPosition = i; }); }
    else if (String(z).includes('sendToBack')) { store.zOrderPosition = 0; others.forEach((s, i) => { s.__store.zOrderPosition = i + 1; }); }
    store._ops.push(['setZOrder', z]);
  };
  store.delete = () => { store.__deleted = true; const arr = store.__sheetShapes; if (arr) { const i = arr.indexOf(store.__proxy); if (i >= 0) arr.splice(i, 1); } };
  store.getAsImage = () => ({ value: 'iVBORw0KGgoAAAANSUhEUg==' });
  return proxy;
}

function makeSheet(name, shapes, sync) {
  return {
    name,
    shapes: {
      items: shapes,
      load() { shapes.itemList = { items: shapes }; return { items: shapes }; },
      getItem(n) { const s = this.items.find(x => x.__store.name === n); if (!s) throw new Error(`ItemNotFound: ${n}`); return s; },
      getActiveShape() { return this.items[0]; },
      addGeometricShape(type, box) {
        const s = makeShape(Object.assign({ name: 'Shape' + (idSeq + 1), type }, box || {}));
        s.__store.__sheetShapes = this.items; this.items.push(s); return s;
      },
      addLine(x1, y1, x2, y2) {
        const s = makeShape({ name: 'Line' + (idSeq + 1), type: 'Line', left: Math.min(x1, x2), top: Math.min(y1, y2), width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) });
        s.__store.__sheetShapes = this.items; this.items.push(s); return s;
      },
      addSvg() { const s = makeShape({ name: 'Svg' + (idSeq + 1), type: 'Image' }); s.__store.__sheetShapes = this.items; this.items.push(s); return s; },
      addTextBox(text) { const s = makeShape({ name: 'TextBox' + (idSeq + 1), type: 'TextBox' }); s.__store.textFrame.textRange.text = text; s.__store.__sheetShapes = this.items; this.items.push(s); return s; },
      addGroup(ids) {
        const members = this.items.filter(s => ids.includes(s.__store.__id));
        const s = makeShape({ name: 'Group' + (idSeq + 1), type: 'Group' });
        s.__store.group = {
          shapes: { items: members, load() { return { items: members }; } },
          ungroup() { members.forEach(m => { if (!this.__sheet.items.includes(m)) this.__sheet.items.push(m); }); const i = this.__sheet.items.indexOf(s); if (i >= 0) this.__sheet.items.splice(i, 1); }
        };
        s.__store.group.__sheet = this;
        s.__store.__sheetShapes = this.items;
        this.items.push(s); return s;
      }
    },
    getRange() { return { values: [[]], load() {} }; },
    load() {}, activate() {}, delete() {}
  };
}

const sheets = {};
const logs = [];
function getSheet(n) {
  if (!sheets[n]) sheets[n] = makeSheet(n, [], () => {});
  return sheets[n];
}

const context = {
  workbook: { worksheets: { items: [], getItem: getSheet, getActiveWorksheet: () => getSheet('Sheet1'), add: (n) => getSheet(n), load() { return this.items; } } },
  sync: async () => {
    // 把 pendingWrites 落到 store（模拟一次 sync 把 setter 生效）
    const apply = (list) => list.forEach(p => { const s = p.__store; const w = p.__pendingWrites; Object.keys(w).forEach(k => { s[k] = w[k]; }); Object.keys(w).forEach(k => delete w[k]); });
    Object.values(sheets).forEach(sh => { apply(sh.shapes.items); });
    logs.push('sync');
  }
};

const sandbox = {
  console,
  setTimeout, clearTimeout, setInterval, clearInterval,
  document: { getElementById: () => null, querySelector: () => null, createElement: () => ({ classList: { toggle() {}, contains: () => false }, style: {}, addEventListener() {} }), addEventListener: () => {}, readyState: 'complete' },
  window: { location: { reload() {}, search: '' }, addEventListener: () => {} },
  WebSocket: class { constructor() { this.readyState = 0; } send() {} close() {} },
  Office: { onReady: () => {}, HostType: { Excel: 'Excel' }, context: { requirements: { isSetSupported: () => true }, diagnostics: {} } },
  Excel: {
    run: async (fn) => await fn(context),
    GeometricShapeType: { rectangle: 'GeometricShapeType.rectangle', roundedRectangle: 'GeometricShapeType.roundedRectangle', oval: 'GeometricShapeType.oval', triangle: 'GeometricShapeType.triangle', star5: 'GeometricShapeType.star5', rightArrow: 'GeometricShapeType.rightArrow', flowChartProcess: 'GeometricShapeType.flowChartProcess', textBox: 'GeometricShapeType.textBox' },
    ShapeZOrder: { bringToFront: 'ShapeZOrder.bringToFront', sendToBack: 'ShapeZOrder.sendToBack', bringForward: 'ShapeZOrder.bringForward', sendBackward: 'ShapeZOrder.sendBackward' },
    PictureFormat: { png: 'PictureFormat.png', jpeg: 'PictureFormat.jpeg' },
    ShapeScaleType: { currentSize: 'ShapeScaleType.currentSize', originalSize: 'ShapeScaleType.originalSize' }
  },
  globalThis: null,
  URL: URL,
  btoa: (s) => Buffer.from(s, 'binary').toString('base64'),
  atob: (s) => Buffer.from(s, 'base64').toString('binary'),
  Function, Object, Array, JSON, Math, Number, String, Boolean, Date, Error, Promise, RegExp, isNaN, parseInt, parseFloat
};
sandbox.globalThis = sandbox;
vm.createContext(sandbox);

// getAsImage 需要返回可读的 base64
const realRun = sandbox.Excel.run;
sandbox.Excel.run = async (fn) => await realRun(async (ctx) => {
  const res = await fn(ctx);
  return res;
});

vm.runInContext(src, sandbox, { filename: 'taskpane.js' });
const H = sandbox.__CAP08;
if (!H) { console.error('FAIL: 未能从 taskpane.js 暴露形状处理函数'); process.exit(1); }

// ── 断言 ───────────────────────────────────────────────────────────────────
let pass = 0, fail = 0;
const check = (label, cond, extra) => {
  if (cond) { pass++; console.log('  ✓ ' + label); }
  else { fail++; console.log('  ✗ ' + label + (extra !== undefined ? '  → ' + JSON.stringify(extra) : '')); }
};

console.log('CAP-08 离线检查（mock 批处理语义，非真机）');
console.log('\n[1] 几何形状：写后读回');
const r1 = await H.handleAddShape({ sheetName: 'Sheet1', kind: 'geometric', shapeType: 'rounded_rectangle', shapeName: 'demo_rect', left: 20, top: 30, width: 160, height: 90, rotation: 15, fillColor: '#2f6feb', lineColor: '12305e', lineWeight: 2, text: '标题' });
check('success=true', r1.success === true, r1);
check('读回 name', r1.shape.name === 'demo_rect', r1.shape);
check('读回 left/top/width/height', r1.shape.left === 20 && r1.shape.top === 30 && r1.shape.width === 160 && r1.shape.height === 90, r1.shape);
check('读回 rotation', r1.shape.rotation === 15, r1.shape);
check('填充色归一为 #RRGGBB', r1.shape.fill.foregroundColor === '#2F6FEB', r1.shape.fill);
check('线条色归一', r1.shape.line.color === '#12305E', r1.shape.line);
check('文字写回', r1.shape.text === '标题', r1.shape);
check('verified=true', r1.verified === true, r1.warnings);
check('#RGB 短写展开', H.normalizeHexColor('#abc') === '#AABBCC', H.normalizeHexColor('#abc'));
check('非法颜色返回 null', H.normalizeHexColor('红色') === null);
check('别名 rounded_rectangle → roundedRectangle', H.resolveGeometricShapeType('rounded_rectangle').name === 'roundedRectangle');
check('未知类型报错且给出候选', (() => { try { H.resolveGeometricShapeType('nope'); return false; } catch (e) { return /可用枚举键/.test(e.message); } })());

console.log('\n[2] 直线 / SVG / 文本框');
const r2 = await H.handleAddShape({ sheetName: 'Sheet1', kind: 'line', shapeName: 'demo_line', x1: 200, y1: 20, x2: 360, y2: 110, lineColor: '#E4572E', lineWeight: 3 });
check('直线读回 name', r2.shape.name === 'demo_line', r2.shape);
check('直线包围盒 left=200 width=160', r2.shape.left === 200 && r2.shape.width === 160, r2.shape);
const r3 = await H.handleAddShape({ sheetName: 'Sheet1', kind: 'svg', shapeName: 'demo_svg', svg: '<svg xmlns="http://www.w3.org/2000/svg"/>' });
check('SVG 读回 name', r3.shape.name === 'demo_svg', r3.shape);
const r4 = await H.handleAddShape({ sheetName: 'Sheet1', kind: 'textbox', shapeName: 'demo_text', text: '说明文字', left: 20, top: 140, width: 200, height: 50 });
check('文本框读回文字', r4.shape.text === '说明文字', r4.shape);
check('SVG 缺参数时报错', await (async () => { try { await H.handleAddShape({ kind: 'svg' }); return false; } catch (e) { return /需要传 svg/.test(e.message); } })());

console.log('\n[3] 层级调整');
const rz1 = await H.handleSetShapeZOrder({ sheetName: 'Sheet1', name: 'demo_rect', zOrder: 'bringToFront' });
check('bringToFront 后读回层级序列', Array.isArray(rz1.zOrderBottomToTop) && rz1.zOrderBottomToTop.length >= 4, rz1);
check('未知层级动作报错', await (async () => { try { await H.handleSetShapeZOrder({ sheetName: 'Sheet1', name: 'demo_rect', zOrder: 'up' }); return false; } catch (e) { return /不认识的层级动作/.test(e.message); } })());

console.log('\n[4] 分组 / 解组');
const rg = await H.handleGroupShapes({ sheetName: 'Sheet1', shapeNames: ['demo_rect', 'demo_line', 'demo_svg'], groupName: 'demo_group' });
check('分组读回成员数=3', rg.memberCount === 3, rg);
check('分组 verified', rg.verified === true, rg.warnings);
const rug = await H.handleUngroupShapes({ sheetName: 'Sheet1', groupName: 'demo_group' });
check('解组读回子形状回归', rug.verified === true, rug);
check('分组少于 2 个报错', await (async () => { try { await H.handleGroupShapes({ shapeNames: ['a'] }); return false; } catch (e) { return /至少 2 个形状/.test(e.message); } })());

console.log('\n[5] 读回列表（含层级/填充）');
const rl = await H.handleListShapes({ sheetName: 'Sheet1' });
check('count >= 4', rl.count >= 4, rl.count);
check('每项带 fill/line/zOrderPosition', rl.shapes.every(s => s.fill && s.line && s.zOrderPosition !== undefined), rl.shapes);
check('zOrderBottomToTop 长度一致', rl.zOrderBottomToTop.length === rl.count, rl);

console.log('\n[6] 更新与删除（读回核对）');
const ru = await H.handleUpdateShape({ sheetName: 'Sheet1', name: 'demo_rect', left: 400, fillColor: '#FF0000', rotationDelta: 10 });
check('更新 changed=true', ru.changed === true, ru);
check('before/after 都返回', Boolean(ru.before) && Boolean(ru.shape), ru);
check('更新后 left=400', ru.shape.left === 400, ru.shape);
check('无字段时明确 changed=false', (await H.handleUpdateShape({ sheetName: 'Sheet1', name: 'demo_rect' })).changed === false);
const rd = await H.handleUpdateShape({ sheetName: 'Sheet1', name: 'demo_svg', action: 'delete' });
check('删除后 verified=true', rd.verified === true, rd);
check('删除后不在 remaining', !rd.remaining.includes('demo_svg'), rd.remaining);

console.log('\n[7] 形状定位错误如实报错');
check('不存在的形状名抛 ItemNotFound', await (async () => { try { await H.handleUpdateShape({ sheetName: 'Sheet1', name: '不存在' }); return false; } catch (e) { return /ItemNotFound/.test(e.message); } })());
check('缺 name/shapeId 抛参数错误', await (async () => { try { await H.handleUpdateShape({ sheetName: 'Sheet1' }); return false; } catch (e) { return /缺少必要参数/.test(e.message); } })());

console.log('\n[8] 能力自检编排（handleShapeSelfTest）');
const st = await H.handleShapeSelfTest({ sheetName: '_cap08_selftest_mock', keep: true });
check('自检返回 total/passed', typeof st.total === 'number' && typeof st.passed === 'number', st);
check('自检逐项都有 check 名', Array.isArray(st.checks) && st.checks.every(c => typeof c.check === 'string'), st.checks);
check('自检覆盖几何形状', st.checks.some(c => /addGeometricShape/.test(c.check)), st.checks.map(c => c.check));
check('自检覆盖直线', st.checks.some(c => /addLine/.test(c.check)));
check('自检覆盖 SVG', st.checks.some(c => /addSvg/.test(c.check)));
check('自检覆盖文本框', st.checks.some(c => /addTextBox/.test(c.check)));
check('自检覆盖层级', st.checks.some(c => /setZOrder/.test(c.check)));
check('自检覆盖分组', st.checks.some(c => /addGroup/.test(c.check)));
check('自检覆盖解组', st.checks.some(c => /ungroup/.test(c.check)));
check('自检覆盖导图', st.checks.some(c => /getAsImage/.test(c.check)));
check('自检覆盖缩放', st.checks.some(c => /scaleWidth/.test(c.check)));
check('自检覆盖删除读回', st.checks.some(c => /delete/.test(c.check)));
check('keep=true 时不清理', st.cleanedUp === false, st.cleanedUp);
const st2 = await H.handleShapeSelfTest({ sheetName: '_cap08_selftest_mock2' });
check('默认清理探测表', st2.cleanedUp === true, st2);

console.log(`\n结果：${pass} 通过 / ${fail} 失败`);
process.exit(fail === 0 ? 0 : 1);
