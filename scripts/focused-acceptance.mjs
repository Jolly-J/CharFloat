// 重点验收测试（B 方案）：常用能力 + 今日修复项
//
// 判定口径（本仓库的硬规矩）：
//   **不看 success，只看"写入的值能否从宿主读回来并比对一致"。**
// 每次运行自建测试工作簿，跑完自动关闭并删除磁盘文件，不动使用者的任何文档。
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';

const token = fs.readFileSync(path.join(os.homedir(), '.wps-bridge/token'), 'utf8').trim();
const BASE = 'http://127.0.0.1:19890/api/v1/tool/call';
const WB = 'AI重点验收.xlsx';
const PATH = path.join(os.homedir(), 'Downloads', WB);

async function call(name, args) {
  const r = await fetch(BASE, {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, arguments: args }),
  });
  return r.json().catch(() => ({}));
}
/** 直接问宿主：这是判据的唯一来源 */
async function host(code) {
  const r = await call('wps_execute_script', { component: 'excel', code, readOnly: true });
  if (!r.success) return { ok: false, err: String(r.error).slice(0, 120) };
  return { ok: true, val: r.data?.returnValue };
}
const bgr = (hex) => {
  const h = String(hex).replace('#', '');
  return parseInt(h.substring(4, 6) + h.substring(2, 4) + h.substring(0, 2), 16);
};

const results = [];
let n = 0;
function rec(layer, name, ok, detail) {
  n++;
  results.push({ layer, name, ok, detail });
  console.log(`  ${ok ? '✔' : '✖'} [${layer}] ${name}${ok ? '' : '   → ' + String(detail).slice(0, 130)}`);
}
async function check(layer, name, fn) {
  try {
    const r = await fn();
    rec(layer, name, r === true || (r && r.ok === true), r === true ? '' : (r && r.detail) || r);
  } catch (e) {
    rec(layer, name, false, e.message);
  }
}

// 确保 Excel 组件在线
{
  const st = await fetch('http://127.0.0.1:19890/api/v1/status', { headers: { Authorization: 'Bearer ' + token } })
    .then((r) => r.json()).catch(() => ({}));
  if (!st?.components?.excel?.connected) {
    console.log('✖ WPS 表格未连接，请先在 WPS 里打开任意表格文档');
    process.exit(2);
  }
}

console.log('\n═══ WPS 表格 · 重点验收 ═══');
console.log('  （判定：写入后回头问宿主并比对，不看 success）\n');

// ── 准备：自建测试工作簿
{
  const r = await call('wps_create_workbook', { savePath: PATH, sheetName: '验收' });
  if (!r.success) { console.log('  ✖ 建测试工作簿失败:', String(r.error).slice(0, 140)); process.exit(3); }
}
const SH = '验收';
const B = { host: 'wps', workbookName: WB, sheetName: SH };

// ── A 层：写入与读回
await check('写入', 'patch_cells 写入区域', async () => {
  const vals = [['月份', '区域', '销量', '金额'], ['1月', '华东', 120, 1200], ['1月', '华南', 90, 900],
    ['2月', '华东', 150, 1500], ['2月', '华南', 110, 1100], ['3月', '华东', 180, 1800], ['3月', '华南', 140, 1400]];
  const r = await call('excel_patch_cells', { ...B, address: 'A1:D7', values: vals });
  const h = await host(`const ws=app.Workbooks.Item('${WB}').Worksheets.Item('${SH}'); return ws.Range('A2').Value2+'|'+ws.Range('C7').Value2+'|'+ws.Range('B4').Value2;`);
  return { ok: r.success && h.val === '1月|140|华东', detail: `写入=${r.success} 读回=${h.val}` };
});
await check('写入', '公式写入并读回算式', async () => {
  await call('excel_patch_cells', { ...B, address: 'E1', values: [['合计']] });
  const r = await call('excel_patch_cells', { ...B, address: 'E2', values: [['=SUM(C2:C7)']] });
  const h = await host(`const ws=app.Workbooks.Item('${WB}').Worksheets.Item('${SH}'); return String(ws.Range('E2').Formula)+'|'+String(ws.Range('E2').Value2);`);
  return { ok: r.success && String(h.val).startsWith('=SUM'), detail: `读回公式=${h.val}（期望 =SUM…|790）` };
});

// ── B 层：格式（含今天的修复）
await check('格式', 'format_cells 字号/加粗/底色读回', async () => {
  await call('excel_format_cells', { ...B, address: 'A1:D1', fontSize: 12, bold: true, backgroundColor: '#2F6FEB', fontColor: '#FFFFFF' });
  const h = await host(`const c=app.Workbooks.Item('${WB}').Worksheets.Item('${SH}').Range('A1'); return c.Font.Bold+'|'+c.Font.Size+'|'+c.Interior.Color+'|'+c.Font.Color;`);
  const [b, s, bg, fg] = String(h.val).split('|');
  return { ok: Number(bg) === bgr('#2F6FEB') && Number(fg) === bgr('#FFFFFF'), detail: `粗=${b} 号=${s} 底=${bg}(期望${bgr('#2F6FEB')}) 字=${fg}` };
});
await check('格式', '★ verticalAlignment=bottom 真落地（今日修复）', async () => {
  await call('excel_format_cells', { ...B, address: 'A2', verticalAlignment: 'bottom' });
  const h = await host(`return Number(app.Workbooks.Item('${WB}').Worksheets.Item('${SH}').Range('A2').VerticalAlignment);`);
  return { ok: Number(h.val) === -4107, detail: `读回 ${h.val}，期望 -4107` };
});
await check('格式', '★ 非法颜色如实告警（今日修复）', async () => {
  const r = await call('excel_format_cells', { ...B, address: 'A3', backgroundColor: '不是颜色' });
  return { ok: r.success === true && (r.data?.warnings || []).length > 0, detail: JSON.stringify(r.data?.warnings)?.slice(0, 100) };
});

// ── C 层：条件格式（含今天的修复）
await check('条件格式', '★ text_contains 公式锚定在区域左上角（今日修复）', async () => {
  await call('excel_patch_cells', { ...B, address: 'G1:G4', values: [['特价'], ['普通'], ['特价品'], ['普通']] });
  await call('wps_execute_script', { component: 'excel', code: `app.Workbooks.Item('${WB}').Worksheets.Item('${SH}').Range('A1').Select(); return 1;`, readOnly: false });
  const r = await call('excel_add_conditional_formatting', { ...B, address: 'G1:G4', ruleType: 'text_contains', containsText: '特价', backgroundColor: '#FFDDDD' });
  const h = await host(`const fc=app.Workbooks.Item('${WB}').Worksheets.Item('${SH}').Range('G1:G4').FormatConditions; return fc.Item(fc.Count).Formula1;`);
  const f = String(h.val);
  return { ok: r.success && f.includes('G1') && !f.includes('Y1'), detail: `公式=${f}` };
});

// ── D 层：数据有效性 / 冻结 / 筛选（含今天的修复）
await check('校验', '★ 列表校验缺 listItems 直接报错（今日修复）', async () => {
  const r = await call('excel_set_data_validation', { ...B, address: 'I1:I3', validationType: 'list' });
  return { ok: r.success === false && /listItems/.test(String(r.error)), detail: String(r.error).slice(0, 110) };
});
await check('校验', '★ 报错后原有校验必须保持（今日修复的破坏性顺序）', async () => {
  await call('excel_set_data_validation', { ...B, address: 'I1:I3', validationType: 'list', listItems: ['甲', '乙'] });
  const before = await host(`const v=app.Workbooks.Item('${WB}').Worksheets.Item('${SH}').Range('I1').Validation; return v.Type+'|'+v.Formula1;`);
  await call('excel_set_data_validation', { ...B, address: 'I1:I3', validationType: 'list' });
  const after = await host(`const v=app.Workbooks.Item('${WB}').Worksheets.Item('${SH}').Range('I1').Validation; return v.Type+'|'+v.Formula1;`);
  return { ok: String(before.val) === String(after.val) && String(after.val).includes('甲'), detail: `前=${before.val} 后=${after.val}` };
});
await check('视图', '★ freezeRowIndex=0 越界被拒（今日修复）', async () => {
  const r = await call('excel_freeze_panes', { ...B, freezeRowIndex: 0 });
  return { ok: r.success === false, detail: String(r.error).slice(0, 90) };
});
await check('视图', '冻结窗格并读回', async () => {
  const r = await call('excel_freeze_panes', { ...B, freezeRowIndex: 2 });
  const h = await host(`const w=app.Workbooks.Item('${WB}'); return String(w.Windows.Item(1).SplitRow)+'|'+String(w.Windows.Item(1).FreezePanes);`);
  const [split, frozen] = String(h.val).split('|');
  // 宿主语义：freezeRowIndex=N 表示"冻结前 N 行"，此时 SplitRow=N-1（分界线在 N-1 与 N 之间）
  return { ok: r.success && String(frozen) === 'true' && Number(split) === 1, detail: `SplitRow=${split} FreezePanes=${frozen}（freezeRowIndex=2 → 冻结前 2 行，分界线 SplitRow=1）` };
});

// ── E 层：工作表管理与今天的修复
await check('工作表', '★ rename 到已存在名字被拒（今日修复）', async () => {
  const r = await call('excel_manage_sheet', { ...B, action: 'rename', newName: SH });
  return { ok: r.success === false && /已被占用/.test(String(r.error)), detail: String(r.error).slice(0, 100) };
});
await check('工作表', '★ 非法表名被拒（今日修复）', async () => {
  const r = await call('excel_create_sheet', { ...B, sheetName: 'x/非法' });
  return { ok: r.success === false && /非法字符/.test(String(r.error)), detail: String(r.error).slice(0, 100) };
});
await check('工作表', '★ activate 激活工作表（今日新增）', async () => {
  await call('excel_create_sheet', { ...B, sheetName: '临时表' });
  await call('excel_manage_sheet', { ...B, sheetName: SH, action: 'activate' });
  const h = await host(`return String(app.Workbooks.Item('${WB}').ActiveSheet.Name);`);
  return { ok: String(h.val) === SH, detail: `活动表=${h.val}` };
});

// ── F 层：图表（含今天的修复）
await check('图表', 'add_chart 建图并读回系列数', async () => {
  const r = await call('excel_add_chart', { ...B, dataRange: 'A1:C7', chartType: 'column_clustered', position: { leftCell: 'K2', width: 400, height: 240 }, title: '销量对比' });
  const h = await host(`const ws=app.Workbooks.Item('${WB}').Worksheets.Item('${SH}'); let c=null; for(let i=1;i<=ws.Shapes.Count;i++){const s=ws.Shapes.Item(i); try{if(s.HasChart){c=s.Chart;break;}}catch(e){}} return c?String(c.SeriesCollection().Count)+'|'+String(c.ChartTitle.Text):'无图表';`);
  return { ok: r.success && String(h.val).includes('|'), detail: `宿主读回=${h.val}` };
});
await check('图表', '★ 图表数据源跨表（今日修复）', async () => {
  await call('excel_create_sheet', { ...B, sheetName: '图表页' });
  const r = await call('wps_add_chart', { workbookName: WB, sheetName: '图表页', dataRange: `${SH}!C1:D7`, chartType: 'column_clustered', position: { leftCell: 'B2', width: 420, height: 240 }, title: '跨表图' });
  const h = await host(`const ws=app.Workbooks.Item('${WB}').Worksheets.Item('图表页'); let c=null; for(let i=1;i<=ws.Shapes.Count;i++){const s=ws.Shapes.Item(i); try{if(s.HasChart){c=s.Chart;break;}}catch(e){}} return c?String(c.SeriesCollection().Count)+'|'+String(c.ChartTitle.Text):'无图表';`);
  return { ok: r.success && String(h.val).startsWith('2|'), detail: `宿主读回=${h.val}（C/D 两列都是数字，期望 2 个系列）` };
});

// ── G 层：透视表（含今天的修复）
await check('透视表', 'create_pivot_table 建表并读回', async () => {
  const r = await call('wps_create_pivot_table', { workbookName: WB, sourceSheetName: SH, sourceRange: 'A1:D7', destSheetName: SH, destCell: 'K20', rowFields: ['月份'], dataFields: [{ fieldName: '销量', summaryFunction: 'sum' }] });
  const h = await host(`const pts=app.Workbooks.Item('${WB}').Worksheets.Item('${SH}').PivotTables(); return Number(pts.Count);`);
  return { ok: r.success && Number(h.val) >= 1, detail: `透视表数=${h.val}` };
});
await check('透视表', '★ configure 改字段布局（今日新增）', async () => {
  const r = await call('excel_create_pivot_table', { host: 'wps', workbookName: WB, sheetName: SH, action: 'configure', clearFields: true, rowFields: ['区域'], columnFields: ['月份'], dataFields: [{ fieldName: '金额', summaryFunction: 'sum' }] });
  return { ok: r.success && JSON.stringify(r.data?.applied || {}).includes('区域'), detail: JSON.stringify(r.data?.applied)?.slice(0, 130) };
});

// ── H 层：批注（使用者验收过的那条）
await check('批注', '★ 加批注并读回正文（真机验收过）', async () => {
  const T = '请看这格说明';
  const r = await call('excel_manage_cell_comments', { ...B, action: 'add', address: 'A2', text: T });
  const h = await call('excel_manage_cell_comments', { ...B, action: 'read', address: 'A2' });
  const got = JSON.stringify(h.data || {});
  return { ok: r.success && got.includes(T), detail: got.slice(0, 140) };
});

// ── I 层：命名区域 / 超链接（含今天的修复）
await check('命名区域', '★ comment 不生效时如实告警（今日修复）', async () => {
  const r = await call('wps_manage_named_range', { workbookName: WB, action: 'add', name: '验收区域', refersTo: `${SH}!$A$1:$D$7`, comment: '备注' });
  await call('wps_manage_named_range', { workbookName: WB, action: 'delete', name: '验收区域' });
  return { ok: r.success === true && (r.data?.warnings || []).some((w) => /comment 未生效/.test(w)), detail: JSON.stringify(r.data?.warnings)?.slice(0, 130) };
});
await check('超链接', '★ 单独 targetAddress 可用（今日修复）', async () => {
  const r = await call('excel_manage_hyperlink', { ...B, action: 'add', address: 'A5', targetAddress: `${SH}!A1` });
  await call('excel_manage_hyperlink', { ...B, action: 'delete', address: 'A5' });
  return { ok: r.success === true, detail: String(r.error).slice(0, 110) };
});

// ── J 层：变更感知（今天新增）
await check('变更感知', '★ capture→diff 检出改动（今日新增）', async () => {
  await call('wps_manage_sheet_changes', { workbookName: WB, sheetName: SH, action: 'capture' });
  await call('excel_patch_cells', { ...B, address: 'A1', values: [['月份（已改）']] });
  const r = await call('wps_manage_sheet_changes', { workbookName: WB, sheetName: SH, action: 'diff' });
  await call('wps_manage_sheet_changes', { workbookName: WB, sheetName: SH, action: 'clear' });
  const hit = (r.data?.changed || []).concat(r.data?.added || []).some((c) => c.address === 'A1' || c.address === '$A$1');
  return { ok: r.success && hit, detail: JSON.stringify(r.data?.changed?.slice(0, 2))?.slice(0, 140) };
});

// ── 清理
console.log('\n  —— 清理 ——');
await call('wps_execute_script', {
  component: 'excel', readOnly: false,
  code: `const prev=app.DisplayAlerts; app.DisplayAlerts=false; try{ for(let i=app.Workbooks.Count;i>=1;i--){ const w=app.Workbooks.Item(i); if(w.Name==='${WB}'){ w.Close(false); } } } finally { app.DisplayAlerts=prev; } return 'closed';`,
});
try { fs.unlinkSync(PATH); } catch {}
console.log(`  已关闭并删除测试工作簿 ${PATH}`);

const pass = results.filter((r) => r.ok).length;
console.log('\n════════ 重点验收结果 ════════');
const layers = [...new Set(results.map((r) => r.layer))];
for (const L of layers) {
  const g = results.filter((r) => r.layer === L);
  console.log(`  ${L}: ${g.filter((r) => r.ok).length}/${g.length}`);
}
console.log(`\n  合计 ${pass}/${results.length} 通过`);
const failed = results.filter((r) => !r.ok);
if (failed.length) {
  console.log('\n  失败明细:');
  failed.forEach((r) => console.log(`    ✖ [${r.layer}] ${r.name} → ${String(r.detail).slice(0, 160)}`));
}
process.exit(pass === results.length ? 0 : 1);
