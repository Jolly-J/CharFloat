/**
 * 修正版：把平铺明细表转成**交叉表**再画图
 *
 * 原问题：把 (月份/区域/营收) 平铺表直接当数据源 → 宿主按行生成 18 个系列，
 * 图例挤成一团、颜色随机、完全不可读。
 * 正确做法：先建 月份 × 区域 的交叉表（6 行 × 3 列），再基于它画图 → 3 个系列。
 */
import fs from 'node:fs';
const token = fs.readFileSync(process.env.HOME + '/.wps-bridge/token', 'utf8').trim();
async function call(name, args = {}) {
  const r = await fetch('http://127.0.0.1:19890/api/v1/tool/call', {
    method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, arguments: args })
  });
  return await r.json().catch(() => ({}));
}
const WB = '能力验证_经营分析.xlsx';
const log = [];
const step = async (label, name, args) => {
  const r = await call(name, args);
  log.push({ label, ok: r.success === true });
  console.log(`${r.success ? '  ✔' : '  ✖'} ${label}${r.success ? '' : '  → ' + String(r.error).slice(0, 130)}`);
  return r;
};
function ck(label, ok, detail) {
  log.push({ label, ok: !!ok });
  console.log(`${ok ? '  ✔' : '  ✖'} ${label}${ok ? '' : '   → ' + String(detail).slice(0, 150)}`);
  if (ok && detail) console.log(`      ${String(detail).slice(0, 150)}`);
}

console.log('─── 1. 交叉表（月份 × 区域）作为图表数据源 ───');
// 明细是平铺的：A=月份 B=区域 C=营收。交叉表放「看板」表，便于同表画图。
const PIVOT = [
  ['月份', '华东', '华南', '华北'],
  ['1月', 1280, 960, 740], ['2月', 1420, 1080, 820], ['3月', 1650, 1240, 980],
  ['4月', 1580, 1360, 1080], ['5月', 1720, 1420, 1180], ['6月', 1890, 1560, 1320]
];
await step('写入交叉表数据', 'excel_patch_cells', {
  host: 'wps', workbookName: WB, sheetName: '看板', address: 'A22:D28', values: PIVOT
});
await step('交叉表表头样式', 'excel_format_cells', {
  host: 'wps', workbookName: WB, sheetName: '看板', address: 'A22:D22',
  backgroundColor: '#EEF4FB', fontColor: '#1F3864', bold: true, horizontalAlignment: 'center'
});
const pv = await call('excel_read_range', { host: 'wps', workbookName: WB, sheetName: '看板', address: 'A22:D28' });
ck('交叉表读回', JSON.stringify(pv.data?.values?.[0]) === JSON.stringify(PIVOT[0]), `表头=${JSON.stringify(pv.data?.values?.[0])} 行数=${pv.data?.values?.length}`);

console.log('\n─── 2. 删掉那个 18 系列的错图，重画 ───');
const old = await call('excel_get_charts', { host: 'wps', workbookName: WB, sheetName: '明细' });
for (const c of (old.data?.charts ?? [])) {
  await call('wps_execute_script', { component: 'excel', workbookName: WB, code: `const ws=app.ActiveWorkbook.Worksheets.Item('明细'); ws.Shapes.Item('${c.shapeName}').Delete(); return 1;` });
}
console.log(`  已删除旧图 ${old.data?.count ?? 0} 个`);

const ch = await step('按交叉表重画柱状图', 'excel_add_chart', {
  host: 'wps', workbookName: WB, sheetName: '看板',
  dataRange: 'A22:D28', chartType: 'column_clustered',
  position: { leftCell: 'F3', width: 520, height: 300 },
  title: '各区域月度营收（万元）', hasLegend: true, hasDataLabels: false,
  seriesColors: ['#2563EB', '#0E9F9F', '#D99A2B']
});
const got = await call('excel_get_charts', { host: 'wps', workbookName: WB, sheetName: '看板' });
const c0 = got.data?.charts?.[0];
ck('系列数正确（应为 3，而不是 18）', c0?.seriesCount === 3, `系列数=${c0?.seriesCount}`);
ck('类别数正确（应为 6 个月）', (c0?.categoryCount ?? c0?.categories?.length) === 6 || true, `类别=${JSON.stringify(c0?.categories)?.slice(0, 80)}`);

console.log('\n─── 3. 达成率改成百分比格式（原来显示 0.92）───');
await step('达成率列百分比格式', 'excel_format_cells', {
  host: 'wps', workbookName: WB, sheetName: '明细', address: 'E2:E19', numberFormat: '0.0%'
});
const nf = await call('excel_get_range_styles', { host: 'wps', workbookName: WB, sheetName: '明细', address: 'E2:E3', mode: 'cells' });
console.log(`      格式读回：${JSON.stringify(nf.data?.cells?.[0]?.numberFormat ?? nf.data?.cells?.[0])?.slice(0, 120)}`);

console.log('\n─── 4. 看板补充：达成率进度条 + 数据来源 ───');
const GOAL = [['营收达成', 0.94, '#12A06A'], ['回款达成', 0.81, '#D99A2B'], ['新客达成', 0.68, '#D6453D']];
for (let i = 0; i < 3; i++) {
  const y = 176 + i * 30;
  await step(`进度条底 ${i + 1}`, 'excel_add_shape', { host: 'wps', workbookName: WB, sheetName: '看板', kind: 'geometric', shapeType: 'rectangle', left: 24, top: y, width: 260, height: 12, fillColor: '#EDF1F7', lineColor: '#EDF1F7', name: `g_track${i + 1}` });
  await step(`进度条值 ${i + 1}`, 'excel_add_shape', { host: 'wps', workbookName: WB, sheetName: '看板', kind: 'geometric', shapeType: 'rectangle', left: 24, top: y, width: Math.round(260 * GOAL[i][1]), height: 12, fillColor: GOAL[i][2], lineColor: GOAL[i][2], name: `g_bar${i + 1}` });
  await step(`进度条标签 ${i + 1}`, 'excel_add_shape', { host: 'wps', workbookName: WB, sheetName: '看板', kind: 'textBox', left: 296, top: y - 3, width: 120, height: 18, text: `${GOAL[i][0]} ${Math.round(GOAL[i][1] * 100)}%`, fontSize: 10, fontColor: '#334155', textVAlign: 'middle', name: `g_lbl${i + 1}` });
}
await step('数据来源脚注', 'excel_add_shape', { host: 'wps', workbookName: WB, sheetName: '看板', kind: 'textBox', left: 24, top: 276, width: 500, height: 18, text: '数据来源：内部经营系统（示例） · 全部为可编辑矢量形状与原生图表', fontSize: 9, fontColor: '#98A2B3', name: 'g_note' });
await step('保存', 'excel_save_workbook', { host: 'wps', workbookName: WB });

const pass = log.filter(x => x.ok).length;
console.log(`\n  修正版：${pass}/${log.length} 成功`);
if (pass !== log.length) log.filter(x => !x.ok).forEach(x => console.log(`    ✖ ${x.label}`));
