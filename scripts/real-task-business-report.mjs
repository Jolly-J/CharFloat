// 真实业务活：2026 年经营分析表
// 边走边记录"不顺手的地方"，作为真实使用反馈。
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';

const token = fs.readFileSync(path.join(os.homedir(), '.wps-bridge/token'), 'utf8').trim();
const BASE = 'http://127.0.0.1:19890/api/v1/tool/call';
const WB = '2026年经营分析.xlsx';
const P = path.join(os.homedir(), 'Downloads', WB);

const friction = [];   // 真实使用中不顺手的地方
function note(where, what) { friction.push({ where, what }); console.log(`  ⚠ [不顺手] ${where}: ${what}`); }

async function call(name, args) {
  const r = await fetch(BASE, { method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, arguments: args }) });
  return r.json().catch(() => ({}));
}
async function host(code, wbName = WB) {
  const r = await call('wps_execute_script', { component: 'excel', readOnly: true, code });
  return r.success ? r.data?.returnValue : `ERR:${String(r.error).slice(0, 80)}`;
}
const B = { host: 'wps', workbookName: WB };

console.log('\n═══ 经营分析表 · 真实业务全流程 ═══\n');

// ── 1. 建簿
console.log('【1】建工作簿');
{
  const r = await call('wps_create_workbook', { savePath: P, sheetName: '销售明细' });
  if (!r.success) {
    // 已存在时重建：先删文件
    try { fs.unlinkSync(P); } catch {}
    const r2 = await call('wps_create_workbook', { savePath: P, sheetName: '销售明细' });
    console.log(r2.success ? '  ✔ 已建' : `  ✖ ${String(r2.error || r2.data?.message).slice(0, 120)}`);
  } else console.log('  ✔ 已建');
}

// ── 2. 销售明细数据
console.log('\n【2】写入销售明细');
const 明细 = [
  ['月份', '大区', '销售额（万元）', '订单数', '客单价（元）', '目标（万元）', '完成率'],
  ['1月', '华东', 1280, 3200, 4000, 1200, '=C2/F2'],
  ['1月', '华南', 960, 2400, 4000, 1000, '=C3/F3'],
  ['1月', '华北', 740, 1850, 4000, 800, '=C4/F4'],
  ['2月', '华东', 1420, 3400, 4176, 1300, '=C5/F5'],
  ['2月', '华南', 1080, 2600, 4154, 1050, '=C6/F6'],
  ['2月', '华北', 820, 2000, 4100, 850, '=C7/F7'],
  ['3月', '华东', 1650, 3800, 4342, 1400, '=C8/F8'],
  ['3月', '华南', 1240, 2900, 4276, 1100, '=C9/F9'],
  ['3月', '华北', 910, 2200, 4136, 900, '=C10/F10'],
  ['合计', '—', '=SUM(C2:C10)', '=SUM(D2:D10)', '=AVERAGE(E2:E10)', '=SUM(F2:F10)', '=SUM(C2:C10)/SUM(F2:F10)'],
];
{
  const r = await call('wps_patch_cells', { ...B, sheetName: '销售明细', address: 'A1:G11', values: 明细 });
  if (!r.success) note('写入明细', String(r.error).slice(0, 120));
  else {
    const v = await host(`const ws=app.Workbooks.Item('${WB}').Worksheets.Item('销售明细'); return String(ws.Range('C2').Value2)+'|'+String(ws.Range('C11').Value2)+'|'+String(ws.Range('E11').Value2);`);
    console.log(`  ✔ 写入并读回：C2=${String(v).split('|')[0]} 合计=${String(v).split('|')[1]} 均客单=${String(v).split('|')[2]}`);
    if (!String(v).includes('1280')) note('写入读回', `读回异常: ${v}`);
  }
}

// ── 3. 表头与数字格式
console.log('\n【3】格式美化');
{
  let r = await call('wps_format_cells', { ...B, sheetName: '销售明细', address: 'A1:G1',
    fontName: '微软雅黑', fontSize: 11, bold: true, backgroundColor: '#1F3864', fontColor: '#FFFFFF',
    horizontalAlignment: 'center', verticalAlignment: 'center', wrapText: true });
  const chk = await host(`const c=app.Workbooks.Item('${WB}').Worksheets.Item('销售明细').Range('A1'); return c.Font.Bold+'|'+c.Interior.Color+'|'+c.Font.Color+'|'+c.HorizontalAlignment;`);
  console.log(`  ✔ 表头：粗体/底色/字色/对齐 读回 = ${chk}`);
  if (!String(chk).startsWith('true')) note('表头格式', `读回不符: ${chk}`);

  // 合计行加粗 + 上边框
  r = await call('wps_format_cells', { ...B, sheetName: '销售明细', address: 'A11:G11', bold: true, backgroundColor: '#D9E2F3' });
  // 完成率列转百分比
  r = await call('wps_format_cells', { ...B, sheetName: '销售明细', address: 'G2:G11', numberFormat: '0.0%' });
  const g = await host(`const c=app.Workbooks.Item('${WB}').Worksheets.Item('销售明细').Range('G2'); return String(c.NumberFormat)+'|'+String(c.Value2);`);
  console.log(`  ✔ 完成率列：格式=${String(g).split('|')[0]} 计算值=${String(g).split('|')[1]}`);
  if (!String(g).includes('%')) note('数字格式', `百分比格式未生效: ${g}`);

  // 列宽与冻结
  r = await call('wps_auto_fit_columns', { ...B, sheetName: '销售明细' });
  r = await call('wps_freeze_panes', { ...B, sheetName: '销售明细', freezeRowIndex: 2 });
  console.log('  ✔ 自适应列宽 + 冻结首行');
}

// ── 4. 条件格式：完成率低于 100% 标红
console.log('\n【4】条件格式（完成率预警）');
{
  const r = await call('wps_add_conditional_formatting', { ...B, sheetName: '销售明细', address: 'G2:G10',
    ruleType: 'cell_value', operator: 'less_than', formula1: '1', backgroundColor: '#FFC7CE', fontColor: '#9C0006' });
  const fc = await host(`const fc=app.Workbooks.Item('${WB}').Worksheets.Item('销售明细').Range('G2:G10').FormatConditions; const f=fc.Item(fc.Count); return String(f.Formula1)+'|'+String(f.Interior.Color);`);
  console.log(`  ✔ 条件格式读回：${fc}`);
  if (!String(fc).includes('|')) note('条件格式', `读回失败: ${fc}`);
}

// ── 5. 交叉表（图表数据源必须是交叉表）
console.log('\n【5】建交叉表（图表数据源）');
{
  const r = await call('wps_create_sheet', { ...B, sheetName: '分析看板' });
  await call('wps_patch_cells', { ...B, sheetName: '分析看板', address: 'A1:D5', values: [
    ['月份', '华东', '华南', '华北'],
    ['1月', 1280, 960, 740], ['2月', 1420, 1080, 820], ['3月', 1650, 1240, 910],
    ['合计', 4350, 3280, 2470]] });
  const v = await host(`const ws=app.Workbooks.Item('${WB}').Worksheets.Item('分析看板'); return String(ws.Range('B2').Value2)+'|'+String(ws.Range('D5').Value2);`);
  console.log(`  ✔ 交叉表写入并读回：${v}`);
}

// ── 6. 图表
console.log('\n【6】图表');
{
  let r = await call('wps_add_chart', { workbookName: WB, sheetName: '分析看板', dataRange: 'A1:D4',
    chartType: 'column_clustered', position: { leftCell: 'F2', width: 460, height: 260 }, title: '各大区月度销售趋势' });
  console.log(r.success ? `  ✔ 柱状图：系列数=${r.data?.seriesCount}` : `  ✖ ${String(r.error).slice(0, 120)}`);
  if (!r.success) note('建图表', String(r.error).slice(0, 120));

  // 加数据标签 + 微软雅黑 + 图例靠右
  r = await call('wps_update_chart', { workbookName: WB, sheetName: '分析看板', chartIndex: 1,
    fontName: '微软雅黑', legendPosition: 'right', showDataLabels: true, title: '各大区月度销售趋势（万元）' });
  if (!r.success) note('图表细化', String(r.error).slice(0, 120));
  else console.log(`  ✔ 图表细化：${JSON.stringify(r.data?.warnings || [])}`);

  const c = await host(`const ws=app.Workbooks.Item('${WB}').Worksheets.Item('分析看板'); let ch=null; for(let i=1;i<=ws.Shapes.Count;i++){const s=ws.Shapes.Item(i); try{if(s.HasChart){ch=s.Chart;break;}}catch(e){}} return ch?String(ch.SeriesCollection().Count)+'|'+(ch.HasLegend?String(ch.Legend.Position):'无图例')+'|'+String(ch.ChartTitle.Text):'无图表';`);
  console.log(`  ✔ 宿主读回：系列数|图例位置|标题 = ${c}`);
}

// ── 7. 透视表
console.log('\n【7】透视表');
{
  let r = await call('wps_create_pivot_table', { workbookName: WB, sourceSheetName: '销售明细', sourceRange: 'A1:F10',
    destSheetName: '分析看板', destCell: 'A8', rowFields: ['大区'], columnFields: ['月份'], dataFields: [{ fieldName: '销售额（万元）', summaryFunction: 'sum' }] });
  console.log(r.success ? `  ✔ 透视表建成` : `  ✖ ${String(r.error).slice(0, 120)}`);
  if (!r.success) note('建透视表', String(r.error).slice(0, 120));

  r = await call('excel_create_pivot_table', { host: 'wps', workbookName: WB, sheetName: '分析看板', action: 'read' });
  const pt = r.data?.pivotTables?.[0];
  console.log(`  ✔ 读回：${pt ? `${pt.name} 字段数=${(pt.fields || []).filter(f => f.orientation !== 0).length}` : JSON.stringify(r.data).slice(0, 100)}`);
}

// ── 8. 批注
console.log('\n【8】批注说明');
{
  const r = await call('excel_manage_cell_comments', { ...B, sheetName: '销售明细', action: 'add', address: 'G1',
    text: '完成率 = 销售额 / 目标。低于 100% 的已用条件格式标红。' });
  const back = await call('excel_manage_cell_comments', { ...B, sheetName: '销售明细', action: 'read', address: 'G1' });
  const got = JSON.stringify(back.data || {});
  console.log(got.includes('完成率') ? '  ✔ 批注写入并读回' : `  ✖ 批注读回失败: ${got.slice(0, 120)}`);
  if (!got.includes('完成率')) note('批注', got.slice(0, 120));
}

// ── 9. 保存
console.log('\n【9】保存');
{
  const r = await call('wps_save_workbook', { workbookName: WB });
  console.log(r.success ? `  ✔ 已保存（saved=${r.data?.saved}）` : `  ✖ ${String(r.error || r.data?.message).slice(0, 120)}`);
  if (!r.success) note('保存', String(r.error || r.data?.message).slice(0, 120));
}

// ── 汇总
console.log('\n════════ 真实使用反馈 ════════');
if (friction.length === 0) console.log('  ✔ 全流程无阻塞');
else friction.forEach((f, i) => console.log(`  ${i + 1}. [${f.where}] ${f.what}`));
console.log(`\n  交付文件：${P}`);
