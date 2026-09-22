/**
 * 能力验证：从零建一个「季度经营分析」工作簿，把已交付能力全用一遍
 *
 * 只用 MCP 工具（不含脚本），每一步都先写后读。
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
/** 断言（与 step 分开：step 看调用是否成功，ck 看**结果对不对**） */
function ck(group, name, ok, detail) {
  log.push({ label: `[${group}] ${name}`, ok: !!ok, name: 'ck' });
  console.log(`${ok ? '  ✔' : '  ✖'} ${name}${ok ? '' : '   → ' + String(detail).slice(0, 140)}`);
  if (ok && detail) console.log(`      ${String(detail).slice(0, 140)}`);
}
const step = async (label, name, args) => {
  const r = await call(name, args);
  const ok = r.success === true;
  log.push({ label, ok, name });
  console.log(`${ok ? '  ✔' : '  ✖'} ${label}${ok ? '' : '  → ' + String(r.error).slice(0, 140)}`);
  return r;
};

console.log('═══ 一、明细表：数据 + 表格对象 + 校验 + 条件格式 + 冻结 + 筛选 ═══');
// 用新工具从零建工作簿（宿主本来就支持，此前只是没暴露成工具）
await step('新建工作簿（wps_create_workbook）', 'wps_create_workbook', { savePath: `/Users/jolin/Downloads/${WB}`, sheetName: '明细' });
await step('确认首表名', 'wps_execute_script', { component: 'excel', workbookName: WB, code: `const w = app.ActiveWorkbook; return w.Name + ' / ' + w.Worksheets.Item(1).Name;` });

const HEAD = ['月份', '区域', '营收(万元)', '订单数', '达成率'];
const ROWS = [
  ['1月', '华东', '1280', '342', '0.92'], ['1月', '华南', '960', '268', '0.81'], ['1月', '华北', '740', '196', '0.74'],
  ['2月', '华东', '1420', '368', '1.02'], ['2月', '华南', '1080', '292', '0.91'], ['2月', '华北', '820', '214', '0.82'],
  ['3月', '华东', '1650', '402', '1.18'], ['3月', '华南', '1240', '326', '1.05'], ['3月', '华北', '980', '248', '0.98'],
  ['4月', '华东', '1580', '388', '1.13'], ['4月', '华南', '1360', '352', '1.15'], ['4月', '华北', '1080', '272', '1.08'],
  ['5月', '华东', '1720', '418', '1.23'], ['5月', '华南', '1420', '364', '1.20'], ['5月', '华北', '1180', '296', '1.18'],
  ['6月', '华东', '1890', '446', '1.35'], ['6月', '华南', '1560', '398', '1.32'], ['6月', '华北', '1320', '328', '1.32'],
];
await step('写入 18 行明细', 'excel_patch_cells', {
  host: 'wps', workbookName: WB, sheetName: '明细', address: `A1:E${ROWS.length + 1}`,
  values: [HEAD, ...ROWS]
});
const back = await call('excel_read_range', { host: 'wps', workbookName: WB, sheetName: '明细', address: 'A1:E19' });
console.log(`      读回核对：${JSON.stringify(back.data?.values?.[0])} … 末行 ${JSON.stringify(back.data?.values?.[18])}`);

await step('表头样式（深底白字加粗居中）', 'excel_format_cells', {
  host: 'wps', workbookName: WB, sheetName: '明细', address: 'A1:E1',
  backgroundColor: '#1F3864', fontColor: '#FFFFFF', bold: true, horizontalAlignment: 'center'
});
await step('结构化表格（含汇总行）', 'excel_manage_table', {
  host: 'wps', workbookName: WB, sheetName: '明细', action: 'apply',
  address: `A1:E${ROWS.length + 1}`, tableName: 'tblRevenue', hasHeaders: true, totalsRow: false
});
const tbl = await call('excel_manage_table', { host: 'wps', workbookName: WB, sheetName: '明细', action: 'list' });
console.log(`      表格读回：${JSON.stringify(tbl.data?.tables?.[0]?.columns)} / ${tbl.data?.tables?.[0]?.rowCount} 行`);

await step('数据有效性：区域下拉', 'excel_set_data_validation', {
  host: 'wps', workbookName: WB, sheetName: '明细', address: 'B2:B19',
  validationType: 'list', listItems: ['华东', '华南', '华北', '西南', '西北']
});
const dv = await call('excel_set_data_validation', { host: 'wps', workbookName: WB, sheetName: '明细', action: 'read', address: 'B2:B19' });
console.log(`      校验读回：type=${dv.data?.validations?.[0]?.type} 候选=${dv.data?.validations?.[0]?.formula1}`);

await step('条件格式：达成率色阶', 'excel_add_conditional_formatting', {
  host: 'wps', workbookName: WB, sheetName: '明细', address: 'E2:E19', ruleType: 'color_scale',
  colorScaleMin: '#F8696B', colorScaleMax: '#63BE7B'
});
const cf = await call('excel_add_conditional_formatting', { host: 'wps', workbookName: WB, sheetName: '明细', action: 'read', address: 'E2:E19' });
console.log(`      条件格式读回：${cf.data?.ruleCount} 条规则 / 类型=${cf.data?.areas?.[0]?.rules?.[0]?.type}`);

await step('冻结首行', 'excel_freeze_panes', { host: 'wps', workbookName: WB, sheetName: '明细', freezeRowIndex: 2 });
const fp = await call('excel_freeze_panes', { host: 'wps', workbookName: WB, sheetName: '明细', action: 'read' });
console.log(`      冻结读回：frozen=${fp.data?.frozen} 行=${fp.data?.freezeRowIndex}`);
await step('自动筛选', 'excel_set_filter_and_sort', { host: 'wps', workbookName: WB, sheetName: '明细', range: 'A1:E19', enableAutoFilter: true });

console.log('\n═══ 二、看板表：矢量形状 + 原生图表 + 读回 ═══');
await step('新建「看板」表', 'excel_create_sheet', { host: 'wps', workbookName: WB, sheetName: '看板' });
await step('隐藏网格线（画布前置）', 'excel_set_sheet_view', { host: 'wps', workbookName: WB, sheetName: '看板', showGridlines: false });
await step('标题条', 'excel_add_shape', { host: 'wps', workbookName: WB, sheetName: '看板', kind: 'geometric', shapeType: 'rounded_rectangle', left: 20, top: 16, width: 560, height: 46, fillColor: '#1F3864', text: '2026 上半年经营分析', fontSize: 16, bold: true, textAlign: 'left', marginLeft: 16, name: 'k_title' });
const KPI = [['1.92 亿', '累计营收', '#12A06A'], ['2,436', '累计订单', '#2563EB'], ['1.11', '平均达成率', '#D99A2B']];
for (let i = 0; i < 3; i++) {
  const x = 20 + i * 192;
  await step(`KPI 卡片 ${i + 1}`, 'excel_add_shape', { host: 'wps', workbookName: WB, sheetName: '看板', kind: 'geometric', shapeType: 'rounded_rectangle', left: x, top: 78, width: 176, height: 74, fillColor: '#EEF4FB', lineColor: '#C7DCF3', name: `k_card${i + 1}` });
  await step(`KPI 数值 ${i + 1}`, 'excel_add_shape', { host: 'wps', workbookName: WB, sheetName: '看板', kind: 'textBox', left: x + 14, top: 88, width: 150, height: 30, text: KPI[i][0], fontSize: 18, bold: true, fontColor: KPI[i][2], textVAlign: 'middle', name: `k_num${i + 1}` });
  await step(`KPI 标签 ${i + 1}`, 'excel_add_shape', { host: 'wps', workbookName: WB, sheetName: '看板', kind: 'textBox', left: x + 14, top: 120, width: 150, height: 18, text: KPI[i][1], fontSize: 10, fontColor: '#6B7A90', textVAlign: 'middle', name: `k_lbl${i + 1}` });
}
await step('艺术字', 'excel_add_shape', { host: 'wps', workbookName: WB, sheetName: '看板', kind: 'wordart', left: 600, top: 14, width: 240, height: 48, text: 'Office Agent Bridge', fontName: '微软雅黑', fontSize: 16, name: 'k_art' });
const shapes = await call('excel_list_shapes', { host: 'wps', workbookName: WB, sheetName: '看板' });
console.log(`      形状读回：共 ${shapes.data?.count} 个`);
const num1 = (shapes.data?.shapes ?? []).find(s => s.name === 'k_num1');
console.log(`      k_num1 读数：${JSON.stringify({ text: num1?.text, fill: num1?.fillColor, align: num1?.textAlign, valign: num1?.textVAlign })}`);

// 原生图表
// ⚠️ 图表数据源必须与图表**同表**：宿主拒绝带表名的跨表区域串
// （实测 dataRange/sourceAddress/dataRanges 三种写法全都报 "Parameter type error source"）
await step('插入原生柱状图（建在数据表上）', 'excel_add_chart', {
  host: 'wps', workbookName: WB, sheetName: '明细',
  dataRange: 'A1:C19', chartType: 'column_clustered', position: { leftCell: 'G2', width: 560, height: 280 }, title: '各区域月度营收'
});
const ch = await call('excel_get_charts', { host: 'wps', workbookName: WB, sheetName: '明细' });
ck('CAP-21', '图表**真的有数据**（系列数 > 0）', (ch.data?.charts?.[0]?.seriesCount ?? 0) > 0, `系列数=${ch.data?.charts?.[0]?.seriesCount} 标题=${JSON.stringify(ch.data?.charts?.[0]?.title)}`);
await step('更新图表标题（CAP-21）', 'wps_update_chart', { workbookName: WB, sheetName: '明细', chartName: ch.data?.charts?.[0]?.shapeName, title: '各区域月度营收（万元）' });
const ch2 = await call('excel_get_charts', { host: 'wps', workbookName: WB, sheetName: '明细' });
ck('CAP-21', '标题已更新且读回一致', ch2.data?.charts?.[0]?.title === '各区域月度营收（万元）', `读回=${JSON.stringify(ch2.data?.charts?.[0]?.title)}`);

console.log('\n═══ 三、说明表：超链接目录 + 命名区域 + 文档属性 ═══');
await step('新建「说明」表', 'excel_create_sheet', { host: 'wps', workbookName: WB, sheetName: '说明' });
for (const [i, [text, target]] of [['经营明细', '#明细!A1'], ['指标看板', '#看板!A1']].entries()) {
  await step(`超链接：${text}`, 'excel_manage_hyperlink', { host: 'wps', workbookName: WB, sheetName: '说明', action: 'add', address: `A${i + 2}`, url: target, displayText: text });
}
const hls = await call('excel_manage_hyperlink', { host: 'wps', workbookName: WB, sheetName: '说明', action: 'list' });
console.log(`      超链接读回：${hls.data?.count} 个 → ${JSON.stringify(hls.data?.hyperlinks?.map(h => h.text))}`);
await step('命名区域', 'excel_manage_named_range', { host: 'wps', workbookName: WB, action: 'add', name: 'RevenueData', refersTo: '明细!$C$2:$C$19' });
const nr = await call('excel_manage_named_range', { host: 'wps', workbookName: WB, action: 'list' });
const mine = (nr.data?.names ?? []).find(n => /RevenueData/.test(String(n.name)));
console.log(`      命名区域读回：${mine?.name} → ${mine?.refersTo}`);
await step('文档属性', 'excel_manage_document_properties', { host: 'wps', workbookName: WB, action: 'apply', properties: { Title: '2026 上半年经营分析', Author: 'Office Agent Bridge', Subject: '能力验证' } });
const props = await call('excel_manage_document_properties', { host: 'wps', workbookName: WB, action: 'read' });
console.log(`      属性读回：Title=${JSON.stringify(props.data?.builtin?.Title)} Author=${JSON.stringify(props.data?.builtin?.Author)}`);

console.log('\n═══ 四、导出与打印 ═══');
await step('打印设置（A4 横向 + 打印标题行）', 'excel_configure_print_layout', { host: 'wps', workbookName: WB, sheetName: '明细', printArea: 'A1:E19', printTitleRows: '1:1', orientation: 'landscape', paperSize: 'a4', centerHorizontally: true, fitToPagesWide: 1 });
const pl = await call('excel_configure_print_layout', { host: 'wps', workbookName: WB, sheetName: '明细', action: 'read' });
console.log(`      打印设置读回：方向=${pl.data?.current?.orientation} 纸张=${pl.data?.current?.paperSize} 区域=${JSON.stringify(pl.data?.current?.printArea)}`);
const pdf = await step('导出 PDF（整表）', 'excel_export_sheet_pdf', { host: 'wps', workbookName: WB, sheetName: '明细', scope: 'sheet' });
const pdfPath = pdf.data?.outputPath;
console.log(`      PDF 落盘：fileWritten=${pdf.data?.fileWritten} ${pdfPath ? (fs.existsSync(pdfPath) ? '✔ ' + fs.statSync(pdfPath).size + ' 字节 → ' + pdfPath : '✖ 未找到 ' + pdfPath) : ''}`);
const chartPng = '/Users/jolin/Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/经营分析_图表.png';
try { fs.unlinkSync(chartPng); } catch (e) {}
await step('导出图表图片', 'excel_export_chart_image', { host: 'wps', workbookName: WB, sheetName: '明细', chartIndex: 1, outputPath: chartPng, format: 'PNG' });
const ci = await call('excel_export_chart_image', { host: 'wps', workbookName: WB, sheetName: '明细', chartIndex: 1, outputPath: chartPng, format: 'PNG' });
console.log(`      图表图片落盘：fileWritten=${ci.data?.fileWritten} ${ci.data?.fileSizeBytes ?? ''} 字节`);
await step('保存工作簿', 'excel_save_workbook', { host: 'wps', workbookName: WB });

const pass = log.filter(x => x.ok).length;
console.log(`\n════════════════════════`);
console.log(`  能力验证：${pass}/${log.length} 步成功`);
if (pass !== log.length) { console.log('  失败步骤：'); log.filter(x => !x.ok).forEach(x => console.log(`    ${x.label} (${x.name})`)); }
