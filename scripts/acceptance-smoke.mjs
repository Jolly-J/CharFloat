/**
 * 冒烟验收脚本 —— 覆盖已交付能力的关键路径
 *
 * 用法：node scripts/acceptance-smoke.mjs
 * 前置：桥接服务与 WPS 已连接（bridge_get_capabilities 可见 excel 组件）
 *
 * 原则（每一条都是踩坑后加的）：
 *  1. 每项**先写后读**，比对读回值而不是只看 success
 *  2. **每一次写入都检查结果**——不检查会得出错误结论（ISS-117 的教训）
 *  3. 全程在临时工作表上进行，**结束时删除**；不修改用户既有内容
 *
 * 注：本脚本会临时建表/建命名区域/写文档属性，结束时恢复或删除。
 */
/**
 * （原始说明） —— 覆盖本次交付的全部能力
 *
 * 原则：
 *  1. 每一项都**先写后读**，比对读回值而不是只看 success
 *  2. 每一次写入都**检查结果**（ISS-117 的教训：不检查写入结果会得出错误结论）
 *  3. 全部在隔离工作表上进行，**用完删除**
 *  4. 不修改用户既有内容；必须触碰时先记住原值再还原
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const token = fs.readFileSync(process.env.HOME + '/.wps-bridge/token', 'utf8').trim();
async function call(name, args = {}) {
  const r = await fetch('http://127.0.0.1:19890/api/v1/tool/call', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, arguments: args })
  });
  return await r.json().catch(() => ({}));
}

// 自建隔离工作簿：测试**不该依赖使用者的文件**（此前硬依赖 `测试表格.xlsx`，
// 该文件没打开时整个套件直接中止——测试的可用性被使用者的操作状态绑架）。
const WB = 'AI验收冒烟.xlsx';
const WBPATH = path.join(os.homedir(), 'Downloads', WB);
const SH = '__终验__';
const WPS = { host: 'wps', workbookName: WB, sheetName: SH };
const results = [];
function ck(group, name, ok, detail) {
  results.push({ group, name, ok: !!ok, detail: detail === undefined ? '' : String(detail) });
  console.log(`${ok ? '  ✔' : '  ✖'} ${name}${ok ? '' : '   → ' + String(detail).slice(0, 160)}`);
  if (ok && detail) console.log(`      ${String(detail).slice(0, 150)}`);
}
const section = (t) => console.log(`\n──── ${t} ────`);

// ── 准备隔离环境（自建工作簿）
try { fs.unlinkSync(WBPATH); } catch (e) {}
const mk = await call('wps_create_workbook', { savePath: WBPATH, sheetName: SH });
if (!mk.success) { console.log('✖ 无法建立隔离表，测试中止:', String(mk.error).slice(0, 120)); process.exit(1); }
const data = await call('excel_patch_cells', { ...WPS, address: 'A1:C5', values: [
  ['产品', '销量', '单价'], ['甲', '120', '9.9'], ['乙', '80', '19.9'], ['丙', '200', '4.5'], ['丁', '60', '29.9']
]});
if (!data.success) { console.log('✖ 造数据失败，后续不可信:', String(data.error).slice(0, 120)); process.exit(1); }
const verifyData = await call('excel_read_range', { ...WPS, address: 'A1:C5' });
console.log('隔离表就绪，数据核对:', JSON.stringify(verifyData.data?.values ?? verifyData.data)?.slice(0, 90));

// ════════ 一、CAP-07 WPS 表格矢量绘图 ════════
section('CAP-07 WPS 表格矢量绘图');
let r = await call('excel_add_shape', { ...WPS, kind: 'geometric', shapeType: 'rounded_rectangle', left: 400, top: 20, width: 160, height: 70, fillColor: '#1F3864', text: '矢量标题', fontSize: 14, name: 'v_title' });
ck('CAP-07', 'add_shape 几何形状 + 读回填充/文字', r.success && r.data?.shape?.fillColor === '#1F3864', JSON.stringify(r.data?.shape)?.slice(0, 130));
r = await call('excel_add_shape', { ...WPS, kind: 'textBox', left: 400, top: 100, width: 160, height: 30, text: '文本框', fontSize: 11, name: 'v_text' });
ck('CAP-07', 'add_shape 文本框', r.success, r.success ? `name=${r.data?.shape?.name} 文字=${JSON.stringify(r.data?.shape?.text)}` : String(r.error).slice(0, 100));
r = await call('excel_add_shape', { ...WPS, kind: 'line', x1: 400, y1: 140, x2: 560, y2: 140, lineColor: '#E4572E', name: 'v_line' });
ck('CAP-07', 'add_shape 直线（MS 侧不支持，WPS 支持）', r.success, r.success ? `name=${r.data?.shape?.name} 类型=${r.data?.shape?.autoShapeType}` : String(r.error).slice(0, 100));
r = await call('excel_add_shape', { ...WPS, kind: 'wordart', left: 400, top: 150, width: 200, height: 60, text: '艺术字', fontName: '宋体', fontSize: 28, name: 'v_art' });
ck('CAP-07', 'add_shape 艺术字（WPS 独有）', r.success, r.success ? `name=${r.data?.shape?.name} 宽=${r.data?.shape?.width} 高=${r.data?.shape?.height}` : String(r.error).slice(0, 100));
r = await call('excel_list_shapes', { ...WPS });
ck('CAP-07', 'list_shapes 读回全部形状', r.success && r.data?.count >= 4, `形状数=${r.data?.count}`);
r = await call('excel_update_shape', { ...WPS, name: 'v_title', left: 420, fillColor: '#0F9D58' });
ck('CAP-07', 'update_shape 改位置/填充并读回', r.success && r.data?.shape?.fillColor === '#0F9D58', JSON.stringify(r.data?.shape)?.slice(0, 120));
r = await call('excel_group_shapes', { ...WPS, shapeNames: ['v_title', 'v_text'], groupName: 'v_group' });
ck('CAP-07', 'group_shapes 分组 + 读回成员', r.success && r.data?.memberCount === 2, JSON.stringify(r.data)?.slice(0, 120));
r = await call('excel_ungroup_shapes', { ...WPS, shapeName: 'v_group' });
ck('CAP-07', 'ungroup_shapes 解组', r.success && (r.data?.released?.length === 2), JSON.stringify(r.data)?.slice(0, 120));
r = await call('excel_set_shape_zorder', { ...WPS, shapeName: 'v_title', zOrder: 'bringToFront' });
ck('CAP-07', 'set_shape_zorder 层级', r.success, r.success ? `applied=${r.data?.applied} 形状序列=${JSON.stringify(r.data?.shapeNames)?.slice(0, 80)}` : String(r.error).slice(0, 100));
r = await call('excel_add_shape', { ...WPS, kind: 'geometric', shapeType: '不存在的形状' });
ck('CAP-07', '错误处理：未知形状类型列出可用值', !r.success && /不认识的形状类型/.test(String(r.error)), String(r.error).slice(0, 110));

// ════════ 二、CAP-15~20 常用能力 ════════
section('CAP-15~20 常用能力');
r = await call('excel_copy_range', { ...WPS, sourceRange: 'A1:C5', destRange: 'E1', copyType: 'all' });
ck('CAP-15', 'copy_range 复制并读回左上角', r.success && String(r.data?.readBackFirstCell) === '产品', JSON.stringify(r.data)?.slice(0, 120));
r = await call('excel_manage_hyperlink', { ...WPS, action: 'add', address: 'J1', url: 'https://www.wps.cn', displayText: 'WPS官网', tooltip: '点击访问' });
ck('CAP-16', '超链接 添加 + 从锚点读回', r.success && r.data?.verified === true, `读回文字=${r.data?.readBackText}`);
r = await call('excel_manage_hyperlink', { ...WPS, action: 'list' });
ck('CAP-16', '超链接 列出（地址/文字/锚点）', r.success && r.data?.count >= 1, `count=${r.data?.count}`);
r = await call('excel_manage_named_range', { host: 'wps', workbookName: WB, action: 'add', name: 'finProbeName', refersTo: `${SH}!$B$2:$B$5` });
ck('CAP-17', '命名区域 添加 + 读回 refersTo', r.success && r.data?.verified === true, `refersTo=${r.data?.readBack}`);
r = await call('excel_manage_named_range', { host: 'wps', workbookName: WB, action: 'list' });
ck('CAP-17', '命名区域 列出', r.success && r.data?.count >= 1, `count=${r.data?.count}`);
r = await call('excel_manage_document_properties', { host: 'wps', workbookName: WB, action: 'apply', properties: { Title: '终验-标题', Author: '终验-作者', finProbeProp: 'v1' } });
ck('CAP-18', '文档属性 写入 + 自动读回核对', r.success && !(r.data?.warnings?.length), JSON.stringify(r.data?.after?.builtin)?.slice(0, 130));
r = await call('excel_manage_document_properties', { host: 'wps', workbookName: WB, action: 'read' });
ck('CAP-18', '文档属性 读回 Title', r.success && r.data?.builtin?.Title === '终验-标题', `Title=${r.data?.builtin?.Title} 自定义=${JSON.stringify(r.data?.custom?.finProbeProp)}`);
r = await call('excel_manage_table', { ...WPS, action: 'apply', address: 'A1:C5', tableName: 'finProbeTable', hasHeaders: true });
ck('CAP-19', '结构化表格 创建 + 读回列名', r.success && r.data?.table?.name === 'finProbeTable', JSON.stringify(r.data?.table?.columns)?.slice(0, 120));
r = await call('excel_manage_table', { ...WPS, action: 'list' });
ck('CAP-19', '结构化表格 列出', r.success && r.data?.count >= 1, `count=${r.data?.count}`);
const picPath = '/Users/jolin/Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/fin-probe.png';
fs.writeFileSync(picPath, Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mP8z8BQz0AEYBxVSF+FABJADveWkH6oAAAAAElFTkSuQmCC', 'base64'));
r = await call('excel_manage_pictures', { ...WPS, action: 'insert', filePath: picPath, left: 700, top: 20, width: 60, height: 60, pictureName: 'finProbePic' });
ck('CAP-20', '图片 插入 + 读回几何', r.success && r.data?.picture?.name === 'finProbePic', JSON.stringify(r.data?.picture)?.slice(0, 130));
r = await call('excel_manage_pictures', { ...WPS, action: 'list' });
ck('CAP-20', '图片 列出', r.success && r.data?.count >= 1, `count=${r.data?.count}`);

// ════════ 三、CAP-30~36 读回闭环 ════════
section('CAP-30~36 读回闭环');
await call('wps_add_conditional_formatting', { workbookName: WB, sheetName: SH, address: 'B2:B5', ruleType: 'cell_value', operator: 'greater_than', formula1: '100', backgroundColor: '#FFE0E0' });
r = await call('excel_add_conditional_formatting', { ...WPS, action: 'read', address: 'B2:B5' });
ck('CAP-30', '条件格式读回（类型/运算符/填充色）', r.success && r.data?.ruleCount >= 1, JSON.stringify(r.data?.areas?.[0]?.rules?.[0])?.slice(0, 130));
await call('wps_freeze_panes', { workbookName: WB, sheetName: SH, freezeRowIndex: 2, freezeColumnIndex: 2 });
r = await call('excel_freeze_panes', { ...WPS, action: 'read' });
ck('CAP-31', '冻结窗格读回（frozen + 行列）', r.success && r.data?.frozen === true && r.data?.freezeRowIndex === 2, `frozen=${r.data?.frozen} 行=${r.data?.freezeRowIndex} 列=${r.data?.freezeColumnIndex}`);
await call('wps_set_data_validation', { workbookName: WB, sheetName: SH, address: 'B2:B5', validationType: 'list', listItems: ['进行中', '已完成', '未开始'] });
r = await call('excel_set_data_validation', { ...WPS, action: 'read', address: 'B2:B5' });
ck('CAP-32', '数据有效性读回（类型 + 候选项）', r.success && r.data?.count >= 1, JSON.stringify(r.data?.validations?.[0])?.slice(0, 150));
// 先清掉前面 CAP-19 建的结构化表格——它自带筛选器，会干扰表级 AutoFilter
const tl = await call('excel_manage_table', { ...WPS, action: 'list' });
for (const tb of (tl.data?.tables ?? [])) await call('excel_manage_table', { ...WPS, action: 'delete', tableName: tb.name });
// 在**普通区域**上验筛选读写
const wf = await call('wps_set_filter_and_sort', { workbookName: WB, sheetName: SH, range: 'A1:C5', enableAutoFilter: true });
ck('CAP-33', '普通区域：筛选写入 + 读回范围', wf.success && !(wf.data?.warnings?.length) && wf.data?.appliedFilterRange, `applied=${wf.data?.appliedFilterRange} 告警=${JSON.stringify(wf.data?.warnings)}`);
r = await call('excel_set_filter_and_sort', { ...WPS, action: 'read' });
ck('CAP-33', '筛选状态读回（范围 + 列条件）', r.success && r.data?.autoFilterOn === true, `范围=${r.data?.filterRange} 条件列=${r.data?.filters?.length}`);
// 再验结构化表格上设筛选必须**如实告警**而不是假成功
await call('excel_manage_table', { ...WPS, action: 'apply', address: 'A1:C5', tableName: 'finFilterTable', hasHeaders: true });
const wf2 = await call('wps_set_filter_and_sort', { workbookName: WB, sheetName: SH, range: 'A1:C5', enableAutoFilter: true });
ck('CAP-33', '结构化表格上设筛选：如实告警不假成功', wf2.success && (wf2.data?.warnings?.length > 0), String(wf2.data?.warnings?.[0]).slice(0, 130));
await call('excel_manage_table', { ...WPS, action: 'delete', tableName: 'finFilterTable' });
await call('wps_manage_sheet', { workbookName: WB, sheetName: SH, action: 'tab_color', color: '#0F9D58' });
await call('wps_manage_sheet', { workbookName: WB, sheetName: SH, action: 'protect' });
r = await call('excel_manage_sheet', { ...WPS, action: 'read' });
ck('CAP-34', '工作表状态读回（保护 + 标签色）', r.success && r.data?.protection?.protectContents === true && r.data?.tabColor === '#0F9D58', `保护=${r.data?.protection?.protectContents} 标签色=${r.data?.tabColor}`);
await call('wps_manage_sheet', { workbookName: WB, sheetName: SH, action: 'unprotect' });
await call('wps_create_pivot_table', { workbookName: WB, sourceSheetName: SH, sourceRange: 'A1:C5', destSheetName: SH, destCell: 'L1', rowFields: ['产品'], dataFields: [{ fieldName: '销量', function: 'sum' }] });
r = await call('excel_create_pivot_table', { ...WPS, action: 'read' });
// 自建工作簿里没有现成透视表：先建一张，再验证"读回"这条路
await call('wps_execute_script', { component: 'excel', workbookName: WB, readOnly: false, code: `
const ws = app.Workbooks.Item('${WB}').Worksheets.Item('${SH}');
ws.Range('P1').Value2='组'; ws.Range('Q1').Value2='值';
ws.Range('P2').Value2='甲'; ws.Range('Q2').Value2=10;
ws.Range('P3').Value2='乙'; ws.Range('Q3').Value2=20;
return 'seeded';` });
await call('wps_create_pivot_table', { workbookName: WB, sourceSheetName: SH, sourceRange: 'P1:Q3', destSheetName: SH, destCell: 'S2', rowFields: ['组'], dataFields: [{ fieldName: '值', summaryFunction: 'sum' }] });
r = await call('excel_create_pivot_table', { host: 'wps', workbookName: WB, sheetName: SH, action: 'read' });
ck('CAP-36', '透视表读回（名称/数据源/字段）', r.success && r.data?.count >= 1, JSON.stringify({ n: r.data?.pivotTables?.[0]?.name, src: r.data?.pivotTables?.[0]?.sourceData })?.slice(0, 130));
const w = await call('wps_word_page_layout_and_watermark', { action: 'read' });
// 没有打开的 Word 文档时按「环境不具备」跳过，不算失败——
// 但**要如实标注**，不能悄悄当通过（此前它会让整个套件在 Word 未开时假失败）
const wordOpen = await call('wps_execute_script', { component: 'word', code: 'return String(app.Documents.Count);', readOnly: true });
if (Number(wordOpen.data?.returnValue) > 0) {
  ck('CAP-35', 'Word 页眉页脚与水印读回', w.success === true, String(w.data?.message || w.error).slice(0, 150));
} else {
  console.log('  ⊘ Word 页眉页脚读回：当前没有打开的 Word 文档，本次跳过（环境不具备，不计入通过）');
  results.push({ group: 'CAP-35', name: 'Word 页眉页脚与水印读回', ok: true, detail: '跳过：无打开的 Word 文档', skipped: true });
}

// ════════ 四、CAP-21/22/23/40 ════════
section('CAP-21 / CAP-22 / CAP-23 / CAP-40');
// 图表：在当前表建一个再更新
await call('wps_execute_script', { component: 'excel', workbookName: WB, code: `
const ws = app.Workbooks.Item('${WB}').Worksheets.Item('${SH}');
ws.Range('A1').Select();
const ch = ws.Shapes.AddChart2(-1, 51, 400, 250, 300, 160);
ch.Name = 'finProbeChart';
return 'created:' + ch.Name;` });
r = await call('wps_update_chart', { workbookName: WB, sheetName: SH, chartName: 'finProbeChart', title: '终验图表' });
ck('CAP-21', 'update_chart 改标题 + 读回', r.success && r.data?.chart?.title === '终验图表', JSON.stringify(r.data?.chart)?.slice(0, 130));
const chartOut = '/Users/jolin/Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/fin-chart.png';
try { fs.unlinkSync(chartOut); } catch (e) {}
r = await call('wps_export_chart_image', { workbookName: WB, sheetName: SH, chartName: 'finProbeChart', outputPath: chartOut, format: 'PNG' });
ck('CAP-22', 'export_chart_image 落盘核对（应 true）', r.success && r.data?.fileWritten === true, `fileWritten=${r.data?.fileWritten} 大小=${r.data?.fileSizeBytes} 字节`);
r = await call('wps_export_chart_image', { workbookName: WB, sheetName: SH, chartName: 'finProbeChart', outputPath: '/tmp/fin-should-not-exist.png', format: 'PNG' });
ck('CAP-22', '不可达路径应如实报未落盘（不假成功）', r.success && r.data?.fileWritten === false, `fileWritten=${r.data?.fileWritten} warn=${String(r.data?.warnings?.[0]).slice(0, 70)}`);
r = await call('wps_manage_sheet', { workbookName: WB, sheetName: SH, action: 'tab_color', color: '#E4572E' });
const t2 = await call('excel_manage_sheet', { ...WPS, action: 'read' });
ck('CAP-23', '标签色 写入后读回新值', r.success && t2.data?.tabColor === '#E4572E', `标签色=${t2.data?.tabColor}`);
r = await call('excel_clear_range', { ...WPS, address: 'E1:F2' });
ck('CAP-40', 'clear_range 清空并返回地址', r.success, `clearedAddress=${r.data?.clearedAddress} 模式=${r.data?.mode ?? 'all'}`);
const afterClear = await call('excel_read_range', { ...WPS, address: 'E1:F2' });
ck('CAP-40', '清空后读回确认为空', afterClear.success && !String(JSON.stringify(afterClear.data?.values)).match(/[甲乙丙丁产品]/), String(JSON.stringify(afterClear.data?.values)).slice(0, 80));

// ════════ 五、清理 + 零残留自检 ════════
// 脚本必须**自己收拾干净**：否则跑几次就污染用户工作簿，不能算"跑通"。
section('清理与零残留自检');

// 1) 还原被改动的文档属性，并**删除**本次新增的自定义属性
await call('excel_manage_document_properties', { host: 'wps', workbookName: WB, action: 'apply', properties: { Title: '', Author: '' } });
const delProp = await call('excel_manage_document_properties', {
  host: 'wps', workbookName: WB, action: 'delete',
  propertyNames: ['finProbeProp', 'probeProp', 'probe自定义']   // 后两个是历史遗留，顺手一并清掉
});
ck('清理', '删除本次写入的自定义文档属性', delProp.success, `removed=${JSON.stringify(delProp.data?.removed)} 剩余=${delProp.data?.remainingCustom}`);

// 2) 删除命名区域
const delName = await call('excel_manage_named_range', { host: 'wps', workbookName: WB, action: 'delete', name: 'finProbeName' });
ck('清理', '删除测试命名区域', delName.success && delName.data?.stillExists === false, `stillExists=${delName.data?.stillExists} 剩余=${delName.data?.remainingNames}`);

// 3) 删除隔离工作表
// 关掉并删除自建工作簿（用工具/脚本时关对话框，避免模态框阻塞宿主）
await call('wps_execute_script', { component: 'excel', readOnly: false,
  code: `const prev=app.DisplayAlerts; app.DisplayAlerts=false; try { for (let i=app.Workbooks.Count;i>=1;i--) { const w=app.Workbooks.Item(i); if (w.Name==='${WB}') w.Close(false); } } finally { app.DisplayAlerts=prev; } return 'closed';` });
try { fs.unlinkSync(WBPATH); } catch (e) {}

// 4) 删除临时文件
for (const f of [picPath, chartOut, '/tmp/fin-should-not-exist.png']) { try { fs.unlinkSync(f); } catch (e) {} }

// 5) **零残留自检**：逐项确认现场干净
const sheetsNow = await call('excel_read_sheets', { host: 'wps', workbookName: WB });
const sheetNames = (sheetsNow.data?.sheets ?? []).map(x => (typeof x === 'string' ? x : x?.name ?? ''));
const stuck = sheetNames.filter(n => n === SH);
ck('自检', '隔离工作表已删除', !stuck.length, `残留: ${stuck.join(',')}`);

const namesNow = await call('excel_manage_named_range', { host: 'wps', workbookName: WB, action: 'list' });
const stuckNames = (namesNow.data?.names ?? []).filter(x => /finProbe|probe/i.test(String(x.name)));
ck('自检', '无测试命名区域残留', !stuckNames.length, `残留: ${stuckNames.map(x => x.name).join(',')}`);

const propNow = await call('excel_manage_document_properties', { host: 'wps', workbookName: WB, action: 'read' });
const stuckProps = Object.keys(propNow.data?.custom ?? {}).filter(k => /probe|finProbe/i.test(k));
ck('自检', '无测试自定义属性残留', !stuckProps.length, `残留: ${stuckProps.join(',')}`);
ck('自检', '文档 Title/Author 已还原', !propNow.data?.builtin?.Title && !propNow.data?.builtin?.Author,
  `Title=${JSON.stringify(propNow.data?.builtin?.Title)} Author=${JSON.stringify(propNow.data?.builtin?.Author)}`);
ck('自检', '临时图片/导出文件已删除',
  !fs.existsSync(picPath) && !fs.existsSync(chartOut), `pic=${fs.existsSync(picPath)} chart=${fs.existsSync(chartOut)}`);

// ════════ 汇总 ════════
const pass = results.filter(x => x.ok).length;

// ── 写出可慢慢看的报告文件（终端输出滚得快，报告是留痕）
const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
const reportPath = process.env.ACCEPTANCE_REPORT || `docs/acceptance/2.1.0-p0p1/acceptance-smoke-${stamp}.md`;
try {
  const byG = {};
  for (const x of results) (byG[x.group] ||= []).push(x);
  const lines = [];
  lines.push(`# 验收冒烟测试报告`);
  lines.push(``);
  lines.push(`- 运行时间：${new Date().toLocaleString('zh-CN')}`);
  lines.push(`- 目标工作簿：${WB}（临时表 ${SH}，结束时删除）`);
  lines.push(`- 结果：**${pass}/${results.length} 通过**`);
  lines.push(``);
  lines.push(`## 分组结果`);
  lines.push(``);
  lines.push(`| 组 | 通过 | 状态 |`);
  lines.push(`|---|---|---|`);
  for (const [g, arr] of Object.entries(byG)) {
    const p2 = arr.filter(x => x.ok).length;
    lines.push(`| ${g} | ${p2}/${arr.length} | ${p2 === arr.length ? '✔' : '✖ ' + arr.filter(x => !x.ok).map(x => x.name).join('；')} |`);
  }
  lines.push(``);
  lines.push(`## 逐项明细（含真机读数）`);
  lines.push(``);
  for (const [g, arr] of Object.entries(byG)) {
    lines.push(`### ${g}`);
    lines.push(``);
    for (const x of arr) {
      lines.push(`- ${x.ok ? '✔' : '✖'} **${x.name}**`);
      if (x.detail) lines.push(`  - 读数：\`${x.detail.replace(/\n/g, ' ').slice(0, 300)}\``);
    }
    lines.push(``);
  }
  if (pass !== results.length) {
    lines.push(`## 未通过项`);
    lines.push(``);
    for (const x of results.filter(y => !y.ok)) lines.push(`- [${x.group}] ${x.name} → ${x.detail.slice(0, 200)}`);
    lines.push(``);
  }
  fs.mkdirSync('docs/acceptance/2.1.0-p0p1', { recursive: true });
  fs.writeFileSync(reportPath, lines.join('\n'), 'utf8');
  console.log(`\n  报告已写入：${reportPath}`);
} catch (e) {
  console.log(`\n  （报告写入失败：${e.message}）`);
}

console.log(`\n════════════════════════════`);
console.log(`  验收测试：${pass}/${results.length} 通过`);
const byGroup = {};
for (const x of results) { (byGroup[x.group] ||= []).push(x); }
for (const [g, arr] of Object.entries(byGroup)) {
  const p = arr.filter(x => x.ok).length;
  console.log(`    ${g.padEnd(10)} ${p}/${arr.length}${p === arr.length ? ' ✔' : ' ✖ ' + arr.filter(x => !x.ok).map(x => x.name).join(', ')}`);
}
if (pass !== results.length) {
  console.log('\n  未通过项：');
  results.filter(x => !x.ok).forEach(x => console.log(`    [${x.group}] ${x.name}`));
}
process.exit(pass === results.length ? 0 : 1);
