/**
 * 参数生效性测试  ——  node scripts/param-effect-test.mjs
 *
 * 存在的理由：此前所有验证都只断言 success === true，于是
 * 「调用成功但没生效」这一类缺陷一条都拦不住（schema 有参数宿主不读、
 * 写进去的值被丢弃、参数名不匹配、单元格说改却没改…）。
 *
 * 断言模式固定三步，缺一不可：
 *   1. 写入一个**特征值**（与默认值不同，便于识别）
 *   2. **直接问宿主**读回真实状态（绕过工具自己的返回体）
 *   3. 比对特征值是否真的落到宿主
 *
 * 只要第 3 步不过，就是「参数没生效」——不管工具返回什么。
 *
 * 注意：颜色常量用 bgr() 换算，不要手写 BGR 整数
 * （首版手写了 7 个期望值，7 个全算错，白白报了 7 个假失败）。
 */
/**
 * （原始说明）
 *
 * 为什么需要它：此前所有验证都只断言 "success === true"，
 * 于是"调用成功但没生效"的一类缺陷一条都拦不住（schema 有参数、宿主不读；
 * 写进去的值被丢弃；参数名不匹配…）。
 *
 * 本测试的断言模式固定为三步：
 *   1. 写入一个**特征值**（不与默认值相同，便于识别）
 *   2. **直接问宿主**读回真实状态（绕过工具自己的返回体）
 *   3. 比对特征值是否真的落到了宿主
 *
 * 只要第 3 步不过，就是"参数没生效"，不管工具返回什么。
 */
import fs from 'node:fs';

const token = fs.readFileSync(process.env.HOME + '/.wps-bridge/token', 'utf8').trim();
async function call(name, args = {}) {
  const r = await fetch('http://127.0.0.1:19890/api/v1/tool/call', {
    method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, arguments: args }) });
  return await r.json().catch(() => ({}));
}
/** 直接问宿主（不经过工具的返回体）——这是本测试的关键 */
async function host(code) {
  const r = await call('wps_execute_script', { component: 'excel', code });
  return r.success ? r.data?.returnValue : null;
}

/** 十六进制颜色 → 宿主读回的 BGR 整数（r + g*256 + b*65536）。
 *  手写这些常量时我算错过 4 个，所以改成函数算。 */
function bgr(hex) {
  const h = hex.replace('#', '');
  const r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
  return r + g * 256 + b * 65536;
}

const WB = '参数生效性测试.xlsx';
const SH = 'T';
const log = [];
function ck(feature, param, ok, detail) {
  log.push({ feature, param, ok: !!ok, detail: String(detail ?? '') });
  console.log(`  ${ok ? '✔' : '✖'} [${feature}] ${param}${ok ? '' : '   → ' + String(detail).slice(0, 150)}`);
}

// 用工具新建工作簿（顺带验证 create_workbook）
const mk = await call('wps_create_workbook', { savePath: `/Users/jolin/Downloads/${WB}`, sheetName: SH });
if (!mk.success) { console.log('✖ 建簿失败，中止:', String(mk.error).slice(0, 120)); process.exit(1); }
console.log('隔离工作簿就绪\n');

// ═══ 1. 单元格写入：values / formulas 都要真的落表 ═══
console.log('── 单元格 ──');
{
  const MARK = '特征值-7382';
  await call('excel_patch_cells', { host: 'wps', workbookName: WB, sheetName: SH, address: 'A1:B3',
    values: [[MARK, 11], ['x', 22], ['y', 33]] });
  const v = await host(`return String(wb.Worksheets.Item('${SH}').Range('A1').Value2);`);
  ck('patch_cells', 'values', v === MARK, `宿主 A1=${v}，期望 ${MARK}`);

  await call('excel_patch_cells', { host: 'wps', workbookName: WB, sheetName: SH, address: 'C1', formulas: [['=B1*3']] });
  const f = await host(`return String(wb.Worksheets.Item('${SH}').Range('C1').Formula);`);
  const fv = await host(`return String(wb.Worksheets.Item('${SH}').Range('C1').Value2);`);
  ck('patch_cells', 'formulas', String(f).includes('B1*3') && Number(fv) === 33, `公式=${f} 值=${fv}（期望公式含 B1*3、值 33）`);
}

// ═══ 2. 格式：每一项都读回宿主 ═══
console.log('── 格式 ──');
{
  await call('excel_format_cells', { host: 'wps', workbookName: WB, sheetName: SH, address: 'A1',
    backgroundColor: '#FF8800', fontColor: '#003366', bold: true, fontSize: 17, horizontalAlignment: 'center', numberFormat: '0.00' });
  const r = await host(`const c=wb.Worksheets.Item('${SH}').Range('A1'); const o=[]; o.push('填充='+c.Interior.Color); o.push('字色='+c.Font.Color); o.push('粗='+c.Font.Bold); o.push('号='+c.Font.Size); o.push('对齐='+c.HorizontalAlignment); o.push('格式='+c.NumberFormat); return o.join('|');`);
  const s = String(r ?? '');
  ck('format_cells', 'backgroundColor', s.includes(`填充=${bgr('#FF8800')}`), `${s}（#FF8800 → BGR ${bgr('#FF8800')}）`);
  ck('format_cells', 'fontColor', s.includes(`字色=${bgr('#003366')}`), `${s}（#003366 → BGR ${bgr('#003366')}）`);
  ck('format_cells', 'bold', /粗=-1|粗=true/i.test(s), s);
  ck('format_cells', 'fontSize', s.includes('号=17'), s);
  ck('format_cells', 'horizontalAlignment', /对齐=-4108/.test(s), s);
  ck('format_cells', 'numberFormat', s.includes('0.00'), s);
}

// ═══ 3. 条件格式：规则类型与颜色 ═══
console.log('── 条件格式 ──');
{
  await call('excel_add_conditional_formatting', { host: 'wps', workbookName: WB, sheetName: SH, address: 'B1:B3',
    ruleType: 'cell_value', operator: 'greater_than', formula1: '5', backgroundColor: '#FFDDDD' });
  const r = await host(`const fc=wb.Worksheets.Item('${SH}').Range('B1:B3').FormatConditions; const f=fc.Item(1); return fc.Count+'|'+f.Type+'|'+f.Operator+'|'+f.Formula1+'|'+f.Interior.Color;`);
  const s = String(r ?? '');
  const [cnt, type, op, formula, color] = s.split('|');
  ck('conditional_format', 'ruleType/operator', Number(cnt) >= 1 && Number(type) === 1 && Number(op) === 5, s);
  ck('conditional_format', 'formula1', String(formula).includes('5'), s);
  ck('conditional_format', 'backgroundColor', Number(color) === bgr('#FFDDDD'), `${s}（#FFDDDD → ${bgr('#FFDDDD')}）`);
}

// ═══ 4. 数据有效性 ═══
console.log('── 数据有效性 ──');
{
  await call('excel_set_data_validation', { host: 'wps', workbookName: WB, sheetName: SH, address: 'D1:D3',
    validationType: 'list', listItems: ['甲选项', '乙选项', '丙选项'] });
  const r = await host(`const v=wb.Worksheets.Item('${SH}').Range('D1').Validation; return v.Type+'|'+v.Formula1;`);
  ck('data_validation', 'listItems', String(r).includes('甲选项') && String(r).startsWith('3'), String(r));
}

// ═══ 5. 冻结窗格 ═══
console.log('── 冻结窗格 ──');
{
  await call('excel_freeze_panes', { host: 'wps', workbookName: WB, sheetName: SH, freezeRowIndex: 3, freezeColumnIndex: 2 });
  const r = await host(`const w=app.ActiveWindow; return w.FreezePanes+'|'+w.SplitRow+'|'+w.SplitColumn;`);
  ck('freeze_panes', 'freezeRowIndex/Column', String(r) === 'true|2|1', `宿主 ${r}，期望 true|2|1（冻 3 行 2 列 → Split=2,1）`);
}

// ═══ 6. 打印设置 ═══
console.log('── 打印设置 ──');
{
  await call('excel_configure_print_layout', { host: 'wps', workbookName: WB, sheetName: SH,
    printArea: 'A1:C3', orientation: 'landscape', paperSize: 'a4', centerHorizontally: true, printGridlines: true });
  const r = await host(`const p=wb.Worksheets.Item('${SH}').PageSetup; return p.PrintArea+'|'+p.Orientation+'|'+p.PaperSize+'|'+p.CenterHorizontally+'|'+p.PrintGridlines;`);
  const s = String(r ?? '');
  ck('print_layout', 'printArea/orientation/paperSize', /\$?A\$?1:\$?C\$?3/.test(s) && s.includes('|2|') && s.includes('|9|'), s);
  ck('print_layout', 'centerHorizontally/printGridlines', /true\|true|-1\|-1/.test(s), s);
}

// ═══ 7. 图表：**每个参数都读回宿主**（此前只有 success 断言，漏掉一堆）═══
console.log('── 图表（重点：参数是否真的落到宿主）──');
{
  await call('excel_patch_cells', { host: 'wps', workbookName: WB, sheetName: SH, address: 'F1:H4',
    values: [['月', '甲', '乙'], ['1月', 10, 20], ['2月', 30, 25], ['3月', 45, 40]] });
  const ch = await call('excel_add_chart', { host: 'wps', workbookName: WB, sheetName: SH,
    dataRange: 'F1:H4', chartType: 'column_clustered', position: { leftCell: 'J1', width: 480, height: 260 },
    title: '参数生效性-图表', hasLegend: true, hasDataLabels: true,
    seriesColors: ['#2563EB', '#0E9F9F', '#D99A2B'] });
  ck('add_chart', '创建成功', ch.success, String(ch.error).slice(0, 120));

  const probe = await host(`const ws=wb.Worksheets.Item('${SH}'); let c=null,sh=null; for(let i=1;i<=ws.Shapes.Count;i++){const x=ws.Shapes.Item(i); try{ if(x.HasChart){c=x.Chart;sh=x;break;} }catch(e){}} if(!c) return 'NO_CHART';
    const o=[];
    o.push('标题='+c.ChartTitle.Text);
    o.push('类型='+c.ChartType);
    o.push('图例='+c.HasLegend);
    o.push('系列数='+c.SeriesCollection().Count);
    const sc=c.SeriesCollection(); const cols=[]; for(let i=1;i<=sc.Count;i++){ let fc=null; try{fc=sc.Item(i).Format.Fill.ForeColor.RGB;}catch(e){} let dl=null; try{dl=sc.Item(i).HasDataLabels;}catch(e){} cols.push(fc+'/'+dl); }
    o.push('系列='+cols.join(','));
    o.push('位置='+Math.round(sh.Left)+','+Math.round(sh.Top)+','+Math.round(sh.Width)+','+Math.round(sh.Height));
    return o.join(' | ');`);
  const p = String(probe ?? '');
  ck('add_chart', 'dataRange（系列数=2）', /系列数=2/.test(p), p);
  ck('add_chart', 'title', p.includes('标题=参数生效性-图表'), p);
  ck('add_chart', 'chartType（51）', p.includes('类型=51'), p);
  ck('add_chart', 'hasLegend', /图例=(true|-1)/.test(p), p);
  ck('add_chart', 'hasDataLabels', /\/true/.test(p) || /\/-1/.test(p), p);
  ck('add_chart', 'seriesColors（2563EB→15426341）', p.includes('15426341'), p);
  ck('add_chart', 'position', /位置=432,0,480,260/.test(p.replace(/\s/g, '')), `${p}（leftCell J1 → Left=432, Top=0, 480x260）`);

  // update_chart 的参数
  const up = await call('wps_update_chart', { workbookName: WB, sheetName: SH, chartIndex: 1,
    title: '改后标题', fontName: '微软雅黑', legendPosition: 'right', dataLabelColorMatchesSeries: true,
    position: { topDelta: 25 } });
  ck('update_chart', '调用成功', up.success, String(up.error).slice(0, 120));
  const p2 = await host(`const ws=wb.Worksheets.Item('${SH}'); let c=null,sh=null; for(let i=1;i<=ws.Shapes.Count;i++){const x=ws.Shapes.Item(i); try{ if(x.HasChart){c=x.Chart;sh=x;break;} }catch(e){}} if(!c) return 'NO_CHART';
    const sc=c.SeriesCollection(); const o=[];
    o.push('标题='+c.ChartTitle.Text);
    o.push('标题字体='+c.ChartTitle.Font.Name);
    o.push('图例位置='+c.Legend.Position);
    const rows=[]; for(let i=1;i<=sc.Count;i++){ let fc=null,lc=null; try{fc=sc.Item(i).Format.Fill.ForeColor.RGB;}catch(e){} try{lc=sc.Item(i).DataLabels().Font.Color;}catch(e){} rows.push(fc===lc); }
    o.push('标签色一致='+rows.join(','));
    o.push('Top='+Math.round(sh.Top));
    o.push('刻度字体='+c.Axes().Item(1).TickLabels.Font.Name);
    return o.join(' | ');`);
  const q = String(p2 ?? '');
  ck('update_chart', 'title', q.includes('标题=改后标题'), q);
  ck('update_chart', 'fontName', q.includes('标题字体=微软雅黑') && q.includes('刻度字体=微软雅黑'), q);
  ck('update_chart', 'legendPosition', q.includes('图例位置=-4152'), q);
  ck('update_chart', 'dataLabelColorMatchesSeries', q.includes('标签色一致=true,true'), q);
  ck('update_chart', 'position.topDelta', /Top=25/.test(q), `${q}（原 Top=0，topDelta:25 → 25）`);
}

// ═══ 8. 形状：文字样式参数 ═══
console.log('── 矢量形状 ──');
{
  await call('excel_add_shape', { host: 'wps', workbookName: WB, sheetName: SH, kind: 'geometric', shapeType: 'rounded_rectangle',
    left: 20, top: 300, width: 200, height: 80, fillColor: '#1F3864', text: '形状参数', fontSize: 18, bold: true, textAlign: 'center', textVAlign: 'middle', name: 'ftShape' });
  const r = await host(`const s=wb.Worksheets.Item('${SH}').Shapes.Item('ftShape'); const tf=s.TextFrame; const o=[];
    o.push('填充='+s.Fill.ForeColor.RGB);
    o.push('文字='+tf.Characters().Text);
    o.push('字号='+tf.Characters().Font.Size);
    o.push('粗='+tf.Characters().Font.Bold);
    o.push('字体='+tf.Characters().Font.Name);
    o.push('H='+tf.HorizontalAlignment);
    o.push('V='+tf.VerticalAlignment);
    return o.join('|');`);
  const s = String(r ?? '');
  ck('add_shape', 'fillColor', s.includes(`填充=${bgr('#1F3864')}`), `${s}（#1F3864 → BGR ${bgr('#1F3864')}）`);
  ck('add_shape', 'text', s.includes('文字=形状参数'), s);
  ck('add_shape', 'fontSize', s.includes('字号=18'), s);
  ck('add_shape', 'bold', /粗=-1|粗=true/.test(s), s);
  ck('add_shape', '默认字体（微软雅黑）', s.includes('字体=微软雅黑'), s);
  ck('add_shape', 'textAlign（-4108 居中）', /H=-4108/.test(s), s);
  ck('add_shape', 'textVAlign（-4108 居中）', /V=-4108/.test(s), s);
}

// ═══ 汇总（清理后输出）═══
await call('wps_execute_script', { component: 'excel', code: `for(let i=app.Workbooks.Count;i>=1;i--){const w=app.Workbooks.Item(i); if(w.Name.indexOf('参数生效性测试')>=0) w.Close(false);} return 1;` });
try { fs.unlinkSync(`/Users/jolin/Downloads/${WB}`); } catch (e) {}

const pass = log.filter(x => x.ok).length;
console.log(`\n════════════════════════`);
console.log(`  参数生效性：${pass}/${log.length} 通过`);
const byF = {};
for (const x of log) (byF[x.feature] ||= []).push(x);
for (const [f, arr] of Object.entries(byF)) {
  const p = arr.filter(x => x.ok).length;
  console.log(`    ${f.padEnd(22)} ${p}/${arr.length}${p === arr.length ? ' ✔' : ' ✖ ' + arr.filter(x => !x.ok).map(x => x.param).join(',')}`);
}
if (pass !== log.length) { console.log('\n  未生效的参数：'); log.filter(x => !x.ok).forEach(x => console.log(`    [${x.feature}] ${x.param}`)); }
process.exit(pass === log.length ? 0 : 1);
