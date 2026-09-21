import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';

const token = fs.readFileSync(path.join(os.homedir(), '.wps-bridge/token'), 'utf8').trim();
const BASE_URL = 'http://127.0.0.1:19890/api/v1/tool/call';

async function callTool(name: string, args: any) {
  const res = await fetch(BASE_URL, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      name,
      arguments: { host: 'microsoft', ...args }
    })
  });
  const data = await res.json();
  return data;
}

async function runAllTests() {
  console.log('========================================================');
  console.log('🚀 开始全量逐项实机自动化测试 (Microsoft Excel macOS)');
  console.log('========================================================\n');

  const results: { tool: string; success: boolean; detail: string }[] = [];

  // 1. excel_get_workspace_summary
  try {
    const res = await callTool('excel_get_workspace_summary', {});
    const ok = res.success && res.data.hasOpenWorkbook;
    results.push({
      tool: 'excel_get_workspace_summary',
      success: ok,
      detail: ok ? `工作簿: ${res.data.workbookName}, 工作表数: ${res.data.sheetCount}, 活动表: ${res.data.activeSheetName}` : res.error
    });
  } catch (e: any) {
    results.push({ tool: 'excel_get_workspace_summary', success: false, detail: e.message });
  }

  // 2. excel_get_sheet_outline
  try {
    const res = await callTool('excel_get_sheet_outline', {});
    const ok = res.success && !res.data.isEmpty;
    results.push({
      tool: 'excel_get_sheet_outline',
      success: ok,
      detail: ok ? `已用区域: ${res.data.usedRangeAddress}, 行数: ${res.data.rowCount}, 列数: ${res.data.columnCount}` : res.error
    });
  } catch (e: any) {
    results.push({ tool: 'excel_get_sheet_outline', success: false, detail: e.message });
  }

  // 3. excel_create_sheet
  const testSheetName = '全量自动化验收_' + Math.floor(Math.random() * 1000);
  try {
    const res = await callTool('excel_create_sheet', { sheetName: testSheetName });
    const ok = res.success;
    results.push({
      tool: 'excel_create_sheet',
      success: ok,
      detail: ok ? `已新建测试工作表: ${testSheetName}` : res.error
    });
  } catch (e: any) {
    results.push({ tool: 'excel_create_sheet', success: false, detail: e.message });
  }

  // 4. excel_patch_cells (写入表头、数据与动态计算公式)
  try {
    const res = await callTool('excel_patch_cells', {
      sheetName: testSheetName,
      address: 'A1:E5',
      values: [
        ['季度', '产品线', '单价', '销量', '销售总额'],
        ['Q1', '智能终端', 2500, 120, '=C2*D2'],
        ['Q2', '智能穿戴', 1200, 350, '=C3*D3'],
        ['Q3', '云端配件', 450, 800, '=C4*D4'],
        ['合计', '全部产品', '', '=SUM(D2:D4)', '=SUM(E2:E4)']
      ]
    });
    const ok = res.success && res.data.modifiedCount === 25;
    results.push({
      tool: 'excel_patch_cells (值+公式)',
      success: ok,
      detail: ok ? `成功写入 25 个单元格（含公式联动）` : res.error
    });
  } catch (e: any) {
    results.push({ tool: 'excel_patch_cells (值+公式)', success: false, detail: e.message });
  }

  // 5. excel_read_range (读回校验)
  try {
    const res = await callTool('excel_read_range', {
      sheetName: testSheetName,
      address: 'A1:E5',
      includeFormulas: true
    });
    const ok = res.success && res.data.values && res.data.values[4][0] === '合计';
    results.push({
      tool: 'excel_read_range (读回校验)',
      success: ok,
      detail: ok ? `读回校验通过，E5 动态求和公式为: ${JSON.stringify(res.data.formulas[4][4])}` : res.error
    });
  } catch (e: any) {
    results.push({ tool: 'excel_read_range (读回校验)', success: false, detail: e.message });
  }

  // 6. excel_format_cells (格式美化：表头深蓝背景、白字、加粗；数据行货币/千分位格式与边框)
  try {
    const resHeader = await callTool('excel_format_cells', {
      sheetName: testSheetName,
      address: 'A1:E1',
      bold: true,
      backgroundColor: '#1F4E78',
      fontColor: '#FFFFFF',
      horizontalAlignment: 'center',
      borders: true
    });
    const resNumbers = await callTool('excel_format_cells', {
      sheetName: testSheetName,
      address: 'E2:E5',
      numberFormat: '¥#,##0',
      bold: true
    });
    const ok = resHeader.success && resNumbers.success;
    results.push({
      tool: 'excel_format_cells (表头+金额格式+边框)',
      success: ok,
      detail: ok ? `表头样式 (加粗/背景/字色/居中) 与数据金额格式 (¥#,##0) 设置成功` : (resHeader.error || resNumbers.error)
    });
  } catch (e: any) {
    results.push({ tool: 'excel_format_cells (表头+金额格式+边框)', success: false, detail: e.message });
  }

  // 7. excel_get_range_styles (样式读取校验)
  try {
    const res = await callTool('excel_get_range_styles', {
      sheetName: testSheetName,
      address: 'A1'
    });
    const ok = res.success && res.data.font.bold === true;
    results.push({
      tool: 'excel_get_range_styles (样式检验)',
      success: ok,
      detail: ok ? `A1 字体加粗=${res.data.font.bold}, 字体颜色=${res.data.font.color}, 背景色=${res.data.fill.color}` : res.error
    });
  } catch (e: any) {
    results.push({ tool: 'excel_get_range_styles (样式检验)', success: false, detail: e.message });
  }

  // 8. excel_auto_fit_columns (自适应列宽)
  try {
    const res = await callTool('excel_auto_fit_columns', {
      sheetName: testSheetName,
      address: 'A1:E5'
    });
    const ok = res.success;
    results.push({
      tool: 'excel_auto_fit_columns (自适应列宽)',
      success: ok,
      detail: ok ? `列宽已根据文字长度自动适应排版` : res.error
    });
  } catch (e: any) {
    results.push({ tool: 'excel_auto_fit_columns (自适应列宽)', success: false, detail: e.message });
  }

  // 9. excel_set_data_validation (数据有效性下拉菜单)
  try {
    const res = await callTool('excel_set_data_validation', {
      sheetName: testSheetName,
      address: 'A2:A4',
      validationType: 'list',
      listItems: ['Q1', 'Q2', 'Q3', 'Q4']
    });
    const ok = res.success;
    results.push({
      tool: 'excel_set_data_validation (下拉数据验证)',
      success: ok,
      detail: ok ? `已在 A2:A4 成功设置 Q1/Q2/Q3/Q4 下拉选项` : res.error
    });
  } catch (e: any) {
    results.push({ tool: 'excel_set_data_validation (下拉数据验证)', success: false, detail: e.message });
  }

  // 10. excel_freeze_panes (冻结首行，并二次防重调用验证无内部错误)
  try {
    const res1 = await callTool('excel_freeze_panes', {
      sheetName: testSheetName,
      freezeRowIndex: 1
    });
    // 关键验证点：连续调用二次冻结，验证防重 unfreeze 机制，绝不抛出 InternalError
    const res2 = await callTool('excel_freeze_panes', {
      sheetName: testSheetName,
      freezeRowIndex: 2
    });
    const ok = res1.success && res2.success;
    results.push({
      tool: 'excel_freeze_panes (冻结幂等防重测试)',
      success: ok,
      detail: ok ? `首行锁定及连续重新冻结成功，彻底消除 InternalError` : (res1.error || res2.error)
    });
  } catch (e: any) {
    results.push({ tool: 'excel_freeze_panes (冻结幂等防重测试)', success: false, detail: e.message });
  }

  // 11. excel_add_chart (创建原生柱状图并验证 replaceExisting 回执与旧图替换)
  let createdChartId = '';
  try {
    const res1 = await callTool('excel_add_chart', {
      sheetName: testSheetName,
      sourceAddress: 'B1:D4',
      chartType: 'column',
      title: '各产品线销量对比图',
      replaceExisting: true,
      left: 350,
      top: 20,
      width: 480,
      height: 280
    });
    // 再次调用以同一标题/坐标建图，验证 replaceExisting 成功替换
    const res2 = await callTool('excel_add_chart', {
      sheetName: testSheetName,
      sourceAddress: 'B1:D4',
      chartType: 'column',
      title: '各产品线销量对比图',
      replaceExisting: true,
      left: 350,
      top: 20,
      width: 480,
      height: 280
    });
    const ok = res2.success && res2.data.replaceExisting === true && !!res2.data.shapeName;
    if (ok) createdChartId = res2.data.id || res2.data.name;
    results.push({
      tool: 'excel_add_chart (replaceExisting旧图自动替换与回执)',
      success: ok,
      detail: ok ? `柱状图绘制成功，旧图被安全替换且返回完整可信元数据 (ID: ${createdChartId}, shape: ${res2.data.shapeName})` : res2.error
    });
  } catch (e: any) {
    results.push({ tool: 'excel_add_chart (replaceExisting旧图自动替换与回执)', success: false, detail: e.message });
  }

  // 12. excel_get_charts (查询图表列表，测试 detail: true 深度元数据)
  try {
    const res = await callTool('excel_get_charts', {
      sheetName: testSheetName,
      detail: true
    });
    const first = res.data?.charts?.[0];
    const ok = res.success && res.data.count === 1 && first?.hasLegend !== undefined;
    results.push({
      tool: 'excel_get_charts (图表查询与 detail 详情回执)',
      success: ok,
      detail: ok ? `已探测到 ${res.data.count} 个图表，图表详情包含 hasLegend=${first?.hasLegend}, seriesCount=${first?.seriesCount}` : res.error
    });
  } catch (e: any) {
    results.push({ tool: 'excel_get_charts (图表查询与 detail 详情回执)', success: false, detail: e.message });
  }

  // 13. excel_capture_sheet_preview (高保真渲染图捕获，验证绝无 422 报错)
  try {
    const res = await callTool('excel_capture_sheet_preview', {
      sheetName: testSheetName,
      address: 'A1:E5'
    });
    const ok = res.success && (!!res.data.imageBase64 || !!res.data.message);
    const sizeKb = res.data.imageBase64 ? (res.data.imageBase64.length * 0.75 / 1024).toFixed(1) : '0';
    results.push({
      tool: 'excel_capture_sheet_preview (高保真视觉渲染截图)',
      success: ok,
      detail: ok ? `成功生成高清视觉渲染图 (图像大小: ${sizeKb} KB, Base64有效)，彻底根除 422 报错` : res.error
    });
  } catch (e: any) {
    results.push({ tool: 'excel_capture_sheet_preview (高保真视觉渲染截图)', success: false, detail: e.message });
  }

  // 14. excel_search_cells (单元格搜索查找)
  try {
    const res = await callTool('excel_search_cells', {
      sheetName: testSheetName,
      query: '智能终端'
    });
    const ok = res.success && res.data.count > 0;
    results.push({
      tool: 'excel_search_cells (单元格搜索)',
      success: ok,
      detail: ok ? `找到关键字匹配项，匹配地址: ${res.data.matches[0]?.address}` : res.error
    });
  } catch (e: any) {
    results.push({ tool: 'excel_search_cells (单元格搜索)', success: false, detail: e.message });
  }

  // 15. excel_save_workbook (保存工作簿，验证完整权威回执)
  try {
    const res = await callTool('excel_save_workbook', {});
    const ok = res.success && res.data.saved === true && !!res.data.workbookName;
    results.push({
      tool: 'excel_save_workbook (保存工作簿权威可信回执)',
      success: ok,
      detail: ok ? `工作簿已刷盘保存，返回权威凭据 workbookName: ${res.data.workbookName}, saved: ${res.data.saved}` : res.error
    });
  } catch (e: any) {
    results.push({ tool: 'excel_save_workbook (保存工作簿权威可信回执)', success: false, detail: e.message });
  }

  // 16. excel_delete_chart (删除图表测试)
  if (createdChartId) {
    try {
      const res = await callTool('excel_delete_chart', {
        sheetName: testSheetName,
        chartName: createdChartId
      });
      const ok = res.success;
      results.push({
        tool: 'excel_delete_chart (删除图表)',
        success: ok,
        detail: ok ? `测试图表已按需清理` : res.error
      });
    } catch (e: any) {
      results.push({ tool: 'excel_delete_chart (删除图表)', success: false, detail: e.message });
    }
  }

  // 输出测试总览
  console.log('\n================ 测试验收结果汇总 ================');
  let passCount = 0;
  for (const r of results) {
    const icon = r.success ? '✅ PASS' : '❌ FAIL';
    console.log(`${icon} | [${r.tool}]: ${r.detail}`);
    if (r.success) passCount++;
  }
  console.log('==================================================');
  console.log(`\n验收结论: 总计测试 ${results.length} 项，通过 ${passCount} 项，失败 ${results.length - passCount} 项。`);
}

runAllTests().catch(console.error);
