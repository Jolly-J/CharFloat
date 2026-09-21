/**
 * WPS Bridge 全梯队能力集成验证脚本 (四大梯队端到端集成测试)
 */

const BRIDGE_HTTP = "http://127.0.0.1:19890";

async function callTool(name: string, args: any = {}) {
  const resp = await fetch(`${BRIDGE_HTTP}/api/v1/tool/call`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name,
      arguments: args,
      clientName: "Tier-4 Integration Verifier"
    })
  });
  const json: any = await resp.json();
  if (!json.success) {
    throw new Error(`[${name}] 调用失败: ${json.error}`);
  }
  return json.data;
}

async function runAllTiersVerification() {
  console.log("==================================================");
  console.log("   WPS Bridge 四大梯队核心能力全景验证执行");
  console.log("==================================================");

  // 0. 检查连接状态
  const statusResp = await fetch(`${BRIDGE_HTTP}/api/v1/status`);
  const status: any = await statusResp.json();
  if (!status.isWpsConnected) {
    console.warn("[提示] WPS 当前未连入 Bridge (守护进程端口 19890)。");
    console.warn("请确保 WPS Office 正在运行，且加载项已连接。");
    return;
  }
  console.log(`[连接] 当前活跃工作簿: [${status.activeWorkbook}]`);

  const testSheetName = "梯队能力全景看板";

  // 1. 第四梯队：新建工作表并设置标签底色
  console.log(`\n1. [第四梯队] 新建工作表【${testSheetName}】并设置标签底色为商务蓝...`);
  await callTool("wps_create_sheet", { sheetName: testSheetName });
  await callTool("wps_manage_sheet", {
    sheetName: testSheetName,
    action: "tab_color",
    color: "#1E3A8A"
  });

  // 2. 写入业务看板数据
  console.log("\n2. [写入数据] 写入标题、副标题、综合指标明细与结论卡片...");
  const tableValues = [
    ["2026年8月综合良率推移与缺陷分析看板", "", "", "", "", ""],
    ["统计口径: 制造一部 M23 全自动化生产线 | 更新周期: 每日滚动更新", "", "", "", "", ""],
    ["日期", "投产片数", "良品片数", "综合良率", "主要异常分类", "处理状态"],
    ["2026-08-01", 5200, 4880, 0.9385, "硅片碎片", "已解决"],
    ["2026-08-02", 5350, 4920, 0.9196, "隐裂", "已解决"],
    ["2026-08-03", 5180, 4610, 0.8899, "虚焊 (超出控制限)", "跟进中"],
    ["2026-08-04", 5400, 5080, 0.9407, "硅片碎片", "已解决"],
    ["2026-08-05", 5320, 4980, 0.9361, "印刷对齐偏差", "已解决"],
    ["2026-08-06", 5290, 4655, 0.8799, "热斑烧结不良", "待复测"],
    ["2026-08-07", 5410, 5130, 0.9482, "外观脏污", "已解决"],
    ["合计 / 均值", 37150, 34255, "=C11/B11", "累计不良分布", "全部完成"],
    ["", "", "", "", "", ""],
    ["管理诊断建议：8月3日与8月6日出现良率异常下滑（低于90%告警阈值），主要集中在虚焊与热斑烧结环节。建议设备工程部调校第三温区热场分布，并对串焊机吸盘进行负压检测。", "", "", "", "", ""]
  ];

  await callTool("wps_patch_cells", {
    sheetName: testSheetName,
    address: "B2:G14",
    values: tableValues,
    reason: "写入全景验证看板数据"
  });

  // 3. 第一梯队：排版美学与合并单元格
  console.log("\n3. [第一梯队] 执行大标题/副标题跨列居中合并、结论卡片整行横跨...");
  // 主标题 B2:G2
  await callTool("wps_format_cells", {
    sheetName: testSheetName,
    address: "B2:G2",
    fontSize: 15,
    bold: true,
    fontColor: "#0F172A",
    horizontalAlignment: "center",
    merge: true,
    rowHeight: 36
  });

  // 副标题 B3:G3
  await callTool("wps_format_cells", {
    sheetName: testSheetName,
    address: "B3:G3",
    fontSize: 9.5,
    fontColor: "#64748B",
    horizontalAlignment: "center",
    merge: true,
    rowHeight: 20
  });

  // 表头 B4:G4
  await callTool("wps_format_cells", {
    sheetName: testSheetName,
    address: "B4:G4",
    fontSize: 10.5,
    bold: true,
    backgroundColor: "#0F172A",
    fontColor: "#FFFFFF",
    horizontalAlignment: "center",
    rowHeight: 28
  });

  // 数据对齐与数字掩码
  await callTool("wps_format_cells", {
    sheetName: testSheetName,
    address: "B5:B11",
    horizontalAlignment: "center",
    numberFormat: "yyyy-mm-dd"
  });
  await callTool("wps_format_cells", {
    sheetName: testSheetName,
    address: "C5:D11",
    horizontalAlignment: "right",
    numberFormat: "#,##0"
  });
  await callTool("wps_format_cells", {
    sheetName: testSheetName,
    address: "E5:E11",
    horizontalAlignment: "right",
    numberFormat: "0.00%"
  });
  await callTool("wps_format_cells", {
    sheetName: testSheetName,
    address: "G5:G11",
    horizontalAlignment: "center"
  });

  // 合计行 B11:G11
  await callTool("wps_format_cells", {
    sheetName: testSheetName,
    address: "B11:G11",
    bold: true,
    backgroundColor: "#F1F5F9",
    rowHeight: 24,
    borders: "#94A3B8"
  });

  // 诊断建议整行合并 B14:G14
  await callTool("wps_format_cells", {
    sheetName: testSheetName,
    address: "B14:G14",
    merge: true,
    wrapText: true,
    rowHeight: 32,
    backgroundColor: "#FEF3C7",
    fontColor: "#92400E",
    fontSize: 9.5
  });

  // 4. 第一梯队：条件格式告警 (良率 < 90% 浅红高亮)
  console.log("\n4. [第一梯队] 添加条件格式：良率 < 0.90 自动浅红背景告警...");
  await callTool("wps_add_conditional_formatting", {
    sheetName: testSheetName,
    address: "E5:E10",
    ruleType: "cell_value",
    operator: "less_than",
    formula1: "0.90",
    backgroundColor: "#FEE2E2",
    fontColor: "#991B1B"
  });

  // 5. 第一梯队：冻结窗格表头吸顶
  console.log("\n5. [第一梯队] 冻结窗格吸顶锁定第 4 行表头...");
  await callTool("wps_freeze_panes", {
    sheetName: testSheetName,
    freezeRowIndex: 5
  });

  // 6. 第二梯队：原生矢量图表嵌入 (并列右侧 I4，开启平滑曲线、品牌系列色、Y轴范围锁定)
  console.log("\n6. [第二梯队] 嵌入原生折线走势图，锚定单元格 I4，配置平滑曲线、系列颜色与 Y 轴刻度锁定...");
  await callTool("wps_add_chart", {
    sheetName: testSheetName,
    chartType: "line",
    dataRange: "B4:E10",
    title: "8月投产与良率推移趋势",
    position: {
      leftCell: "I4",
      width: 520,
      height: 300
    },
    hasLegend: true,
    hasDataLabels: false,
    smoothLine: true,
    seriesColors: ["#3B82F6", "#10B981", "#F59E0B"],
    yAxis: {
      min: 0.85,
      max: 1.00,
      step: 0.05,
      numberFormat: "0.0%",
      title: "良率 / 达成比率"
    }
  });

  // 7. 第三梯队：开启表头自动筛选与初始排序
  console.log("\n7. [第三梯队] 开启表头自动筛选...");
  await callTool("wps_set_filter_and_sort", {
    sheetName: testSheetName,
    range: "B4:G10",
    enableAutoFilter: true
  });

  // 8. 第三梯队：处理状态列注入下拉验证
  console.log("\n8. [第三梯队] 在处理状态列 (G5:G10) 注入下拉列表有效性验证...");
  await callTool("wps_set_data_validation", {
    sheetName: testSheetName,
    address: "G5:G10",
    validationType: "list",
    listItems: ["已解决", "跟进中", "待复测", "需挂牌"],
    promptTitle: "状态选择",
    promptMessage: "请从标准化选项中选择当前处理状态"
  });

  // 9. 智能列宽适配
  console.log("\n9. 执行智能自适应列宽排版...");
  await callTool("wps_auto_fit_columns", {
    sheetName: testSheetName,
    columnRules: [
      { colIndex: 2, minWidth: 14, maxWidth: 16 },
      { colIndex: 3, minWidth: 12, maxWidth: 16 },
      { colIndex: 4, minWidth: 12, maxWidth: 16 },
      { colIndex: 5, minWidth: 13, maxWidth: 16 },
      { colIndex: 6, minWidth: 18, maxWidth: 26 },
      { colIndex: 7, minWidth: 12, maxWidth: 16 }
    ]
  });

  console.log("\n==================================================");
  console.log("   四大梯队全部新能力验证流水线执行完毕！");
  console.log("==================================================");
}

runAllTiersVerification().catch(console.error);
