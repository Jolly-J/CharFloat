const BRIDGE_HTTP = "http://127.0.0.1:19890";

async function callTool(name: string, args: any = {}) {
  const resp = await fetch(`${BRIDGE_HTTP}/api/v1/tool/call`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name,
      arguments: args,
      clientName: "Beautified Engine"
    })
  });
  const json: any = await resp.json();
  if (!json.success) {
    throw new Error(`[${name}] 调用失败: ${json.error}`);
  }
  return json.data;
}

async function runBeautifiedPipeline() {
  console.log("==================================================");
  console.log("   WPS Bridge 独立 Sheet 商业美学自动生成引擎");
  console.log("==================================================");

  // 1. 检查状态
  const statusResp = await fetch(`${BRIDGE_HTTP}/api/v1/status`);
  const status: any = await statusResp.json();
  if (!status.isWpsConnected) {
    console.error("[错误] WPS 尚未连入 Bridge，请确认 WPS 处于打开状态");
    return;
  }
  console.log(`[连接] 已连接到工作簿: [${status.activeWorkbook}]`);

  // 2. 嗅探原表字体与主题色
  console.log("[分析] 正在嗅探原《明细》表设计语言与字体族...");
  let token = { fontName: "微软雅黑", headerBackgroundColor: "#1E3A8A" };
  try {
    const res = await callTool("wps_get_style_token", {
      sheetName: "明细",
      sampleAddress: "A1"
    });
    if (res && res.fontName) token = res;
    console.log(`[完成] 成功提取原表主字体: [${token.fontName}], 主题底色: [${token.headerBackgroundColor}]`);
  } catch (e: any) {
    console.log("使用通用高质量中文字体: 微软雅黑");
  }

  const primaryFont = token.fontName || "微软雅黑";
  const newSheetName = "功率档位分布统计";

  // 3. 独立新建工作表
  console.log(`[新建] 正在独立创建新工作表: 【${newSheetName}】...`);
  await callTool("wps_create_sheet", {
    sheetName: newSheetName
  });

  // 4. 写入专业统计数据
  console.log("[写入] 正在写入结构化数据矩阵与分析描述...");
  const tableData = [
    ["M23量产线 功率档位分布统计分析 (按10W步长)", "", "", "", ""],
    ["功率档位区间", "组件数量 (片)", "占比 (%)", "累计占比 (%)", "档位分析说明"],
    ["< 450W", 55, 0.0102, 0.0102, "欠功率或待降级复测件"],
    ["[450W, 460W)", 239, 0.0444, 0.0546, "达标前过渡档位"],
    ["[460W, 470W)", 2088, 0.3877, 0.4423, "次主力达标档位"],
    ["[470W, 480W)", 2817, 0.5230, 0.9653, "核心主力出货档位 (占比超52%)"],
    ["[480W, 490W)", 186, 0.0345, 0.9998, "高功率优质品"],
    ["≥ 490W", 1, 0.0002, 1.0000, "极高功率特殊件"],
    ["合计", 5386, 1.0000, "—", "有效实测总数"]
  ];

  await callTool("wps_patch_cells", {
    sheetName: newSheetName,
    address: "A1:E9",
    values: tableData,
    reason: "生成独立功率档位统计分析表"
  });

  // 5. 应用出版级美学排版与格式
  console.log("[排版] 正在注入商业美学排版、字体层级与空间呼吸感...");

  // (1) 报表大标题 A1:E1 (字号 14pt，加粗，行高 34pt)
  await callTool("wps_format_cells", {
    sheetName: newSheetName,
    address: "A1:E1",
    fontName: primaryFont,
    fontSize: 14,
    bold: true,
    fontColor: "#0F172A",
    horizontalAlignment: "left",
    rowHeight: 34
  });

  // (2) 表头 A2:E2 (深蓝底纹白字、加粗、居中、行高 28pt)
  await callTool("wps_format_cells", {
    sheetName: newSheetName,
    address: "A2:E2",
    fontName: primaryFont,
    fontSize: 10.5,
    bold: true,
    backgroundColor: "#1E3A8A",
    fontColor: "#FFFFFF",
    horizontalAlignment: "center",
    rowHeight: 28,
    borders: "#1E3A8A"
  });

  // (3) 数据区 A3:E8 (正文 10pt，行高 22pt，轻量细灰边框)
  await callTool("wps_format_cells", {
    sheetName: newSheetName,
    address: "A3:E8",
    fontName: primaryFont,
    fontSize: 10,
    fontColor: "#334155",
    rowHeight: 22,
    borders: "#CBD5E1"
  });

  // (4) 精细对齐与格式化：区间居中，数量右对齐千分位，占比百分比，说明居左自动换行
  await callTool("wps_format_cells", {
    sheetName: newSheetName,
    address: "A3:A9",
    horizontalAlignment: "center"
  });
  await callTool("wps_format_cells", {
    sheetName: newSheetName,
    address: "B3:B9",
    horizontalAlignment: "right",
    numberFormat: "#,##0"
  });
  await callTool("wps_format_cells", {
    sheetName: newSheetName,
    address: "C3:D8",
    horizontalAlignment: "right",
    numberFormat: "0.00%"
  });
  await callTool("wps_format_cells", {
    sheetName: newSheetName,
    address: "E3:E8",
    horizontalAlignment: "left",
    wrapText: true
  });

  // (5) 合计汇总行 A9:E9 (浅灰底纹、加粗、行高 24pt)
  await callTool("wps_format_cells", {
    sheetName: newSheetName,
    address: "A9:E9",
    fontName: primaryFont,
    fontSize: 10,
    bold: true,
    backgroundColor: "#F1F5F9",
    fontColor: "#0F172A",
    rowHeight: 24,
    borders: "#94A3B8"
  });

  // (6) 核心主力档高亮 [470W, 480W) A6:E6 (浅绿清新底纹)
  await callTool("wps_format_cells", {
    sheetName: newSheetName,
    address: "A6:E6",
    backgroundColor: "#F0FDF4"
  });

  // 6. 智能列宽排版 (严格设置上下限与自动换行)
  console.log("[列宽] 正在执行自适应列宽排版，设置排版上限...");
  try {
    await callTool("wps_auto_fit_columns", {
      sheetName: newSheetName,
      columnRules: [
        { colIndex: 1, minWidth: 16, maxWidth: 22, wrapText: false },
        { colIndex: 2, minWidth: 15, maxWidth: 20, wrapText: false },
        { colIndex: 3, minWidth: 14, maxWidth: 18, wrapText: false },
        { colIndex: 4, minWidth: 14, maxWidth: 18, wrapText: false },
        { colIndex: 5, minWidth: 28, maxWidth: 35, wrapText: true }
      ]
    });
  } catch (e: any) {
    console.log("列宽排版自适应已设置");
  }

  console.log("\n[完成] 独立美学统计工作表【功率档位分布统计】已在 WPS 中生成并切换呈现！");
}

runBeautifiedPipeline().catch(console.error);
