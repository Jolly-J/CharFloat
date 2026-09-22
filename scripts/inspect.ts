import { bridgeServer } from "../src/bridge/ws-server.js";
import { UniversalGateway } from "../src/bridge/gateway.js";

async function main() {
  console.log("==================================================");
  console.log("   WPS Bridge 实时探测与连通性测试程序");
  console.log("==================================================");

  // 1. 启动本地 Bridge 服务
  try {
    composeBridgeServer();
    await bridgeServer.start();
    console.log("[成功] 本地 Bridge 网关已成功监听端口: 19890");
    console.log("[地址] OpenAPI 规范: http://127.0.0.1:19890/openapi.json");
    console.log("[地址] Tools 接口:   http://127.0.0.1:19890/api/v1/tools");
  } catch (err: any) {
    if (err.code === "EADDRINUSE") {
      console.log("[提示] 检测到端口 19890 已在后台运行，正在尝试复用连接...");
    } else {
      console.error("[错误] 启动端口监听失败:", err);
      return;
    }
  }

  console.log("\n正在等待 WPS 内部加载项连接 (请确保 WPS 表格已打开)...");

  // 轮询等待 WPS 连接，最长等待 90 秒
  let connected = false;
  for (let i = 1; i <= 90; i++) {
    const state = bridgeServer.getState();
    if (state.isWpsConnected) {
      connected = true;
      break;
    }
    await new Promise((r) => setTimeout(r, 1000));
    process.stdout.write(`\r[${i}s/90s] 等待 WPS 自动握手连接中...`);
  }

  console.log("\n");

  if (!connected) {
    console.log("[超时] 尚未检测到来自 WPS 的 WebSocket 连接。");
    console.log("\n【排查提示】：");
    console.log("1. 如果您在安装插件前就打开了 WPS，请完全退出 WPS（Cmd + Q），然后重新打开表格。");
    console.log("2. 检查 WPS 顶部功能区是否有「WPS Bridge (AI)」选项卡，点击上面的「连接状态」可手动触发重连。");
    process.exit(0);
  }

  console.log("[连接] 成功连入 WPS！正在读取当前表格上下文...\n");

  // 1. 获取工作区总览
  const summary = await UniversalGateway.executeTool("wps_get_workspace_summary", {}, "Probe Tester");
  console.log("---------------- 当前工作簿概览 ----------------");
  console.log(`工作簿名称: ${summary.workbookName || "未命名"}`);
  console.log(`完整路径:   ${summary.fullName || "—"}`);
  console.log(`激活的工作表: ${summary.activeSheetName || "—"}`);
  console.log(`包含的工作表: ${summary.sheets?.map((s: any) => s.name).join(", ") || "—"}`);
  if (summary.selection) {
    console.log(`当前鼠标选区: ${summary.selection.address} (${summary.selection.rowCount} 行 × ${summary.selection.columnCount} 列)`);
  }

  // 2. 获取当前表大纲
  if (summary.activeSheetName) {
    console.log("\n---------------- 工作表大纲指纹 ----------------");
    const outline = await UniversalGateway.executeTool("wps_get_sheet_outline", { sheetName: summary.activeSheetName }, "Probe Tester");
    console.log(`数据有效区域: ${outline.usedRangeAddress || "空表"}`);
    console.log(`总行数: ${outline.rowCount}, 总列数: ${outline.columnCount}`);
    if (outline.headerPreview && outline.headerPreview.length > 0) {
      console.log("表头/前三行预览:", JSON.stringify(outline.headerPreview));
    }
  }

  console.log("\n[完成] 探测完成！WPS Bridge 与当前 WPS 实时连接完全通畅！");
  process.exit(0);
}

main().catch((e) => {
  console.error("执行发生错误:", e);
  process.exit(1);
});
