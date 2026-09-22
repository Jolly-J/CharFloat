// 本文件由 scripts/build-office-addon.mjs 生成，请勿手改；改动请改 office-addon/src/**
// 构建命令: node scripts/build-office-addon.mjs（按固定顺序拼接 src 下 19 个片段，无打包器、无依赖）
// 拼接顺序:
//   01. src/state.js
//   02. src/ui.js
//   03. src/lifecycle.js
//   04. src/connection.js
//   05. src/rpc.js
//   06. src/excel/shared.js
//   07. src/excel/workbook.js
//   08. src/excel/sheets.js
//   09. src/excel/range.js
//   10. src/excel/formula.js
//   11. src/excel/format.js
//   12. src/excel/filter-sort.js
//   13. src/excel/table.js
//   14. src/excel/chart.js
//   15. src/excel/pivot.js
//   16. src/excel/shape.js
//   17. src/excel/comment.js
//   18. src/excel/script.js
//   19. src/bootstrap.js

/**
 * WPS Bridge - Microsoft Office (Excel) 官方 Office.js 核心运行时
 * 运行在 Microsoft Excel 任务窗格 WebView (WebKit / Edge WebView2)
 */

(function () {

// ── 模块: src/state.js — 全局常量与跨模块共享状态 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 7-15 行（原样搬迁，未改写）。
  const ADDON_VERSION = "2.0.0";
  // 部署探针：区分"磁盘上已构建的新能力"与"Excel 当前跑的那份窗格代码"。
  // 任务窗格是宿主 WebView 里长期驻留的页面，改完源码不重开窗格就不会生效；
  // 这个探针让任意 host=microsoft 的读工具都能回答"窗格里有没有本轮新分支"。
  const ADDON_CAPABILITIES = { cap08Shapes: true, shapeMethods: ["add_shape", "list_shapes", "update_shape", "group_shapes", "ungroup_shapes", "set_shape_zorder", "get_active_shape", "export_shape_image", "shape_self_test", "insert_image", "reload"] };
  const DEFAULT_WS_URL = "wss://localhost:19891/office-addon";
  let ws = null;
  let reconnectTimer = null;
  let isConnected = false;
  let activeWorkbookName = "—";
  let activeSheetName = "—";
  let activeSelectionAddress = "—";

// ── 模块: src/ui.js — 任务窗格 UI 渲染：运行日志、活动流、状态徽标 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 16-56 行（原样搬迁，未改写）。
  function log(msg, data) {
    const timestamp = new Date().toLocaleTimeString();
    const formatted = `[${timestamp}] ${msg}` + (data ? ` ${JSON.stringify(data)}` : "");
    console.log(formatted);
    const logEl = document.getElementById("logArea");
    if (logEl) {
      logEl.innerText = (formatted + "\n" + logEl.innerText).slice(0, 5000);
    }
  }

  function addActivityItem(actionTitle) {
    const feed = document.getElementById("activityFeed");
    if (!feed) return;
    const empty = feed.querySelector(".feed-empty");
    if (empty) empty.remove();
    const now = new Date().toLocaleTimeString("zh-CN", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
    const item = document.createElement("div");
    item.className = "feed-item";
    item.innerHTML = `<span class="feed-item-action">${actionTitle}</span><span class="feed-item-time">${now}</span>`;
    feed.prepend(item);
    while (feed.children.length > 8) {
      feed.removeChild(feed.lastChild);
    }
  }

  function updateStatusUI(connected, text) {
    const badge = document.getElementById("connectionBadge");
    const connectionText = document.getElementById("connectionText");
    const docTitle = document.getElementById("docTitle");
    const activeSheetEl = document.getElementById("activeSheet");
    const activeSelEl = document.getElementById("activeSelection");

    if (badge && connectionText) {
      badge.className = "status-badge " + (connected ? "connected" : text === "连接中…" ? "connecting" : "error");
      connectionText.innerText = connected ? "已就绪" : (text || "未连接");
    }
    if (docTitle) docTitle.innerText = activeWorkbookName;
    if (activeSheetEl) activeSheetEl.innerText = activeSheetName;
    if (activeSelEl) activeSelEl.innerText = activeSelectionAddress;
  }

// ── 模块: src/lifecycle.js — Office.js 生命周期与宿主状态感知 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 57-116 行（原样搬迁，未改写）。
  // ==========================================
  // Office.js 初始化与状态感知
  // ==========================================

  Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
      log("Office.js 环境就绪 (Microsoft Excel)");
      initExcelEventHooks();
      connect();
    } else {
      log("当前宿主不是 Excel: " + info.host);
      updateStatusUI(false, "非 Excel 环境");
    }
  });

  async function updateActiveSummary() {
    try {
      const docUrl = Office?.context?.document?.url;
      if (docUrl) {
        const decoded = decodeURIComponent(docUrl.split("/").pop().split("\\").pop());
        if (decoded) activeWorkbookName = decoded;
      }
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        const selection = context.workbook.getSelectedRange();
        selection.load("address");
        await context.sync();

        if (activeWorkbookName === "—") activeWorkbookName = "工作簿1.xlsx";
        activeSheetName = sheet.name || "Sheet1";
        activeSelectionAddress = selection.address || "—";
        updateStatusUI(isConnected);
      });
    } catch (e) {}
  }

  function initExcelEventHooks() {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onSelectionChanged.add(async () => {
        await updateActiveSummary();
        if (ws && isConnected) {
          sendPacket({
            type: "event",
            event: "selection_change",
            host: "microsoft",
            data: {
              workbookName: activeWorkbookName,
              sheetName: activeSheetName,
              address: activeSelectionAddress
            }
          });
        }
      });
      await context.sync();
      await updateActiveSummary();
    }).catch(() => {});
  }

// ── 模块: src/connection.js — WebSocket 连接、重连、注册与数据包发送 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 117-184 行（原样搬迁，未改写）。
  // ==========================================
  // WebSocket 通信与注册
  // ==========================================

  function connect() {
    if (ws && (ws.readyState === WebSocket.CONNECTING || ws.readyState === WebSocket.OPEN)) return;

    try {
      log("正在连接 Bridge 服务端: " + DEFAULT_WS_URL);
      ws = new WebSocket(DEFAULT_WS_URL);

      ws.onopen = async function () {
        isConnected = true;
        log("已成功连上 WPS Bridge 服务端 (MS Office 通道)");
        await updateActiveSummary();

        sendPacket({
          type: "register",
          client: "ms-excel-addon",
          host: "microsoft",
          version: ADDON_VERSION,
          summary: {
            workbookName: activeWorkbookName,
            activeSheetName: activeSheetName,
            selection: { address: activeSelectionAddress }
          }
        });
        updateStatusUI(true);
      };

      ws.onmessage = function (event) {
        handleIncomingMessage(event.data);
      };

      ws.onerror = function (err) {
        log("WebSocket 异常", err);
      };

      ws.onclose = function () {
        isConnected = false;
        log("连接断开，2.5 秒后自动重试…");
        updateStatusUI(false, "等待连接…");
        scheduleReconnect();
      };
    } catch (e) {
      log("创建 WebSocket 异常: " + e.message);
      scheduleReconnect();
    }
  }

  function scheduleReconnect() {
    if (reconnectTimer) clearTimeout(reconnectTimer);
    reconnectTimer = setTimeout(connect, 2500);
  }

  window.reconnect = function () {
    if (ws) {
      try { ws.close(); } catch (e) {}
    }
    connect();
  };

  function sendPacket(packet) {
    if (ws && ws.readyState === WebSocket.OPEN) {
      ws.send(JSON.stringify(packet));
    }
  }

// ── 模块: src/rpc.js — RPC 入站分发与 dispatchExcelTool 方法路由 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 185-388 行（原样搬迁，未改写）。
  // ==========================================
  // RPC 分发中心与全量能力执行
  // ==========================================

  async function handleIncomingMessage(rawText) {
    let message;
    try {
      message = JSON.parse(rawText);
    } catch (e) {
      return;
    }

    const id = message.id;
    if (!id) return;

    const rawMethod = String(message.method || "");
    const method = rawMethod.replace(/^(wps|excel)[._]/, "").toLowerCase();
    const params = message.params || {};

    const readableActionMap = {
      'read_range': '读取单元格数据',
      'write_range': '写入数据到表格',
      'create_chart': '生成并排版图表',
      'create_table': '创建结构化表格',
      'sort_range': '执行数据排序',
      'add_comment': '添加批注与审阅',
      'add_worksheet': '新建工作表',
      'delete_worksheet': '删除工作表',
      'set_range_format': '设置单元格样式',
      'calculate': '重新计算公式'
    };
    const actionDesc = readableActionMap[method] || `协同操作 [${rawMethod}]`;
    addActivityItem(actionDesc);
    log(`收到指令 [${rawMethod}]`, params);

    try {
      const result = await dispatchExcelTool(method, params);
      sendPacket({
        id,
        type: "rpc_response",
        result: result === undefined ? { success: true } : result
      });
      log(`指令执行成功 [${rawMethod}]`);
    } catch (error) {
      log(`指令执行失败 [${rawMethod}]: ` + error.message);
      sendPacket({
        id,
        type: "rpc_response",
        error: error.message || String(error)
      });
    }
  }

  async function dispatchExcelTool(method, params) {
    switch (method) {
      // 1. 工作簿与元数据
      case "get_workbook_info":
      case "get_workbook_state":
      case "get_workspace_summary":
        return await handleGetWorkbookInfo(params);
      case "get_sheet_outline":
        return await handleGetSheetOutline(params);
      case "save":
      case "save_workbook":
        return await handleSave(params);
      case "calculate":
        return await handleCalculate(params);
      case "list_named_items":
        return await handleListNamedItems(params);
      case "update_named_item":
        return await handleUpdateNamedItem(params);
      case "get_document_properties":
        return await handleGetProperties(params);
      case "update_document_properties":
        return await handleUpdateProperties(params);

      // 2. 工作表生命周期
      case "list_sheets":
      case "get_sheets":
        return await handleListSheets(params);
      case "add_sheet":
      case "create_sheet":
        return await handleAddSheet(params);
      case "update_sheet":
      case "rename_sheet":
        return await handleUpdateSheet(params);
      case "delete_sheet":
        return await handleDeleteSheet(params);
      case "copy_sheet":
      case "duplicate_sheet":
      case "manage_sheet":
        return await handleManageSheet(params);

      // 3. 区域读写与结构操作
      case "read_range":
      case "get_range_data":
      case "get_used_range":
        return await handleReadRange(params);
      case "get_range_styles":
        return await handleGetRangeStyles(params);
      case "write_range":
      case "patch_cells":
      case "set_range_data":
        return await handleWriteRange(params);
      case "update_range_structure":
      case "insert_rows":
      case "delete_rows":
      case "modify_rows_columns":
      case "manage_rows_and_columns":
      case "insert_dimension":
        return await handleUpdateRangeStructure(params);
      case "clear_range":
      case "clear_cells":
        return await handleClearRange(params);
      case "copy_range":
        return await handleCopyRange(params);
      case "set_hyperlink":
        return await handleSetHyperlink(params);
      case "set_data_validation":
        return await handleSetDataValidation(params);
      case "find_replace_cells":
      case "find_and_replace":
      case "find_replace":
      case "search_cells":
        return await handleFindReplace(params);

      // 4. 公式
      case "set_formula":
      case "set_cell_formula":
        return await handleSetFormula(params);

      // 5. 格式与排版
      case "format_cells":
      case "format_range":
      case "set_range_format":
        return await handleFormatRange(params);
      case "auto_fit_columns":
      case "autofit_columns":
        return await handleAutofitColumns(params);
      case "freeze_panes":
        return await handleFreezePanes(params);
      case "add_conditional_formatting":
        return await handleAddConditionalFormatting(params);
      // 危险别名修复（问题台账 ISS-91）：`list_conditional_formats` / `update_conditional_format`
      // 原来和"新增条件格式"走同一个写处理函数 —— 调"读条件格式"会**新增一条规则**。
      // Office.js 侧暂无对应的读/改实现，这里**显式报错**，绝不再落到写路径。
      case "list_conditional_formats":
      case "update_conditional_format":
        throw new Error(
          `Office.js 通道尚未实现 "${method}"：原先它会落到"新增条件格式"的写路径，造成误写，已阻断。` +
          `如需读取或修改条件格式，请改用 host=wps 的对应能力，或先用 wps_execute_script 探测。`
        );

      // 6. 排序与筛选
      case "sort_range":
      case "apply_filter":
      case "set_filter_and_sort":
        return await handleSetFilterAndSort(params);

      // 7. 表格 Table
      case "create_table":
        return await handleCreateTable(params);
      case "update_table":
        return await handleUpdateTable(params);

      // 8. 图表 Chart
      case "get_charts":
        return await handleGetCharts(params);
      case "add_chart":
      case "create_chart":
        return await handleCreateChart(params);
      case "delete_chart":
        return await handleDeleteChart(params);
      case "update_chart":
        return await handleUpdateChart(params);
      case "export_chart_image":
        return await handleExportChartImage(params);
      case "capture_sheet_preview":
        return await handleCaptureSheetPreview(params);

      // 9. 数据透视表 PivotTable
      case "create_pivot_table":
        return await handleCreatePivotTable(params);
      case "update_pivot_table":
        return await handleUpdatePivotTable(params);

      // 10. 形状与图片（CAP-08：MS 侧矢量绘图）
      case "insert_image":
        return await handleInsertImage(params);
      case "list_shapes":
        return await handleListShapes(params);
      case "update_shape":
        return await handleUpdateShape(params);
      case "add_shape":
      case "add_geometric_shape":
        return await handleAddShape(params);
      case "group_shapes":
        return await handleGroupShapes(params);
      case "ungroup_shapes":
        return await handleUngroupShapes(params);
      case "set_shape_zorder":
      case "set_z_order":
        return await handleSetShapeZOrder(params);
      case "get_active_shape":
        return await handleGetActiveShape(params);
      case "export_shape_image":
        return await handleExportShapeImage(params);
      case "shape_self_test":
      case "probe_shape_api":
        return await handleShapeSelfTest(params);

      // 11. 审阅与批注
      // 危险别名修复（问题台账 ISS-92）：`handleManageComments` 的 action 默认是 "add"，
      // 原来 `list_comments` 走同一条路 —— 调"列批注"会在 A1 **插一条空批注**；`update_comment` 则静默返回成功。
      case "manage_cell_comments":
      case "add_comment":
        return await handleManageComments({ ...params, action: params.action || "add" });
      case "list_comments":
        return await handleManageComments({ ...params, action: "list" });
      case "update_comment":
        throw new Error(
          'Office.js 通道尚未实现 "update_comment"：原实现会落到未知 action 的静默成功分支，不做任何事却报成功，已阻断。' +
          '如需修改批注，请先 list_comments 取 id、delete 后再 add。'
        );

      // 12. 任意脚本自由运行
      case "run_script":
      case "execute_script":
        return await handleRunScript(params);
      // 热重载：与 WPS 加载项的同名分支对齐（`wps_reload_addon` 走的就是 reload）。
      // 这里不再用 `window.location.reload(true)`：forceGet 参数已废弃，WKWebView 不保证重新拉取
      // 子资源。改写 URL query 再跳转，保证 taskpane.html / taskpane.js 真的重新请求。
      case "reload":
      case "reload_addon": {
        try {
          const url = new URL(window.location.href);
          url.searchParams.set("_reload", String(Date.now()));
          setTimeout(() => window.location.replace(url.toString()), 50);
          return { success: true, reloading: true, target: url.toString() };
        } catch (e) {
          setTimeout(() => window.location.reload(), 50);
          return { success: true, reloading: true, fallback: "location.reload()" };
        }
      }

      default:
        throw new Error(`Office.js 暂未映射该工具：${method}`);
    }
  }

// ── 模块: src/excel/shared.js — Excel 操作公共辅助（目标工作表解析） ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 389-399 行（原样搬迁，未改写）。
  // ==========================================
  // 具体 API 业务实现 (借助 Excel.run)
  // ==========================================

  function getTargetSheet(context, sheetName) {
    // 顺手把 `name` 排进加载队列：调用方普遍会把 `sheet.name` 写进返回体，
    // 而 Office.js 读未 load 的属性会直接抛「属性"name"不可用」。
    // 在这里统一 load 一次，避免每个处理器各写一遍（真机实测踩到过：形状工具全部报该错）。
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    return sheet;
  }

// ── 模块: src/excel/workbook.js — 工作簿元数据、保存、计算、命名项与文档属性 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 400-531 行（原样搬迁，未改写）。
  // 1. 工作簿操作
  async function handleGetWorkbookInfo(params) {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets.load("items/name, items/visibility, items/position");
      const names = context.workbook.names.load("items/name");
      const activeSheet = context.workbook.worksheets.getActiveWorksheet().load("name");
      try {
        context.workbook.load("name");
      } catch {}
      await context.sync();

      let wbName = (params && params.workbookName) || "工作簿1.xlsx";
      try {
        if (context.workbook.name) wbName = context.workbook.name;
      } catch {}

      return {
        hasOpenWorkbook: true,
        workbookName: wbName,
        fullName: wbName,
        activeSheet: activeSheet.name,
        activeSheetName: activeSheet.name,
        sheetCount: sheets.items.length,
        sheets: sheets.items.map(s => ({ name: s.name, visibility: s.visibility, position: s.position })),
        namedItemCount: names.items.length,
        namedItems: names.items.map(n => n.name)
      };
    });
  }

  async function handleGetSheetOutline(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params?.sheetName);
      sheet.load("name");
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("address, rowCount, columnCount, rowIndex, columnIndex, values");
      await context.sync();

      if (usedRange.isNullObject) {
        return {
          sheetName: sheet.name,
          isEmpty: true,
          usedRangeAddress: "A1",
          startRow: 1,
          startColumn: 1,
          rowCount: 0,
          columnCount: 0,
          headerPreview: []
        };
      }

      const headers = (usedRange.values && usedRange.values.length > 0) ? [usedRange.values[0]] : [];
      return {
        sheetName: sheet.name,
        isEmpty: false,
        usedRangeAddress: usedRange.address,
        startRow: usedRange.rowIndex + 1,
        startColumn: usedRange.columnIndex + 1,
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount,
        headerPreview: headers
      };
    });
  }

  async function handleSave(params) {
    return await Excel.run(async (context) => {
      context.workbook.save();
      try {
        context.workbook.load("name");
      } catch {}
      await context.sync();
      const wbName = context.workbook.name || (params && params.workbookName) || "工作簿1.xlsx";
      return {
        success: true,
        saved: true,
        workbookName: wbName,
        fullName: wbName,
        message: `工作簿 [${wbName}] 已成功保存到磁盘`
      };
    });
  }

  async function handleCalculate(params) {
    return await Excel.run(async (context) => {
      const type = params.type || "FullRebuild";
      context.application.calculate(type);
      await context.sync();
      return { success: true, calculationType: type };
    });
  }

  async function handleListNamedItems() {
    return await Excel.run(async (context) => {
      const items = context.workbook.names.load("items/name, items/value, items/type, items/visible");
      await context.sync();
      return {
        items: items.items.map(i => ({ name: i.name, value: i.value, type: i.type, visible: i.visible }))
      };
    });
  }

  async function handleUpdateNamedItem(params) {
    return await Excel.run(async (context) => {
      const { name, reference, comment, action } = params;
      if (action === "delete") {
        context.workbook.names.getItem(name).delete();
      } else {
        context.workbook.names.add(name, reference, comment);
      }
      await context.sync();
      return { success: true, name };
    });
  }

  async function handleGetProperties() {
    return await Excel.run(async (context) => {
      const p = context.workbook.properties.load("title, subject, author, keywords, comments, category, manager, company");
      await context.sync();
      return { properties: p };
    });
  }

  async function handleUpdateProperties(params) {
    return await Excel.run(async (context) => {
      const props = params.properties || params;
      context.workbook.properties.set(props);
      await context.sync();
      return { success: true, updated: props };
    });
  }

// ── 模块: src/excel/sheets.js — 工作表生命周期 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 532-613 行（原样搬迁，未改写）。
  // 2. 工作表操作
  async function handleListSheets() {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets.load("items/name, items/visibility, items/tabColor, items/position");
      await context.sync();
      return {
        sheets: sheets.items.map(s => ({ name: s.name, visibility: s.visibility, tabColor: s.tabColor, position: s.position }))
      };
    });
  }

  async function handleAddSheet(params) {
    return await Excel.run(async (context) => {
      const name = params.name || params.sheetName;
      const sheet = context.workbook.worksheets.add(name);
      sheet.load("name, position");
      await context.sync();
      return { success: true, name: sheet.name, position: sheet.position };
    });
  }

  async function handleUpdateSheet(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName || params.oldName);
      if (params.name || params.newName) sheet.name = params.name || params.newName;
      if (params.visibility) sheet.visibility = params.visibility;
      if (params.tabColor || params.color) sheet.tabColor = params.tabColor || params.color;
      if (params.activate) sheet.activate();
      await context.sync();
      return { success: true };
    });
  }

  async function handleDeleteSheet(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName || params.name);
      sheet.delete();
      await context.sync();
      return { success: true };
    });
  }

  async function handleCopySheet(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sourceSheet);
      const copied = sheet.copy(params.positionType || "After", sheet);
      if (params.newName) copied.name = params.newName;
      copied.load("name");
      await context.sync();
      return { success: true, name: copied.name };
    });
  }

  /** 取整数，失败返回 null（用于把 targetIndex / position / index 归一）。 */
  function readSheetNumber(...candidates) {
    for (const raw of candidates) {
      if (raw === undefined || raw === null || raw === "") continue;
      const n = Number(raw);
      if (Number.isFinite(n)) return n;
    }
    return null;
  }

  async function handleManageSheet(params) {
    return await Excel.run(async (context) => {
      const rawAction = String(params.action || "").trim();
      const action = rawAction || (params.newSheetName || params.newName ? "rename" : "");
      if (!action) {
        throw new Error(
          '[Office.js 通道] manage_sheet 缺少 action：支持 rename | move | tab_color | protect | unprotect。' +
          '原实现在缺失 action 时会返回 success 却什么都不做，已阻断。'
        );
      }

      const sheet = getTargetSheet(context, params.sheetName || params.oldName || params.sourceSheet);
      sheet.load("name");
      await context.sync();
      const sheetName = sheet.name;

      // 归一化 action 别名：网关用 WPS 语义（rename/move/tab_color/protect/unprotect），
      // 早期 Office.js 实现用 copy/hide/show/color —— 两套都认，避免任何一侧改名后静默 no-op（ISS-93）。
      const actionAliases = {
        copy: "copy", duplicate: "copy", clone: "copy",
        rename: "rename", set_name: "rename",
        move: "move", activate: "move", set_position: "move", reorder: "move",
        tab_color: "tab_color", color: "tab_color", set_tab_color: "tab_color",
        protect: "protect", unprotect: "unprotect",
        hide: "hide", show: "show", unhide: "show", visible: "show"
      };
      const normalized = actionAliases[action.toLowerCase()];

      const unsupported = {
        protect: '[Office.js 通道] manage_sheet(action="protect") 在 Microsoft 宿主未实现：Excel Office.js 的 WorksheetProtection 在 Excel on Mac 桌面版不可用/不可靠，本通道拒绝执行以免回报假成功。替代路径：① 改用 host="wps" 的 wps_manage_sheet(action="protect")；② 在 Excel 里用「审阅 → 保护工作表」手动加保护，然后 read_range 读回确认。',
        unprotect: '[Office.js 通道] manage_sheet(action="unprotect") 在 Microsoft 宿主未实现（同上）。替代路径：① 改用 host="wps" 的 wps_manage_sheet(action="unprotect")；② 在 Excel 里用「审阅 → 撤销工作表保护」手动解除。'
      };

      if (!normalized) {
        throw new Error(
          `[Office.js 通道] manage_sheet 无法识别的 action: "${rawAction}"。` +
          '支持 rename | move | tab_color | protect | unprotect（另兼容 copy/duplicate/activate/color/hide/show）。' +
          '原实现对未知 action 返回 success 但什么都不做，已改为显式报错。'
        );
      }
      if (unsupported[normalized]) throw new Error(unsupported[normalized]);

      if (normalized === "copy") {
        const copied = sheet.copy(Excel.WorksheetPositionType.after, sheet);
        const newName = params.newSheetName || params.newName;
        if (newName) copied.name = newName;
        copied.load("name, position");
        await context.sync();
        return { success: true, action: "copy", sheetName: copied.name, sourceSheetName: sheetName, position: copied.position };
      }

      if (normalized === "rename") {
        const newName = params.newName || params.newSheetName;
        if (!newName) {
          throw new Error('[Office.js 通道] manage_sheet(action="rename") 缺少 newName：重命名必须给出新名称，未执行任何修改。');
        }
        sheet.name = newName;
        sheet.load("name");
        await context.sync();
        return { success: true, action: "rename", sheetName: sheet.name, previousName: sheetName };
      }

      if (normalized === "move") {
        // targetIndex 从 1 开始；Office.js 的 Worksheet.position 从 0 开始。
        const targetIndex = readSheetNumber(params.targetIndex, params.position, params.index, params.targetPosition);
        if (targetIndex === null) {
          throw new Error(
            '[Office.js 通道] manage_sheet(action="move") 缺少 targetIndex：请给出目标位置序号（从 1 开始，1=最前）。' +
            '原实现会返回 success 但不移动工作表，已阻断。'
          );
        }
        const total = context.workbook.worksheets.getCount();
        await context.sync();
        if (targetIndex < 1 || targetIndex > total.value) {
          throw new Error(`[Office.js 通道] manage_sheet(action="move") 的 targetIndex=${targetIndex} 越界：当前工作簿共 ${total.value} 张工作表（合法范围 1..${total.value}）。`);
        }
        sheet.position = Math.round(targetIndex) - 1;
        sheet.load("name, position");
        await context.sync();
        return { success: true, action: "move", sheetName: sheet.name, position: sheet.position, targetIndex: Math.round(targetIndex) };
      }

      if (normalized === "tab_color") {
        const color = params.color || params.tabColor;
        if (!color) {
          throw new Error(
            '[Office.js 通道] manage_sheet(action="tab_color") 缺少 color：请给出十六进制标签底色（如 "#EF4444"）。' +
            '原实现只在 params.tabColor 存在时才设色，网关传的是 color → 静默 no-op，现已阻断。'
          );
        }
        sheet.tabColor = color;
        sheet.load("name, tabColor");
        await context.sync();
        return { success: true, action: "tab_color", sheetName: sheet.name, color: sheet.tabColor };
      }

      if (normalized === "hide" || normalized === "show") {
        sheet.visibility = normalized === "hide" ? Excel.SheetVisibility.hidden : Excel.SheetVisibility.visible;
        sheet.load("name, visibility");
        await context.sync();
        return { success: true, action: normalized, sheetName: sheet.name, visibility: sheet.visibility };
      }

      // 到这里说明别名表与分支不同步（防回归）。
      throw new Error(`[Office.js 通道] manage_sheet 内部错误：action "${rawAction}" 已归一为 "${normalized}"，但没有对应实现分支。`);
    });
  }

// ── 模块: src/excel/range.js — 区域读写、结构与查找替换 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 614-918 行（原样搬迁，未改写）。
  // 3. 区域核心读写
  async function handleReadRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      let range;
      if (params.address) {
        range = sheet.getRange(params.address);
      } else {
        range = sheet.getUsedRange();
      }

      range.load("address, rowCount, columnCount, values, formulas, text, numberFormat");
      await context.sync();

      if (params.metadata_only) {
        return {
          address: range.address,
          rowCount: range.rowCount,
          columnCount: range.columnCount
        };
      }

      return {
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values,
        formulas: range.formulas,
        text: range.text
      };
    });
  }

  async function handleGetRangeStyles(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);
      range.load("address, format/font/name, format/font/size, format/font/bold, format/font/color, format/fill/color, format/horizontalAlignment, format/verticalAlignment, numberFormat");
      await context.sync();

      return {
        success: true,
        address: range.address,
        font: {
          name: range.format.font.name,
          size: range.format.font.size,
          bold: range.format.font.bold,
          color: range.format.font.color
        },
        fill: {
          color: range.format.fill.color
        },
        alignment: {
          horizontal: range.format.horizontalAlignment,
          vertical: range.format.verticalAlignment
        },
        numberFormat: range.numberFormat ? range.numberFormat[0]?.[0] : null
      };
    });
  }

  async function handleWriteRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      let range;
      if (params.address) {
        range = sheet.getRange(params.address);
      } else {
        throw new Error("write_range 必须指定 address 参数 (如 'A1' 或 'B2:D5')");
      }

      let values = params.values || params.data;
      let formulas = params.formulas;

      // 标准化 values 为规范二维数组
      if (values !== undefined && values !== null) {
        if (!Array.isArray(values)) {
          values = [[values]];
        } else if (!Array.isArray(values[0])) {
          values = [values];
        }
      }

      // 标准化 formulas 为规范二维数组
      if (formulas !== undefined && formulas !== null) {
        if (!Array.isArray(formulas)) {
          formulas = [[formulas]];
        } else if (!Array.isArray(formulas[0])) {
          formulas = [formulas];
        }
      }

      const primaryMatrix = formulas || values;
      if (!primaryMatrix) {
        throw new Error("write_range 必须提供 values 或 formulas 参数");
      }

      const rowCount = primaryMatrix.length;
      const columnCount = primaryMatrix[0].length;

      // 如果给的是单一左上角单元格且提供了二维数组，自动按矩阵尺寸扩展 range
      if (!params.address.includes(":") && (rowCount > 1 || columnCount > 1)) {
        range = range.getResizedRange(rowCount - 1, columnCount - 1);
      }

      // 如果仅有 values，但 values 中包含以 '=' 开头的字符串，自动提取为公式矩阵
      if (values && !formulas) {
        const hasFormulasInValues = values.some(row => Array.isArray(row) && row.some(cell => typeof cell === 'string' && cell.startsWith('=')));
        if (hasFormulasInValues) {
          formulas = values.map(row => (Array.isArray(row) ? row.map(cell => (typeof cell === 'string' && cell.startsWith('=')) ? cell : null) : [null]));
        }
      }

      // 赋值 values (先写入所有常规数值与文字)
      if (values) {
        range.values = values;
      }

      // 赋值 formulas (仅针对实际包含公式的单元格精准注入，杜绝空字符串抹除数据)
      if (formulas) {
        for (let r = 0; r < rowCount; r++) {
          for (let c = 0; c < columnCount; c++) {
            const f = formulas[r]?.[c];
            if (f && typeof f === 'string' && f.startsWith('=')) {
              range.getCell(r, c).formulas = [[f]];
            }
          }
        }
      }

      // 关键：在 sync 前必须显式 load 需要返回的属性，绝不裸读！
      range.load("address");
      sheet.load("name");
      await context.sync();

      return {
        success: true,
        address: range.address || params.address,
        sheetName: sheet.name,
        rowCount: rowCount,
        columnCount: columnCount,
        modifiedCount: rowCount * columnCount
      };
    });
  }

  /** 列号 → 列字母（1 → A，27 → AA）。 */
  function columnIndexToLetters(colIdx) {
    let temp = Math.floor(colIdx);
    let letter = '';
    while (temp > 0) {
      const mod = (temp - 1) % 26;
      letter = String.fromCharCode(mod + 65) + letter;
      temp = Math.floor((temp - mod) / 26);
    }
    return letter;
  }

  /**
   * 维度归一（ISS-93）：网关用 WPS 语义 `targetType`（'row' | 'column'），
   * 早期 Office.js 实现读的是 `dimension`（'rows' | 'columns'）且**默认 rows** ——
   * 结果"想插列却插行"。这里两套字段名都认，且都识别不了时**抛错**，绝不默认成行。
   */
  function resolveDimension(params) {
    const raw = params.targetType !== undefined ? params.targetType : params.dimension;
    if (raw === undefined || raw === null || raw === '') {
      throw new Error(
        '[Office.js 通道] manage_rows_and_columns 缺少 targetType：必须显式给出 "row" 或 "column"。' +
        '原实现默认按行处理 → 传 "column" 时静默插行，已阻断。'
      );
    }
    const s = String(raw).trim().toLowerCase();
    if (['row', 'rows', '行', 'r'].includes(s)) return 'rows';
    if (['column', 'columns', 'col', 'cols', '列', 'c'].includes(s)) return 'columns';
    throw new Error(`[Office.js 通道] manage_rows_and_columns 无法识别的 targetType: "${raw}"（支持 "row" | "column"）。`);
  }

  /** index 归一：行用数字（1 基），列允许数字或列字母（'B'）。 */
  function resolveStartIndex(params, dimension) {
    const raw = params.index !== undefined ? params.index : params.startIndex;
    if (raw === undefined || raw === null || raw === '') {
      throw new Error('[Office.js 通道] manage_rows_and_columns 缺少 index：请给出起始行号或列号（数字，从 1 开始；列也接受字母如 "B"）。');
    }
    if (dimension === 'columns' && typeof raw === 'string' && /^[A-Za-z]{1,3}$/.test(raw.trim())) {
      const letters = raw.trim().toUpperCase();
      let idx = 0;
      for (const ch of letters) idx = idx * 26 + (ch.charCodeAt(0) - 64);
      return idx;
    }
    const n = Number(raw);
    if (!Number.isFinite(n) || n < 1) {
      throw new Error(`[Office.js 通道] manage_rows_and_columns 的 index="${raw}" 非法：需要 ≥1 的数字或列字母。`);
    }
    return Math.floor(n);
  }

  async function handleUpdateRangeStructure(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const rawAction = String(params.action || 'insert').trim().toLowerCase();
      const actionAliases = {
        insert: 'insert', add: 'insert',
        delete: 'delete', remove: 'delete',
        hide: 'hide',
        unhide: 'unhide', show: 'unhide',
        set_size: 'set_size', set_row_height: 'set_size', set_column_width: 'set_size'
      };
      const action = actionAliases[rawAction];
      if (!action) {
        throw new Error(
          `[Office.js 通道] manage_rows_and_columns 无法识别的 action: "${params.action}"。` +
          '支持 insert | delete | hide | unhide | set_size（原实现对未知 action 返回 success 但什么都不做，已改为显式报错）。'
        );
      }

      const dimension = resolveDimension(params);
      const count = Math.floor(Number(params.count ?? 1));
      if (!Number.isFinite(count) || count < 1) {
        throw new Error(`[Office.js 通道] manage_rows_and_columns 的 count=${params.count} 非法：需要 ≥1 的整数。`);
      }

      let range;
      let targetAddress;
      if (params.address) {
        targetAddress = params.address;
        range = sheet.getRange(targetAddress);
      } else {
        const start = resolveStartIndex(params, dimension);
        if (dimension === 'columns') {
          targetAddress = `${columnIndexToLetters(start)}:${columnIndexToLetters(start + count - 1)}`;
        } else {
          targetAddress = `${start}:${start + count - 1}`;
        }
        range = sheet.getRange(targetAddress);
      }

      if (action === 'set_size') {
        const size = Number(params.size);
        if (params.size === undefined || params.size === null || !Number.isFinite(size) || size <= 0) {
          throw new Error(
            `[Office.js 通道] manage_rows_and_columns(action="set_size") 缺少合法 size：` +
            `targetType='row' 时为磅值行高（如 24），'column' 时为字符列宽（如 15）。未执行任何修改。`
          );
        }
        if (dimension === 'columns') {
          range.format.columnWidth = size;
        } else {
          range.format.rowHeight = size;
        }
        await context.sync();

        // 写后读回：不把"请求成功"当成"尺寸已变"。
        range.load('format/rowHeight, format/columnWidth');
        await context.sync();
        const applied = dimension === 'columns' ? range.format.columnWidth : range.format.rowHeight;
        return {
          success: true,
          action,
          targetType: dimension === 'columns' ? 'column' : 'row',
          address: targetAddress,
          requestedSize: size,
          appliedSize: applied,
          warnings: Math.abs(Number(applied) - size) > 0.01 ? [`设置后读回 ${applied}，与请求 ${size} 不一致（宿主可能按内容或缩放换算）。`] : []
        };
      }

      if (action === 'hide' || action === 'unhide') {
        range.hidden = action === 'hide';
        await context.sync();
        range.load('hidden');
        await context.sync();
        return { success: true, action, targetType: dimension === 'columns' ? 'column' : 'row', address: targetAddress, hidden: range.hidden };
      }

      const shift = params.shift || (dimension === 'columns' ? 'Right' : 'Down');
      if (action === 'insert') {
        range.insert(shift);
      } else {
        range.delete(dimension === 'columns' ? 'Left' : 'Up');
      }
      // 关键：属性必须显式 load 再 sync 才能读；直接读未 load 的 sheet.name 会抛
      // 「属性"name"不可用」——且此时**写入已经执行**，调用方会误判为整体失败（实机踩到）。
      sheet.load('name');
      await context.sync();
      return {
        success: true,
        action,
        targetType: dimension === 'columns' ? 'column' : 'row',
        address: targetAddress,
        count,
        shift: action === 'insert' ? shift : undefined,
        sheetName: sheet.name
      };
    });
  }

  async function handleClearRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const targetRange = params.address || params.range;
      if (!targetRange) {
        throw new Error('[Office.js 通道] clear_range 缺少必要参数: address（如 "A1:E20"）。未执行任何修改。');
      }
      const range = sheet.getRange(targetRange);
      const applyTo = params.applyTo || "All";
      range.clear(applyTo);
      await context.sync();
      return { success: true, address: targetRange, cleared: applyTo };
    });
  }

  async function handleCopyRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const source = params.sourceAddress || params.sourceRange || params.from;
      const destination = params.destinationAddress || params.destinationRange || params.to;
      if (!source || !destination) {
        throw new Error('[Office.js 通道] copy_range 需要 sourceAddress 与 destinationAddress（别名 sourceRange/destinationRange）。未执行任何修改。');
      }
      const sourceRange = sheet.getRange(source);
      const destRange = sheet.getRange(destination);
      destRange.copyFrom(sourceRange, params.copyType || "All");
      await context.sync();
      return { success: true, sourceAddress: source, destinationAddress: destination, copyType: params.copyType || "All" };
    });
  }

  async function handleSetHyperlink(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const address = params.address || params.cell || params.range;
      if (!address || !params.url) {
        throw new Error('[Office.js 通道] set_hyperlink 需要 address 与 url。未执行任何修改。');
      }
      const range = sheet.getRange(address);
      range.hyperlink = {
        address: params.url,
        textToDisplay: params.textToDisplay || params.displayText || params.url,
        screenTip: params.screenTip || ""
      };
      await context.sync();
      return { success: true, address, url: params.url };
    });
  }

  /** WPS 语义 validationType → Office.js DataValidationRule 的规则键。 */
  const VALIDATION_TYPE_MAP = {
    list: 'list',
    number_range: 'wholeNumber',
    number: 'wholeNumber',
    whole_number: 'wholeNumber',
    integer: 'wholeNumber',
    decimal: 'decimal',
    date: 'date',
    time: 'time',
    text_length: 'textLength',
    textlength: 'textLength',
    custom: 'custom'
  };

  /** 网关 operator（WPS 语义）→ Excel.DataValidationOperator 的字符串值（用字符串以免旧宿主缺枚举）。 */
  const VALIDATION_OPERATOR_MAP = {
    between: 'Between',
    not_between: 'NotBetween',
    greater_than: 'GreaterThan',
    greater_than_or_equal: 'GreaterThanOrEqualTo',
    less_than: 'LessThan',
    less_than_or_equal: 'LessThanOrEqualTo',
    equal: 'EqualTo',
    equals: 'EqualTo',
    not_equal: 'NotEqualTo'
  };

  /** 数值/日期公式统一转 number；日期字符串保持字符串（ISO）。 */
  function toValidationFormula(value, ruleKey, label) {
    if (value === undefined || value === null || value === '') {
      throw new Error(`[Office.js 通道] set_data_validation 的 ${label} 缺失：当前验证类型需要该阈值。`);
    }
    if (ruleKey === 'date' || ruleKey === 'time') return String(value);
    const n = Number(value);
    if (!Number.isFinite(n)) {
      throw new Error(`[Office.js 通道] set_data_validation 的 ${label}="${value}" 不是合法数值。`);
    }
    return n;
  }

  /** 去掉 Excel 内联列表来源外层的成对引号（桥接侧可能已加引号）。 */
  function stripOuterQuotes(source) {
    const s = String(source).trim();
    return s.length >= 2 && s.startsWith('"') && s.endsWith('"') ? s.slice(1, -1) : s;
  }

  /**
   * 构造数据有效性规则（ISS-93-a）。
   *
   * 网关（WPS 语义）发的是 validationType / listItems / operator / minVal / maxVal / prompt* / error*；
   * 原 Office.js 实现只读 `params.rule` → 先 `clear()` 再什么都不设，**静默清空既有校验**。
   * 现在：两套字段名都认（WPS 语义优先），且**构造失败时在 clear() 之前抛错**，不清空任何东西。
   */
  function buildDataValidationRule(params) {
    // 兼容路径：调用方（含桥接侧 normalizer）直接给 Office.js 原生规则对象。
    if (params.rule !== undefined && params.rule !== null && !params.validationType) {
      if (typeof params.rule !== 'object') {
        throw new Error('[Office.js 通道] set_data_validation 的 rule 必须是 Office.js DataValidationRule 对象。');
      }
      const rule = { ...params.rule };
      if (rule.list && rule.list.source !== undefined) {
        rule.list = { ...rule.list, source: stripOuterQuotes(rule.list.source) };
      }
      return { rule, typeLabel: 'rule(原生对象)' };
    }

    const rawType = params.validationType !== undefined && params.validationType !== null && params.validationType !== ''
      ? String(params.validationType)
      : 'list';
    const ruleKey = VALIDATION_TYPE_MAP[rawType.trim().toLowerCase()];
    if (!ruleKey) {
      throw new Error(
        `[Office.js 通道] set_data_validation 无法识别的 validationType: "${rawType}"。` +
        '支持 list | number_range | decimal | date | time | text_length | custom（另兼容直接传 Office.js 原生 rule 对象）。'
      );
    }

    if (ruleKey === 'list') {
      const items = params.listItems ?? params.items ?? params.source;
      let source = null;
      if (Array.isArray(items)) {
        const flat = items.map(v => String(v)).filter(v => v.length > 0);
        if (flat.length === 0) {
          throw new Error('[Office.js 通道] set_data_validation(validationType="list") 的 listItems 为空数组：下拉列表至少需要一项。');
        }
        source = flat.join(',');
      } else if (typeof items === 'string' && items.trim()) {
        source = stripOuterQuotes(items);
      }
      if (!source) {
        throw new Error(
          '[Office.js 通道] set_data_validation 缺少 listItems：validationType="list" 必须给出候选项数组（如 ["已通过","待复测"]）。' +
          '原实现在缺少 rule 时会先 clear() 再什么都不设，静默清空既有校验，已阻断。'
        );
      }
      return { rule: { list: { inCellDropDown: true, source } }, typeLabel: 'list', source };
    }

    if (ruleKey === 'custom') {
      const formula = params.formula || params.customFormula || params.minVal;
      if (!formula || typeof formula !== 'string' || !formula.trim().startsWith('=')) {
        throw new Error('[Office.js 通道] set_data_validation(validationType="custom") 需要 formula，且必须以 "=" 开头（如 "=ISNUMBER(A1)"）。');
      }
      return { rule: { custom: { formula: formula.trim() } }, typeLabel: 'custom' };
    }

    const opRaw = params.operator !== undefined && params.operator !== null && params.operator !== '' ? String(params.operator) : 'between';
    const operator = VALIDATION_OPERATOR_MAP[opRaw.trim().toLowerCase()];
    if (!operator) {
      throw new Error(
        `[Office.js 通道] set_data_validation 无法识别的 operator: "${opRaw}"。` +
        '支持 between | not_between | greater_than | greater_than_or_equal | less_than | less_than_or_equal | equal | not_equal。'
      );
    }

    const isTernary = operator === 'Between' || operator === 'NotBetween';
    const primary = isTernary ? (params.minVal ?? params.formula1) : (params.minVal ?? params.maxVal ?? params.formula1);
    const label = isTernary ? 'minVal' : (params.minVal !== undefined ? 'minVal' : 'maxVal');
    const body = { operator, formula1: toValidationFormula(primary, ruleKey, label) };
    if (isTernary) {
      body.formula2 = toValidationFormula(params.maxVal ?? params.formula2, ruleKey, 'maxVal');
    }

    const rule = {};
    rule[ruleKey] = body;
    return { rule, typeLabel: ruleKey };
  }

  async function handleSetDataValidation(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const address = params.address || params.range;
      if (!address) {
        throw new Error('[Office.js 通道] set_data_validation 缺少必要参数: address（如 "E5:E50"）。未执行任何修改。');
      }

      // 关键顺序（ISS-93-a）：先构造规则，构造失败就抛错 —— 此时**还没有 clear()**，既有校验不被破坏。
      const built = buildDataValidationRule(params);
      const range = sheet.getRange(address);

      range.dataValidation.clear();
      range.dataValidation.rule = built.rule;
      if (params.ignoreBlanks !== undefined) range.dataValidation.ignoreBlanks = Boolean(params.ignoreBlanks);

      // 提示与报错文案：桥接侧 normalizer 会把它们列在 unsupportedFields 里但**不下发值**，
      // 这里两处都收：顶层参数（WPS 语义）+ unsupportedFields 里列出的同名字段。
      const extraListed = Array.isArray(params.unsupportedFields) ? params.unsupportedFields : [];
      const hasValue = (n) => params[n] !== undefined && params[n] !== null && String(params[n]) !== '';
      const picked = (...names) => {
        for (const n of names) if (hasValue(n)) return String(params[n]);
        return null;
      };
      const promptTitle = picked('promptTitle');
      const promptMessage = picked('promptMessage');
      if (promptTitle || promptMessage) {
        range.dataValidation.prompt = { showPrompt: true, title: promptTitle || '', message: promptMessage || '' };
      }
      const errorTitle = picked('errorTitle');
      const errorMessage = picked('errorMessage');
      if (errorTitle || errorMessage) {
        range.dataValidation.errorAlert = {
          showAlert: true,
          style: 'Stop',
          title: errorTitle || '',
          message: errorMessage || ''
        };
      }
      const auxiliaryCandidateFields = ['promptTitle', 'promptMessage', 'errorTitle', 'errorMessage'];
      const auxiliaryApplied = auxiliaryCandidateFields.filter(hasValue);
      const auxiliaryNotApplied = extraListed.filter(f => auxiliaryCandidateFields.includes(f) && !hasValue(f));

      await context.sync();

      // 写后读回：不把"请求成功"当成"校验已设上"。
      range.load('address, dataValidation/type, dataValidation/rule, dataValidation/prompt, dataValidation/errorAlert');
      sheet.load('name');
      await context.sync();

      const dv = range.dataValidation;
      const readType = dv ? dv.type : null;
      const warnings = [];
      if (!readType || String(readType).toLowerCase() === 'none') {
        warnings.push(
          `设置后读回 dataValidation.type=${readType || 'null'}：宿主没有保留该规则，` +
          '请用 read_range/execute_script 复核，不要直接重试。'
        );
      }
      if (built.typeLabel === 'list' && built.source && dv && dv.rule && dv.rule.list && dv.rule.list.source) {
        const readSource = String(dv.rule.list.source);
        if (!readSource.includes(built.source.replace(/^=/, '').split(',')[0])) {
          warnings.push(`读回的下拉来源为 "${readSource}"，与请求 "${built.source}" 不一致（Excel 可能把逗号列表改写为区域引用）。`);
        }
      }

      return {
        success: true,
        address: range.address || address,
        sheetName: sheet.name,
        validationType: built.typeLabel,
        rule: dv ? dv.rule : null,
        readBackType: readType,
        prompt: dv ? dv.prompt : null,
        errorAlert: dv ? dv.errorAlert : null,
        auxiliaryFieldsApplied: auxiliaryApplied,
        auxiliaryFieldsNotApplied: auxiliaryNotApplied,
        warnings
      };
    });
  }

  async function handleFindReplace(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      // 网关（WPS 语义）传的是 searchRange；Office.js 侧原来只认 address（ISS-93-c）。
      const searchRange = params.searchRange || params.address || params.range;
      const range = searchRange ? sheet.getRange(searchRange) : sheet.getUsedRange();
      // 字段名兼容：网关（WPS 语义）传的是 searchQuery，Office.js 侧原来只认 text/query/findText。
      // 名字对不上 → text 落成空串 → `replaceAll("", …)` 会命中整片区域并可能破坏内容（问题台账 ISS-93）。
      const text = params.text || params.query || params.findText || params.searchQuery || "";
      if (!text) {
        throw new Error(
          "find_and_replace 缺少搜索文本（Office.js 通道识别 text / query / findText / searchQuery，当前都为空）。" +
          "空搜索串会命中整个区域并可能破坏内容，已阻断；请显式提供要查找的文本。"
        );
      }
      const replaceText = params.replaceText;
      const matchCase = !!params.matchCase;
      const matchEntireCell = !!params.matchEntireCell;
      // maxResults（网关默认 50；桥接侧 normalizer 也会做一次截断）：<=0 视为不限量。
      const rawMax = Number(params.maxResults);
      const maxResults = Number.isFinite(rawMax) && rawMax > 0 ? Math.floor(rawMax) : Infinity;

      if (replaceText !== undefined) {
        // 替换前先确认范围可读，替换后回读文本以给出可核验的读数（不谎报命中数）。
        range.load("text, address");
        await context.sync();
        const beforeText = JSON.stringify(range.text || []);
        range.replaceAll(text, replaceText, { completeMatch: matchEntireCell, matchCase: matchCase });
        await context.sync();
        range.load("text");
        await context.sync();
        const afterText = JSON.stringify(range.text || []);
        return {
          success: true,
          address: range.address,
          searchRange: searchRange || "(usedRange)",
          replacedWith: replaceText,
          changed: beforeText !== afterText,
          textChanged: beforeText !== afterText
        };
      }

      // 跨平台安全搜索：macOS Office.js 缺少 range.findAll，采用纯 JS 内存矩阵检索
      range.load("values, text, rowIndex, columnIndex, address");
      await context.sync();

      function colToLetters(c) {
        let temp = c + 1;
        let letter = "";
        while (temp > 0) {
          const mod = (temp - 1) % 26;
          letter = String.fromCharCode(65 + mod) + letter;
          temp = Math.floor((temp - mod) / 26);
        }
        return letter;
      }

      const baseRow = range.rowIndex || 0;
      const baseCol = range.columnIndex || 0;
      const matches = [];
      const queryStr = matchCase ? String(text) : String(text).toLowerCase();
      const vals = range.text || range.values || [];
      let truncated = false;

      for (let r = 0; r < vals.length; r++) {
        const row = vals[r] || [];
        for (let c = 0; c < row.length; c++) {
          const cellVal = row[c];
          if (cellVal === null || cellVal === undefined) continue;
          const str = String(cellVal);
          const targetStr = matchCase ? str : str.toLowerCase();
          const matched = matchEntireCell ? targetStr === queryStr : targetStr.includes(queryStr);
          if (matched) {
            if (matches.length >= maxResults) { truncated = true; break; }
            const cellAddr = `${colToLetters(baseCol + c)}${baseRow + r + 1}`;
            matches.push({
              address: cellAddr,
              text: str,
              row: baseRow + r + 1,
              column: baseCol + c + 1
            });
          }
        }
        if (truncated) break;
      }

      return {
        success: true,
        address: range.address,
        searchRange: searchRange || "(usedRange)",
        count: matches.length,
        truncated,
        maxResults: maxResults === Infinity ? null : maxResults,
        matches: matches
      };
    });
  }

// ── 模块: src/excel/formula.js — 公式写入 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 919-936 行（原样搬迁，未改写）。
  // 4. 公式
  async function handleSetFormula(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);
      let formula = params.formula;

      if (Array.isArray(formula)) {
        range.formulas = formula;
      } else {
        if (!formula.startsWith("=")) formula = "=" + formula;
        range.formulas = [[formula]];
      }
      await context.sync();
      return { success: true, address: params.address, formula };
    });
  }

// ── 模块: src/excel/format.js — 格式排版、冻结窗格与条件格式 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 937-1136 行（原样搬迁，未改写）。
  // 5. 格式与排版
  async function handleFormatRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const targetAddress = params.address || params.range || params.cellRange;
      if (!targetAddress) throw new Error("缺少必要参数: address 或 range");
      const range = sheet.getRange(targetAddress);

      // 0. 企业级命名样式模板 (Named Style Templates)
      const styleTemplates = {
        "OPPO_Banner": { fillColor: "#046A38", fontColor: "#FFFFFF", bold: true, fontSize: 14, horizontalAlignment: "Center" },
        "OPPO_Subtitle": { fillColor: "#046A38", fontColor: "#E2ECE6", fontSize: 10, horizontalAlignment: "Center" },
        "OPPO_Card_Header": { fillColor: "#046A38", fontColor: "#FFFFFF", bold: true, fontSize: 11 },
        "OPPO_Table_Header": { fillColor: "#EAF6EE", fontColor: "#046A38", bold: true, fontSize: 10 },
        "OPPO_Total_Row": { fillColor: "#EAF6EE", fontColor: "#046A38", bold: true },
        "OPPO_KPI_Value": { fillColor: "#FFFFFF", fontColor: "#046A38", bold: true, fontSize: 18, horizontalAlignment: "Center" },
        "OPPO_KPI_Sub": { fillColor: "#FFFFFF", fontColor: "#10B981", fontSize: 9, horizontalAlignment: "Center" },
        "OPPO_Canvas_Bg": { fillColor: "#F4F6F8" },
        "OPPO_Card_Body": { fillColor: "#FFFFFF" }
      };

      const preset = params.style ? styleTemplates[params.style] : null;
      if (preset) {
        if (preset.fillColor) range.format.fill.color = preset.fillColor;
        if (preset.fontColor) range.format.font.color = preset.fontColor;
        if (preset.bold !== undefined) range.format.font.bold = preset.bold;
        if (preset.fontSize) range.format.font.size = preset.fontSize;
        if (preset.horizontalAlignment) range.format.horizontalAlignment = preset.horizontalAlignment;
      }

      // 1. 字体样式 (兼容扁平入参与嵌套入参)
      const fontName = params.fontName || params.font?.name;
      if (fontName) range.format.font.name = fontName;

      const fontSize = params.fontSize || params.size || params.font?.size;
      if (fontSize) range.format.font.size = fontSize;

      const fontColor = params.fontColor || params.color || params.textColor || params.font?.color;
      if (fontColor) range.format.font.color = fontColor;

      const bold = params.bold !== undefined ? params.bold : params.font?.bold;
      if (bold !== undefined) range.format.font.bold = bold;

      const italic = params.italic !== undefined ? params.italic : params.font?.italic;
      if (italic !== undefined) range.format.font.italic = italic;

      // 2. 单元格背景填充色 (兼容 backgroundColor, fillColor, bg, fill.color)
      const bgColor = params.backgroundColor || params.fillColor || params.bg || params.fill?.color;
      if (bgColor) {
        range.format.fill.color = bgColor;
      }

      // 3. 对齐方式
      const hAlign = params.horizontalAlignment || params.alignment?.horizontal;
      if (hAlign) {
        const hMap = {
          left: "Left", center: "Center", right: "Right",
          justify: "Justify", general: "General"
        };
        range.format.horizontalAlignment = hMap[String(hAlign).toLowerCase()] || hAlign;
      }

      const vAlign = params.verticalAlignment || params.alignment?.vertical;
      if (vAlign) {
        const vMap = {
          top: "Top", center: "Center", bottom: "Bottom",
          distributed: "Distributed", justify: "Justify"
        };
        range.format.verticalAlignment = vMap[String(vAlign).toLowerCase()] || vAlign;
      }

      const wrapText = params.wrapText !== undefined ? params.wrapText : params.alignment?.wrapText;
      if (wrapText !== undefined) {
        range.format.wrapText = wrapText;
      }

      // 4. 数字/百分比/千分位/货币/日期格式
      if (params.numberFormat) {
        range.numberFormat = [[params.numberFormat]];
      }

      // 5. 边框 (borders)
      if (params.borders) {
        const borderColor = params.borderColor || "#D9D9D9";
        const borderStyle = typeof params.borders === 'string' && ['Continuous', 'Dash', 'Dot', 'Double'].includes(params.borders)
          ? params.borders : 'Continuous';
        const edges = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideHorizontal', 'InsideVertical'];
        for (const edge of edges) {
          try {
            const b = range.format.borders.getItem(edge);
            b.style = borderStyle;
            b.color = borderColor;
          } catch {}
        }
      }

      // 6. 行高与列宽
      if (params.rowHeight) range.rowHeight = params.rowHeight;
      if (params.columnWidth) range.columnWidth = params.columnWidth;

      // 7. 合并 / 拆分
      if (params.merge) range.merge(params.across || false);
      if (params.unmerge) range.unmerge();

      range.load("address");
      sheet.load("name");
      await context.sync();
      return { success: true, address: range.address || params.address, sheetName: sheet.name };
    });
  }

  async function handleAutofitColumns(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params?.sheetName);
      const range = params?.address ? sheet.getRange(params.address) : sheet.getUsedRange();
      range.format.autofitColumns();
      if (params?.autofitRows) range.format.autofitRows();
      await context.sync();
      return { success: true };
    });
  }

  async function handleFreezePanes(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params?.sheetName);
      if (params?.unfreeze) {
        sheet.freezePanes.unfreeze();
        await context.sync();
        return { success: true, message: "已解除冻结窗格" };
      }

      // 避免 Office.js 重复冻结报 InternalError: An internal error has occurred，先重置解冻
      sheet.freezePanes.unfreeze();

      const r = Number(params?.freezeRowIndex || 0);
      const c = Number(params?.freezeColumnIndex || 0);
      if (r > 0 && c > 0) {
        sheet.freezePanes.freezeAt(sheet.getCell(r, c));
      } else if (r > 0) {
        sheet.freezePanes.freezeRows(r);
      } else if (c > 0) {
        sheet.freezePanes.freezeColumns(c);
      } else {
        sheet.freezePanes.freezeRows(1);
      }
      await context.sync();
      return { success: true, message: `已冻结窗格 (行: ${r}, 列: ${c})` };
    });
  }

  async function handleAddConditionalFormatting(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);

      if (params.clearExisting) {
        range.conditionalFormats.clearAll();
      }

      const ruleType = params.ruleType || params.type || "cell_value";
      if (ruleType === "cell_value" || ruleType === "cellValue") {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        const opMap = {
          less_than: Excel.ConditionalCellValueOperator.lessThan,
          greater_than: Excel.ConditionalCellValueOperator.greaterThan,
          equal: Excel.ConditionalCellValueOperator.equalTo,
          not_equal: Excel.ConditionalCellValueOperator.notEqualTo,
          between: Excel.ConditionalCellValueOperator.between
        };
        const op = opMap[params.operator] || Excel.ConditionalCellValueOperator.lessThan;
        cf.cellValue.operator = op;
        if (params.formula1 !== undefined) cf.cellValue.formula1 = String(params.formula1);
        if (params.formula2 !== undefined) cf.cellValue.formula2 = String(params.formula2);
        if (params.backgroundColor) cf.cellValue.format.fill.color = params.backgroundColor;
        if (params.fontColor) cf.cellValue.format.font.color = params.fontColor;
      } else if (ruleType === "color_scale" || ruleType === "colorScale") {
        range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
      } else if (ruleType === "data_bar" || ruleType === "dataBar") {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
        const color = params.barColor || params.color || "#00C05E";
        try {
          if (cf.dataBar.positiveFormat) {
            cf.dataBar.positiveFormat.fillColor = color;
            if (params.borderColor) {
              cf.dataBar.positiveFormat.borderColor = params.borderColor;
            }
            if (params.gradientFill !== undefined) {
              cf.dataBar.positiveFormat.gradientFill = Boolean(params.gradientFill);
            }
          }
        } catch (dbErr) {
          console.warn("设置数据条颜色警告:", dbErr);
        }
      }

      await context.sync();
      return { success: true, message: "已应用条件格式" };
    });
  }

// ── 模块: src/excel/filter-sort.js — 排序与筛选 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1137-1168 行（原样搬迁，未改写）。
  // 6. 排序与筛选
  async function handleSetFilterAndSort(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);
      const action = params.action || (params.filterColumn !== undefined ? "filter" : "sort");

      if (action === "filter") {
        const col = params.filterColumn !== undefined ? params.filterColumn : 0;
        const criteria = params.criteria ? (Array.isArray(params.criteria) ? params.criteria : [String(params.criteria)]) : [];
        sheet.autoFilter.apply(range, col, {
          filterOn: Excel.FilterOn.values,
          values: criteria
        });
      } else if (action === "clear") {
        sheet.autoFilter.remove();
        range.sort.clear();
      } else {
        const col = params.sortColumn !== undefined ? params.sortColumn : (params.key || 0);
        const ascending = params.sortOrder === "desc" ? false : (params.ascending !== false);
        range.sort.apply([{
          key: col,
          ascending: ascending,
          sortOn: Excel.SortOn.value
        }], params.hasHeader !== false);
      }

      await context.sync();
      return { success: true, action, address: params.address };
    });
  }

// ── 模块: src/excel/table.js — 结构化表格 Table ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1169-1194 行（原样搬迁，未改写）。
  // 7. 结构化表格
  async function handleCreateTable(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);
      const table = sheet.tables.add(range, params.hasHeaders !== false);
      if (params.name) table.name = params.name;
      if (params.style) table.style = params.style;
      table.load("name, id");
      await context.sync();
      return { success: true, name: table.name, id: table.id };
    });
  }

  async function handleUpdateTable(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const table = sheet.tables.getItem(params.name || params.tableName);
      if (params.style) table.style = params.style;
      if (params.showTotals !== undefined) table.showTotals = params.showTotals;
      if (params.resizeRange) table.resize(sheet.getRange(params.resizeRange));
      await context.sync();
      return { success: true };
    });
  }

// ── 模块: src/excel/chart.js — 图表 Chart 与工作表预览渲染 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1195-1662 行（原样搬迁，未改写）。
  // 8. 图表

  // ── ISS-96（高，伪渲染）───────────────
  // 旧实现的 capture_sheet_preview 在"无原生图表"时用 Canvas 按 cellW=110 / rowH=28 硬编码
  // 合成一张"看起来像表格"的图，并返回 success:true + "高保真渲染图" —— AI 的视觉自查会据此
  // 得出与真实文件不符的结论。**这比报错更危险：报错会让人停下来，假图不会。**
  // 现在的契约只有两条：
  //   ① 真实渲染路径：只有 Office.js 的 chart.getImage()（宿主原生导出）算真实渲染；
  //   ② 其余一律**抛错**，并在错误里写清替代路径（host=wps 的截图能力 / 在 Excel 里自行截图）。
  // 绝不再返回任何合成图、示意图或"已完成排版自检"之类无法核验的措辞。
  const CHART_NAME_PATTERN = /^(?:chart|图表|图)\s*(\d+)$/i;

  /** 图表名归一：把 "图表 2" / "chart2" 与宿主返回的 "Chart 2" 视为同一个。 */
  function normalizeChartName(raw) {
    const s = String(raw === undefined || raw === null ? '' : raw).trim().toLowerCase().replace(/\s+/g, '');
    if (!s) return '';
    const m = CHART_NAME_PATTERN.exec(s);
    if (m) return `chart${m[1]}`;
    return s;
  }

  /** 按名称（兼容 "Chart 2"/"图表 2"）或 id 查找图表。 */
  function findChartByName(items, wanted) {
    const target = normalizeChartName(wanted);
    if (!target) return null;
    return (items || []).find(c => normalizeChartName(c.name) === target || String(c.id || '').toLowerCase() === String(wanted).trim().toLowerCase()) || null;
  }

  function chartNamesOf(items) {
    return (items || []).map(c => {
      const title = c.title && c.title.text ? `「${c.title.text}」` : '';
      return `${c.name}${title}`;
    });
  }

  /** 渲染目标与尺寸归一，任一非法值都抛错（不再静默用默认值顶替用户的请求）。 */
  function readRenderSize(params) {
    const rawW = params.width === undefined || params.width === null || params.width === '' ? 800 : Number(params.width);
    const rawH = params.height === undefined || params.height === null || params.height === '' ? 450 : Number(params.height);
    if (!Number.isFinite(rawW) || !Number.isFinite(rawH) || rawW < 100 || rawH < 100 || rawW > 4000 || rawH > 4000) {
      throw new Error(`[Office.js 通道] capture_sheet_preview 的 width/height 非法（收到 ${params.width}×${params.height}）：需为 100..4000 的像素值。`);
    }
    return { width: Math.round(rawW), height: Math.round(rawH) };
  }

  async function handleGetCharts(params) {
    // CAP-08 通道：用已注册的 get_charts 作为"形状能力自检"的触发口。
    // 本机 MCP 工具面里没有任何入口能把任意 Office.js 代码送进任务窗格，
    // 形状方法在桥接路由表登记完成前也调不到；这个开关让**现有工具**就能取到真机读数。
    // 只在显式传 shapeSelfTest 时才跑，正常读图表的行为不变。
    if (params && params.shapeSelfTest === true) {
      const selfTest = await handleShapeSelfTest({ sheetName: params.selfTestSheetName, keep: params.selfTestKeep === true });
      return { success: selfTest.success, sheetName: selfTest.sheetName, count: 0, totalCount: 0, filtered: false, selector: null, charts: [], shapeSelfTest: selfTest, addon: { version: ADDON_VERSION, capabilities: ADDON_CAPABILITIES } };
    }
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const isDetail = !!params.detail;
      const charts = sheet.charts.load(
        isDetail
          ? "items/name, items/id, items/title/text, items/chartType, items/top, items/left, items/width, items/height, items/legend/visible, items/series/items/name"
          : "items/name, items/id, items/title/text, items/chartType, items/top, items/left, items/width, items/height"
      );
      // 错误分支里要用 sheet.name 组文案：必须一起 load，否则报「属性"name"不可用」而盖掉真正的错误。
      sheet.load("name");
      await context.sync();

      const all = charts.items.map((c, idx) => {
        const item = {
          chartIndex: idx + 1,
          name: c.name,
          shapeName: c.name,
          id: c.id,
          title: c.title ? c.title.text : "",
          chartType: c.chartType,
          top: c.top,
          left: c.left,
          width: c.width,
          height: c.height,
          hasLegend: isDetail && c.legend ? !!c.legend.visible : true,
          seriesCount: isDetail && c.series && c.series.items ? c.series.items.length : 1
        };
        if (isDetail && c.series && c.series.items) {
          item.series = c.series.items.map((s, sIdx) => ({
            seriesIndex: sIdx + 1,
            name: s.name || `系列 ${sIdx + 1}`
          }));
        }
        return item;
      });

      // ISS-93-d：网关传 shapeName / chartIndex / chartTitle 三个选择器，旧实现全部忽略、恒返回全表。
      // 现在按选择器过滤；选择器命中 0 张就**抛错并列出实际图表**，而不是悄悄返回别的图。
      const wantedName = params.shapeName || params.name || params.chartName;
      const wantedTitle = params.chartTitle;
      const wantedIndex = params.chartIndex === undefined || params.chartIndex === null || params.chartIndex === ''
        ? null
        : Number(params.chartIndex);
      if (wantedIndex !== null && (!Number.isFinite(wantedIndex) || wantedIndex < 1)) {
        throw new Error(`[Office.js 通道] get_charts 的 chartIndex=${params.chartIndex} 非法：需要 ≥1 的整数。`);
      }

      const hasSelector = Boolean(wantedName) || Boolean(wantedTitle) || wantedIndex !== null;
      let result = all;
      const selector = { shapeName: wantedName || null, chartTitle: wantedTitle || null, chartIndex: wantedIndex };

      if (hasSelector) {
        if (all.length === 0) {
          throw new Error(`[Office.js 通道] get_charts：工作表 [${sheet.name}] 中没有任何原生图表，选择器 ${JSON.stringify(selector)} 无对象可匹配。`);
        }
        if (wantedName) {
          const nameNorm = normalizeChartName(wantedName);
          result = result.filter(c => normalizeChartName(c.name) === nameNorm || normalizeChartName(c.id) === nameNorm);
        }
        if (wantedTitle) {
          const t = String(wantedTitle).toLowerCase();
          result = result.filter(c => String(c.title || '').toLowerCase().includes(t));
        }
        if (wantedIndex !== null) {
          result = result.filter(c => c.chartIndex === wantedIndex);
        }
        if (result.length === 0) {
          throw new Error(
            `[Office.js 通道] get_charts 选择器 ${JSON.stringify(selector)} 在 [${sheet.name}] 未匹配到任何图表（未返回其它图表充数）。` +
            `当前工作表实际有 ${all.length} 张：${chartNamesOf(all).join('、')}。`
          );
        }
      }

      return {
        success: true,
        sheetName: sheet.name,
        count: result.length,
        totalCount: all.length,
        filtered: hasSelector,
        selector: hasSelector ? selector : null,
        charts: result,
        // 部署探针（CAP-08）：让调用方一眼看出窗格里跑的是哪份代码与哪几条形状分支。
        // 桥接层 get_charts 的响应转换是"先铺开宿主原始响应再覆盖统一字段"（ISS-99），
        // 因此这个字段能原样到达调用方，不会被白名单丢掉。
        addon: { version: ADDON_VERSION, capabilities: ADDON_CAPABILITIES }
      };
    });
  }

  function applyChartPosition(chart, sheet, params, targetLeft, targetTop, targetWidth, targetHeight) {
    const targetCellRange = params.cellRange || params.position?.cellRange;
    const targetStartCell = params.startCell || params.position?.startCell || params.leftCell || params.position?.leftCell;
    const targetEndCell = params.endCell || params.position?.endCell;

    if (targetCellRange && targetCellRange.includes(":")) {
      const parts = targetCellRange.split(":").map(x => x.trim());
      try {
        chart.setPosition(sheet.getRange(parts[0]), sheet.getRange(parts[1]));
        return "cellRange";
      } catch (e1) {
        try { chart.setPosition(parts[0], parts[1]); return "cellRange"; } catch (e2) {}
      }
    } else if (targetStartCell && targetEndCell) {
      try {
        chart.setPosition(sheet.getRange(targetStartCell), sheet.getRange(targetEndCell));
        return "start/endCell";
      } catch (e1) {
        try { chart.setPosition(targetStartCell, targetEndCell); return "start/endCell"; } catch (e2) {}
      }
    } else if (targetStartCell) {
      try {
        chart.setPosition(sheet.getRange(targetStartCell));
        return "startCell";
      } catch (e1) {
        try { chart.setPosition(targetStartCell); return "startCell"; } catch (e2) {}
      }
    }
    if (targetLeft !== undefined) chart.left = targetLeft;
    if (targetTop !== undefined) chart.top = targetTop;
    if (targetWidth !== undefined) chart.width = targetWidth;
    if (targetHeight !== undefined) chart.height = targetHeight;
    return "pixels";
  }

  async function handleCreateChart(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const sourceRange = params.dataRange || params.sourceAddress;
      const extraRanges = Array.isArray(params.dataRanges) ? params.dataRanges.filter(r => typeof r === 'string' && r.trim()) : [];
      if (!sourceRange && extraRanges.length === 0) {
        throw new Error('[Office.js 通道] add_chart 缺少数据源：请提供 dataRange（如 "A4:E19"）或 dataRanges（多段区域数组）。');
      }
      const source = sourceRange ? sheet.getRange(sourceRange) : sheet.getRange(extraRanges[0]);

      const targetTitle = params.title || "";
      const targetLeft = params.left !== undefined ? Number(params.left) : (params.position?.left ? Number(params.position.left) : 350);
      const targetTop = params.top !== undefined ? Number(params.top) : (params.position?.top ? Number(params.position.top) : 20);
      const targetWidth = params.width !== undefined ? Number(params.width) : (params.position?.width ? Number(params.position.width) : 480);
      const targetHeight = params.height !== undefined ? Number(params.height) : (params.position?.height ? Number(params.position.height) : 280);

      // 处理 replaceExisting：若开启，先清理重叠位置或同名旧图，彻底避免图表堆叠
      const replaceExisting = params.replaceExisting !== false;
      const warnings = [];
      const removedCharts = [];
      if (replaceExisting) {
        try {
          const existingCharts = sheet.charts.load("items/name, items/id, items/title/text, items/left, items/top");
          await context.sync();
          for (const c of existingCharts.items) {
            const titleMatch = targetTitle && c.title && c.title.text === targetTitle;
            const posMatch = Math.abs(c.left - targetLeft) < 40 && Math.abs(c.top - targetTop) < 40;
            if (titleMatch || posMatch) {
              removedCharts.push(c.name);
              c.delete();
            }
          }
          await context.sync();
        } catch (cleanErr) {
          warnings.push(`清理已有图表时出错（未阻断建图）：${cleanErr && cleanErr.message ? cleanErr.message : cleanErr}`);
        }
      }

      const typeMap = {
        column: "ColumnClustered",
        column_clustered: "ColumnClustered",
        columnclustered: "ColumnClustered",
        clustered_column: "ColumnClustered",
        bar: "BarClustered",
        bar_clustered: "BarClustered",
        barclustered: "BarClustered",
        clustered_bar: "BarClustered",
        line: "Line",
        line_marker: "LineMarkers",
        pie: "Pie",
        doughnut: "Doughnut",
        donut: "Doughnut",
        area: "Area",
        scatter: "Scatter",
        xy_scatter: "Scatter"
      };
      const rawType = String(params.chartType || "ColumnClustered").trim().toLowerCase().replace(/[-\s]/g, '_');
      // ISS-17 同类问题的防线：枚举外的类型**报错**，不静默降级。
      // pareto 在 Office.js 没有对应 ChartType —— 明确拒绝并给出可执行替代，绝不悄悄建成柱状图。
      if (rawType === 'pareto') {
        throw new Error(
          '[Office.js 通道] add_chart 不支持 chartType="pareto"：Office.js 没有 Pareto 图表类型，' +
          '不会静默降级成柱状图。替代路径：① 用 chartType="column_clustered" 建柱状图，另加一列累计占比系列（或改用 host="wps" 的 add_chart）；' +
          '② 需要柏拉图外观时，自行对数据降序排序后再建图。'
        );
      }
      const chartType = typeMap[rawType];
      if (!chartType) {
        throw new Error(
          `[Office.js 通道] add_chart 无法识别的 chartType: "${params.chartType}"。` +
          '支持 line | column | column_clustered | bar | bar_clustered | pie | doughnut | area | scatter。'
        );
      }

      const chart = sheet.charts.add(chartType, source, params.seriesBy || "Auto");
      const positionMode = applyChartPosition(chart, sheet, params, targetLeft, targetTop, targetWidth, targetHeight);

      if (targetTitle) chart.title.text = targetTitle;

      let legendApplied = null;
      if (params.hasLegend !== undefined) {
        chart.legend.visible = Boolean(params.hasLegend);
        legendApplied = Boolean(params.hasLegend);
      }

      // ── ISS-93-e：dataRanges（多段数据源）────────────────
      // 诚实契约：Office.js 的 `chart.setData(range)` 是**整体替换**图表数据，不是"追加系列"
      // （实机验证：连调两次 setData 后读回仍是 1 个系列，第一段被第二段顶掉）。
      // 因此多段时只做一次 setData 覆盖，并在 warnings 里如实说明"多段未合并"，绝不谎报多系列。
      let dataRangesApplied = [sourceRange || extraRanges[0]];
      if (extraRanges.length > 0) {
        dataRangesApplied = extraRanges.slice();
        try {
          chart.setData(sheet.getRange(extraRanges[0]), params.seriesBy || "Auto");
          await context.sync();
          if (extraRanges.length > 1) {
            warnings.push(
              `dataRanges 共 ${extraRanges.length} 段，但 Office.js 的 chart.setData 不支持合并多段不连续区域：` +
              `本次只取第 1 段 "${extraRanges[0]}" 作为图表数据源（其余 ${extraRanges.slice(1).join('、')} 未进入图表）。` +
              '需要多系列请把数据整理成一段连续区域，或改用 host="wps" 的 add_chart。'
            );
          }
        } catch (multiErr) {
          warnings.push(
            `dataRanges 数据源设置失败：${multiErr && multiErr.message ? multiErr.message : multiErr}。` +
            '可改用一段连续区域，或改用 host="wps" 的 add_chart。'
          );
        }
      }

      const seriesList = chart.series.load("items");
      await context.sync();
      const seriesItems = seriesList.items || [];
      if (dataRangesApplied.length > 1 && seriesItems.length < dataRangesApplied.length) {
        warnings.push(`请求 dataRanges 共 ${dataRangesApplied.length} 段，宿主读回只有 ${seriesItems.length} 个系列（多段数据源未合并，见上条）。`);
      }

      // ── 系列着色 ────────────────────────────────
      let colors = params.seriesColors;
      if (typeof colors === "string") colors = [colors];
      if (!colors && params.seriesColor) colors = [params.seriesColor];
      if (!colors && params.color) colors = [params.color];
      if (!Array.isArray(colors)) colors = null;

      const isPieLike = rawType.includes("pie") || rawType.includes("doughnut") || rawType.includes("donut");
      try {
        if (isPieLike) {
          if (seriesItems.length > 0) {
            const points = seriesItems[0].points.load("items");
            await context.sync();
            if (colors && colors.length > 0) {
              // 饼/环图的"系列颜色"实际是逐点颜色：只给一个颜色就整圈同色，给 N 个就按点顺序取用。
              for (let pIdx = 0; pIdx < points.items.length; pIdx++) {
                points.items[pIdx].format.fill.setSolidColor(colors[pIdx % colors.length]);
              }
            } else {
              const piePalette = ["#046A38", "#00A854", "#2CFF73", "#52C41A", "#A3D4B6", "#145A32", "#7DCEA0"];
              for (let pIdx = 0; pIdx < points.items.length; pIdx++) {
                points.items[pIdx].format.fill.setSolidColor(piePalette[pIdx % piePalette.length]);
              }
            }
          }
        } else if (colors && colors.length > 0) {
          for (let sIdx = 0; sIdx < seriesItems.length; sIdx++) {
            seriesItems[sIdx].format.fill.setSolidColor(colors[sIdx % colors.length]);
          }
        }
      } catch (colorErr) {
        warnings.push(`图表系列颜色设置失败：${colorErr && colorErr.message ? colorErr.message : colorErr}`);
      }

      // ── ISS-93-e：hasDataLabels / smoothLine / yAxis / seriesSettings ──
      const labelsRequested = params.hasDataLabels;
      if (labelsRequested !== undefined) {
        try {
          const labels = chart.dataLabels.load("showValue");
          await context.sync();
          labels.showValue = Boolean(labelsRequested);
          await context.sync();
          labels.load("showValue");
          await context.sync();
          if (Boolean(labels.showValue) !== Boolean(labelsRequested)) {
            warnings.push(`hasDataLabels 请求 ${Boolean(labelsRequested)}，读回 ${Boolean(labels.showValue)}（宿主未接受）。`);
          }
        } catch (labelErr) {
          warnings.push(`hasDataLabels 设置失败：${labelErr && labelErr.message ? labelErr.message : labelErr}`);
        }
      }

      const smoothRequested = params.smoothLine;
      const isLineLike = /line|scatter/i.test(rawType);
      if (smoothRequested !== undefined && !isLineLike) {
        warnings.push(`smoothLine 仅对折线/散点图有意义，当前 chartType=${rawType}，已忽略。`);
      } else if (smoothRequested !== undefined) {
        let smoothApplied = 0;
        for (const s of seriesItems) {
          try { s.smooth = Boolean(smoothRequested); smoothApplied++; } catch (smoothErr) { /* 单系列不支持时跳过 */ }
        }
        if (smoothApplied === 0) {
          warnings.push('smoothLine 设置失败：宿主未接受任何系列的平滑属性。');
        } else {
          try {
            await context.sync();
          } catch (smoothSyncErr) {
            warnings.push(`smoothLine 写入未生效：${smoothSyncErr && smoothSyncErr.message ? smoothSyncErr.message : smoothSyncErr}`);
          }
        }
      }

      const seriesSettings = Array.isArray(params.seriesSettings) ? params.seriesSettings : [];
      for (const setting of seriesSettings) {
        const idx = Number(setting && setting.seriesIndex);
        if (!Number.isFinite(idx) || idx < 1) {
          throw new Error('[Office.js 通道] add_chart 的 seriesSettings.seriesIndex 非法：需要 ≥1 的整数（从 1 开始）。');
        }
        if (idx > seriesItems.length) {
          throw new Error(
            `[Office.js 通道] add_chart 的 seriesSettings.seriesIndex=${idx} 越界：本图只有 ${seriesItems.length} 个系列。` +
            '原实现会静默跳过，现改为显式报错（已建图不会回滚，请按需 delete_chart 后重建）。'
          );
        }
        const target = seriesItems[idx - 1];
        if (setting.color) {
          try { target.format.fill.setSolidColor(setting.color); } catch (cErr) { warnings.push(`seriesSettings[${idx}].color 设置失败：${cErr && cErr.message ? cErr.message : cErr}`); }
        }
        if (setting.smooth !== undefined) {
          try { target.smooth = Boolean(setting.smooth); } catch (smErr) { warnings.push(`seriesSettings[${idx}].smooth 设置失败：${smErr && smErr.message ? smErr.message : smErr}`); }
        }
      }
      if (seriesSettings.length > 0) {
        try { await context.sync(); } catch (ssErr) { warnings.push(`seriesSettings 写入未生效：${ssErr && ssErr.message ? ssErr.message : ssErr}`); }
      }

      const yAxisParam = params.yAxis || {};
      const yAxisRequested = ['min', 'max', 'step', 'numberFormat', 'title'].some(k => yAxisParam[k] !== undefined && yAxisParam[k] !== null && yAxisParam[k] !== '');
      let yAxisApplied = null;
      if (yAxisRequested) {
        try {
          const va = chart.axes.valueAxis;
          if (yAxisParam.min !== undefined && yAxisParam.min !== null) va.minimum = Number(yAxisParam.min);
          if (yAxisParam.max !== undefined && yAxisParam.max !== null) va.maximum = Number(yAxisParam.max);
          if (yAxisParam.step !== undefined && yAxisParam.step !== null) va.majorUnit = Number(yAxisParam.step);
          if (yAxisParam.numberFormat) va.numberFormat = String(yAxisParam.numberFormat);
          if (yAxisParam.title) va.title.text = String(yAxisParam.title);
          await context.sync();
          if (yAxisParam.numberFormat) va.format.numberFormat = String(yAxisParam.numberFormat);
          if (yAxisParam.title) va.title.text = String(yAxisParam.title);
          await context.sync();
          va.load("minimum, maximum, majorUnit, numberFormat, title/text");
          await context.sync();
          yAxisApplied = {
            minimum: va.minimum,
            maximum: va.maximum,
            majorUnit: va.majorUnit,
            numberFormat: va.numberFormat,
            title: va.title ? va.title.text : null
          };
          const mismatches = [];
          if (yAxisParam.min !== undefined && yAxisParam.min !== null && Math.abs(Number(yAxisApplied.minimum) - Number(yAxisParam.min)) > 1e-9) mismatches.push(`min 请求 ${yAxisParam.min} 读回 ${yAxisApplied.minimum}`);
          if (yAxisParam.max !== undefined && yAxisParam.max !== null && Math.abs(Number(yAxisApplied.maximum) - Number(yAxisParam.max)) > 1e-9) mismatches.push(`max 请求 ${yAxisParam.max} 读回 ${yAxisApplied.maximum}`);
          if (mismatches.length > 0) warnings.push(`yAxis 未完全生效：${mismatches.join('；')}`);
        } catch (axisErr) {
          warnings.push(`yAxis 设置失败：${axisErr && axisErr.message ? axisErr.message : axisErr}`);
        }
      }

      chart.load("name, id, top, left, width, height");
      sheet.load("name");
      await context.sync();

      return {
        success: true,
        id: chart.id,
        name: chart.name,
        shapeName: chart.name,
        chartIndex: 1,
        title: targetTitle,
        chartType: rawType,
        dataRange: sourceRange || extraRanges[0],
        dataRangesApplied,
        sheetName: sheet.name,
        left: chart.left !== undefined ? chart.left : targetLeft,
        top: chart.top !== undefined ? chart.top : targetTop,
        width: chart.width !== undefined ? chart.width : targetWidth,
        height: chart.height !== undefined ? chart.height : targetHeight,
        positionMode,
        hasLegendApplied: legendApplied,
        yAxisApplied,
        seriesCount: seriesItems.length,
        seriesSettingsApplied: seriesSettings.length,
        replaceExisting: replaceExisting,
        removedCharts,
        warnings,
        readBackNote: '左侧/顶部/宽高为宿主读回值；用 get_charts(detail=true) 复核标题、系列数与类型。',
        message: `已成功在 [${sheet.name}] 创建 ${rawType} 原生图表，数据源为 ${dataRangesApplied.join(' + ')}`
      };
    });
  }

  async function handleDeleteChart(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      if (params.clearAll) {
        const charts = sheet.charts.load("items");
        await context.sync();
        const count = charts.items.length;
        for (const c of charts.items) {
          c.delete();
        }
        await context.sync();
        return { success: true, count, message: `已成功清空当前工作表中的全部 ${count} 个图表` };
      }
      const wanted = params.chartName || params.name || params.id || params.shapeName;
      if (!wanted) {
        throw new Error('[Office.js 通道] delete_chart 需要 chartName（或 shapeName/id），否则不知道删哪一张；如需清空全部请显式传 clearAll=true。');
      }
      const chart = sheet.charts.getItem(wanted);
      chart.delete();
      await context.sync();
      return { success: true, deleted: wanted, message: `图表 [${wanted}] 已成功删除` };
    });
  }

  async function handleUpdateChart(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const wanted = params.name || params.chartName || params.id || params.shapeName;
      if (!wanted) {
        throw new Error('[Office.js 通道] update_chart 需要 name/chartName/shapeName/id 指定目标图表。');
      }
      const chart = sheet.charts.getItem(wanted);
      if (params.title) chart.title.text = params.title;
      if (params.legendPosition) chart.legend.position = params.legendPosition;
      const targetLeft = params.left !== undefined ? Number(params.left) : (params.position?.left ? Number(params.position.left) : undefined);
      const targetTop = params.top !== undefined ? Number(params.top) : (params.position?.top ? Number(params.position.top) : undefined);
      const targetWidth = params.width !== undefined ? Number(params.width) : (params.position?.width ? Number(params.position.width) : undefined);
      const targetHeight = params.height !== undefined ? Number(params.height) : (params.position?.height ? Number(params.position.height) : undefined);

      applyChartPosition(chart, sheet, params, targetLeft, targetTop, targetWidth, targetHeight);
      await context.sync();
      return { success: true, name: wanted, message: `图表 [${wanted}] 已成功更新位置与配置` };
    });
  }

  async function handleExportChartImage(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const wanted = params.name || params.chartName || params.shapeName;
      if (!wanted) {
        throw new Error('[Office.js 通道] export_chart_image 需要 chartName（或 name/shapeName）指定图表。');
      }
      const chart = sheet.charts.getItem(wanted);
      const size = readRenderSize({ width: params.width, height: params.height });
      const imageResult = chart.getImage(size.width, size.height);
      await context.sync();
      return { success: true, chartName: wanted, imageBase64: imageResult.value, imageMimeType: 'image/png', width: size.width, height: size.height, renderedBy: 'chart.getImage (Office.js 原生导出)' };
    });
  }

  /**
   * ISS-96：capture_sheet_preview —— 只做真实渲染，做不到就明确报错。
   *
   * 真实渲染的**唯一**来源是 `chart.getImage()`（宿主原生导出图表）。
   * 工作表区域截图在 Office.js 交付面上没有可用 API（macOS 桌面版既无 Range 截图，
   * 也没有可用的 Workbook 渲染导出），因此区域预览一律抛错并给出替代路径。
   */
  async function handleCaptureSheetPreview(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params?.sheetName);
      // 报错文案要用工作表名，而每个分支都要 sync：load 必须在**第一次 sync 之前**排进队列，
      // 否则在"区域截图"这条路径上永远 sync 不到 name，真正的报错会被
      // 「属性"name"不可用」盖掉（实机踩到）。
      sheet.load("name");
      const chartSelector = params?.chartName || params?.name || null;
      const address = params?.address || params?.range || null;
      const requestedMode = params?.mode === undefined || params?.mode === null || params?.mode === '' ? null : String(params.mode).toLowerCase();
      const targetChartName = requestedMode === 'sheet' ? null : chartSelector;
      const wantsSheetArea = requestedMode === 'sheet' || (!targetChartName && Boolean(address));
      const charts = sheet.charts.load("items/name, items/id, items/title/text, items/width, items/height");
      await context.sync();

      if (requestedMode && !['chart', 'sheet', 'auto'].includes(requestedMode)) {
        throw new Error(`[Office.js 通道] capture_sheet_preview 无法识别的 mode: "${params.mode}"（支持 chart | sheet | auto）。`);
      }
      if (requestedMode === 'sheet' && chartSelector) {
        throw new Error(
          `[Office.js 通道] capture_sheet_preview 参数冲突：mode="sheet" 与 chartName="${chartSelector}" 同时给出。` +
          '区域截图不受支持（见下），如需导出图表请去掉 mode 或传 mode="chart"。'
        );
      }

      if (targetChartName) {
        if (charts.items.length === 0) {
          throw new Error(`[Office.js 通道] capture_sheet_preview 无法渲染：工作表 [${sheet.name}] 没有任何原生图表，而你请求的是图表 "${targetChartName}"。`);
        }
        const target = findChartByName(charts.items, targetChartName);
        if (!target) {
          throw new Error(
            `[Office.js 通道] capture_sheet_preview 未找到图表 "${targetChartName}"（按名称或 id 精确匹配，兼容 "图表 2" 与 "Chart 2"）。` +
            `[${sheet.name}] 现有图表：${chartNamesOf(charts.items).join('、')}。`
          );
        }
        const size = readRenderSize(params || {});
        const imgResult = target.getImage(size.width, size.height);
        await context.sync();
        if (!imgResult || !imgResult.value) {
          throw new Error(`[Office.js 通道] capture_sheet_preview：宿主对图表 [${target.name}] 的 getImage 未返回图像数据，无法提供真实渲染图。`);
        }
        const title = target.title && target.title.text ? target.title.text : '';
        return {
          success: true,
          kind: 'chart',
          renderedBy: 'chart.getImage (Office.js 原生导出)',
          workbookName: params?.workbookName || null,
          sheetName: sheet.name,
          address: target.name,
          chartName: target.name,
          chartTitle: title,
          imageBase64: imgResult.value,
          imageMimeType: "image/png",
          width: size.width,
          height: size.height,
          message: `已导出 [${sheet.name}] 图表 [${target.name}] 的原生渲染图（${size.width}×${size.height}，由 Office.js chart.getImage 生成，非合成图）。`
        };
      }

      // 走到这里说明请求的是工作表区域（或既没给图表也没给区域）。
      const scope = wantsSheetArea ? `区域 ${address}` : '工作表已用区域';
      throw new Error(
        `[Office.js 通道] capture_sheet_preview 不支持真实渲染：无法为 [${sheet.name}] 的${scope}生成截图。` +
        'Office.js 交付面没有工作表/区域截图 API（只有图表有 chart.getImage），因此本通道不提供该图，' +
        '也不会用 Canvas 合成"示意图"充数（旧版本会返回假渲染图，已移除）。' +
        '替代路径：① 改用 host="wps" 的 wps_capture_sheet_preview（WPS 原生渲染，可截指定区域）；' +
        '② 在 Excel 里自行截图（选区后 Shift+Cmd+4/Ctrl+C 复制为图片）再把图交给 AI；' +
        '③ 需要"读回核对"时改用 read_range / get_range_styles / get_charts(detail=true) 做数据与格式自查。'
      );
    });
  }

// ── 模块: src/excel/pivot.js — 数据透视表 PivotTable ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1663-1687 行（原样搬迁，未改写）。
  // 9. 数据透视表
  async function handleCreatePivotTable(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const source = sheet.getRange(params.sourceAddress);
      const dest = sheet.getRange(params.destinationAddress || "R1C1");
      const name = params.name || "PivotTable1";

      const pt = sheet.pivotTables.add(name, source, dest);
      pt.load("name");
      await context.sync();
      return { success: true, name: pt.name };
    });
  }

  async function handleUpdatePivotTable(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const pt = sheet.pivotTables.getItem(params.name || params.pivotTableName);
      if (params.refresh) pt.refresh();
      await context.sync();
      return { success: true };
    });
  }

// ── 模块: src/excel/shape.js — 形状与图片（CAP-08：MS 侧矢量绘图） ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1688-1731 行。
//
// 设计约束（CAP-08）：
// 1. **写后必读**：每个写操作都在同一批次里把真实落位的属性读回来返回给调用方，
//    不做"发了请求就报成功"的假成功；宿主没接受的属性写进 `warnings` 如实暴露。
// 2. **没实现就不假装**：某项属性设置失败只把该项写进 `warnings`，不阻断其它属性，
//    也不把失败折叠成成功。
// 3. 不用 Canvas / 任何本地合成去"造"图形 —— 形状只能由宿主真实创建。
  // 10. 形状与图片

  /** 几何形状类型的别名表（未知类型时给出可执行提示，而不是甩一句"参数错误"）。 */
  const SHAPE_TYPE_ALIASES = {
    rectangle: "rectangle", rect: "rectangle", square: "rectangle",
    rounded_rectangle: "roundedRectangle", roundedrectangle: "roundedRectangle", round_rect: "roundedRectangle",
    oval: "oval", ellipse: "oval", circle: "oval",
    triangle: "triangle", right_triangle: "rightTriangle",
    diamond: "diamond", rhombus: "diamond",
    pentagon: "pentagon", hexagon: "hexagon", octagon: "octagon",
    star: "star5", star5: "star5", star4: "star4", star8: "star8",
    arrow: "rightArrow", right_arrow: "rightArrow",
    left_arrow: "leftArrow", up_arrow: "upArrow", down_arrow: "downArrow",
    double_arrow: "leftRightArrow",
    cloud: "cloud", heart: "heart", sun: "sun", moon: "moon",
    lightning_bolt: "lightningBolt",
    flow_chart_process: "flowChartProcess", flow_chart_decision: "flowChartDecision", flow_chart_data: "flowChartData",
    flowchart_process: "flowChartProcess", flowchart_decision: "flowChartDecision", flowchart_data: "flowChartData",
    trapezoid: "trapezoid", parallelogram: "parallelogram", cross: "cross", plus: "cross",
    can: "can", cube: "cube", donut: "donut", plaque: "plaque",
    text_box: "textBox", textbox: "textBox"
  };

  /** 把调用方给的类型名解析成 `Excel.GeometricShapeType` 的枚举值。 */
  function resolveGeometricShapeType(raw) {
    const key = String(raw || "rectangle").trim();
    const alias = SHAPE_TYPE_ALIASES[key.toLowerCase()] || key;
    const all = (typeof Excel !== "undefined" && Excel.GeometricShapeType) || {};
    if (all[alias] !== undefined) return { value: all[alias], name: alias };
    for (const k of Object.keys(all)) {
      if (k.toLowerCase() === String(alias).toLowerCase()) return { value: all[k], name: k };
    }
    throw new Error(`不认识的几何形状类型 "${raw}"。可用枚举键（部分）：${Object.keys(all).slice(0, 40).join(", ")}`);
  }

  /** 解析要操作的目标形状；优先 name（稳定），其次 id。 */
  function resolveShape(sheet, params) {
    const name = params.name || params.shapeName;
    if (name) return sheet.shapes.getItem(name);
    if (params.shapeId) return sheet.shapes.getItem(params.shapeId);
    throw new Error("缺少必要参数：name（形状名，推荐用 list_shapes 读到的 name）或 shapeId");
  }

  /** 十六进制颜色归一：#RGB / #RRGGBB / 无 # 都接受；非法值返回 null（由调用方如实报告）。 */
  function normalizeHexColor(raw) {
    if (raw === undefined || raw === null || raw === "") return null;
    let s = String(raw).trim();
    if (!s.startsWith("#")) s = "#" + s;
    if (/^#[0-9a-fA-F]{3}$/.test(s)) s = "#" + s[1] + s[1] + s[2] + s[2] + s[3] + s[3];
    return /^#[0-9a-fA-F]{6}$/.test(s) ? s.toUpperCase() : null;
  }

  /** getAsImage 的格式参数：用枚举值，取不到时退回字符串字面量（Office.js 接受枚举字符串）。 */
  function resolvePictureFormat(raw) {
    const wantJpeg = String(raw || "png").toLowerCase() === "jpeg" || String(raw || "").toLowerCase() === "jpg";
    try {
      if (typeof Excel !== "undefined" && Excel.PictureFormat) return wantJpeg ? Excel.PictureFormat.jpeg : Excel.PictureFormat.png;
    } catch (e) { /* 该版本没有枚举则走字符串 */ }
    return wantJpeg ? "JPEG" : "PNG";
  }

  /** 一次批次里把形状的全部可读属性挂上 load（读回口径只在这里维护一份）。 */
  /**
   * 按形状类型加载属性。
   *
   * ⚠️ **必须分两阶段**：`type` 本身也要 load 才读得到。
   * 一阶段写法（先 `shape.load("...type...")` 再立刻 `String(shape.type)` 判断）拿到的是
   * **undefined**，于是 `isLine`/`isGroup` 恒为 false，照样对 Line/Group load `fill`/`textFrame`
   * → `context.sync()` 抛「当前对象不允许此操作」。
   * 这正是「直线建不出来」「分组建不出来」看起来像宿主限制的原因**（真机复核：对象其实建成了）。
   *
   * 注意：`try/catch` 挡不住这个错——Office.js 的 load() 是延迟的，失败发生在 sync() 时。
   *
   * 调用方约定：调用本函数后需自行 `await context.sync()`；如需按类型决定加载项，
   * 请用 `loadShapeDetailTyped(context, shape)`（它会自己 sync 一次）。
   */
  function loadShapeDetail(shape) {
    shape.load("name,id,type,left,top,width,height,rotation,zOrderPosition,visible");
    shape.load("lineFormat/color,lineFormat/weight,lineFormat/visible");
    return shape;
  }

  /**
   * 分两阶段加载：先只 load `type` 并 sync，拿到真实类型后再补 load 其余属性。
   * 对 Line 不 load `fill`，对 Group 不 load `fill`/`textFrame`——它们没有这些属性。
   */
  async function loadShapeDetailTyped(context, shape) {
    shape.load("type");
    await context.sync();                 // 第一阶段：只为了让 type 可用
    const typeName = (() => { try { return String(shape.type); } catch (e) { return ""; } })();
    const isLine = /Line/i.test(typeName);
    const isGroup = /Group/i.test(typeName);
    shape.load("name,id,left,top,width,height,rotation,zOrderPosition,visible");
    if (!isLine && !isGroup) shape.load("fill/type,fill/foregroundColor,fill/transparency");
    shape.load("lineFormat/color,lineFormat/weight,lineFormat/visible");
    if (!isGroup) {
      shape.load("textFrame/textRange/text");
      shape.load("textFrame/textRange/font/bold,textFrame/textRange/font/size,textFrame/textRange/font/color");
      shape.load("textFrame/horizontalAlignment,textFrame/verticalAlignment");
    }
    await context.sync();                 // 第二阶段：补属性
    return shape;
  }

  /** 把已 load 的 shape 折算成扁平 JSON（不返回宿主代理对象）。 */
  function shapeToJson(shape) {
    const out = {
      name: shape.name,
      id: shape.id,
      type: String(shape.type),
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
      rotation: shape.rotation,
      zOrderPosition: shape.zOrderPosition === undefined || shape.zOrderPosition === null ? null : String(shape.zOrderPosition),
      visible: shape.visible
    };
    try { out.fill = { type: shape.fill ? String(shape.fill.type) : null, foregroundColor: shape.fill ? shape.fill.foregroundColor : null, transparency: shape.fill ? shape.fill.transparency : null }; }
    catch (e) { out.fill = { unsupported: true, error: String(e.message || e) }; }
    try { out.line = { color: shape.lineFormat ? shape.lineFormat.color : null, weight: shape.lineFormat ? shape.lineFormat.weight : null, visible: shape.lineFormat ? shape.lineFormat.visible : null }; }
    catch (e) { out.line = { unsupported: true, error: String(e.message || e) }; }
    try {
      const tf = shape.textFrame;
      if (tf && tf.textRange) {
        out.text = tf.textRange.text;
        out.textFormat = {
          bold: tf.textRange.font ? tf.textRange.font.bold : null,
          size: tf.textRange.font ? tf.textRange.font.size : null,
          color: tf.textRange.font ? tf.textRange.font.color : null,
          horizontalAlignment: tf.horizontalAlignment === undefined ? null : String(tf.horizontalAlignment),
          verticalAlignment: tf.verticalAlignment === undefined ? null : String(tf.verticalAlignment)
        };
      } else out.text = null;
    } catch (e) { out.text = null; out.textUnsupported = String(e.message || e); }
    return out;
  }

  /** 读回核对：把"请求值 vs 宿主实际值"逐项比对，只报真实差异。 */
  function verifyAgainstRequest(actual, expected) {
    const mismatches = [];
    for (const key of Object.keys(expected)) {
      let want;
      if (key === "fillColor") want = actual.fill && actual.fill.foregroundColor;
      else if (key === "lineColor") want = actual.line && actual.line.color;
      else if (key === "lineWeight") want = actual.line && actual.line.weight;
      else if (key === "text") want = actual.text;
      else if (key === "name") want = actual.name;
      else if (key === "rotationDelta" || key === "scaleWidth" || key === "scaleHeight") continue; // 增量类无法按等值比对
      else want = actual[key];
      if (want === undefined || want === null) { mismatches.push(`${key}: 宿主未回读该属性`); continue; }
      const a = typeof want === "string" ? want.toUpperCase() : want;
      const b = (typeof expected[key] === "string" && key !== "text") ? String(expected[key]).toUpperCase() : expected[key];
      if (a !== b) mismatches.push(`${key}: 请求 ${expected[key]} / 实际 ${want}`);
    }
    return mismatches;
  }

  /**
   * 添加矢量图形：矩形/椭圆/箭头等几何形状、直线/连接符、SVG、文本框。
   * 一次调用只加一个形状；返回值里带**宿主读回**的实际位置尺寸与颜色。
   */
  async function handleAddShape(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const warnings = [];
      const requested = {};
      const kind = String(params.kind || "geometric").toLowerCase();

      let shape;
      if (kind === "line" || kind === "connector") {
        // addLine 收的是起止点，不是 left/top/width/height。
        const x1 = params.x1 !== undefined ? params.x1 : (params.left !== undefined ? params.left : 0);
        const y1 = params.y1 !== undefined ? params.y1 : (params.top !== undefined ? params.top : 0);
        const x2 = params.x2 !== undefined ? params.x2 : (params.right !== undefined ? params.right : x1 + (params.width || 100));
        const y2 = params.y2 !== undefined ? params.y2 : (params.bottom !== undefined ? params.bottom : y1 + (params.height || 0));
        shape = sheet.shapes.addLine(x1, y1, x2, y2);
        requested.x1 = x1; requested.y1 = y1; requested.x2 = x2; requested.y2 = y2;
      } else if (kind === "svg") {
        const svg = params.svg
          || (params.base64Image && typeof atob === "function" ? atob(params.base64Image) : null);
        if (!svg) throw new Error("kind=svg 需要传 svg（SVG 源码字符串）或 base64Image（base64 编码的 SVG）");
        const b64 = params.base64Image || (typeof btoa === "function" ? btoa(svg) : "");
        if (!b64) throw new Error("当前环境无法把 SVG 转成 base64（缺少 btoa）");
        shape = sheet.shapes.addSvg(b64);
      } else if (kind === "textbox") {
        shape = sheet.shapes.addTextBox(params.text !== undefined && params.text !== null ? String(params.text) : "");
      } else {
        shape = sheet.shapes.addGeometricShape(resolveGeometricShapeType(params.shapeType || params.type).value);
      }

      if (params.shapeName || params.name) shape.name = params.shapeName || params.name;

      for (const pair of [["left", params.left], ["top", params.top], ["width", params.width], ["height", params.height]]) {
        if (pair[1] === undefined || pair[1] === null) continue;
        shape[pair[0]] = pair[1]; requested[pair[0]] = pair[1];
      }
      if (params.rotation !== undefined && params.rotation !== null) { shape.rotation = params.rotation; requested.rotation = params.rotation; }

      // 文本框的初始文字已在 addTextBox 里给出；其余形状用 textFrame 写文字。
      if (kind !== "textbox" && params.text !== undefined && params.text !== null) {
        shape.textFrame.textRange.text = String(params.text);
        requested.text = String(params.text);
      }

      const fillHex = normalizeHexColor(params.fillColor || params.fill);
      if (fillHex) { shape.fill.setSolidColor(fillHex); requested.fillColor = fillHex; }
      else if (params.fillColor || params.fill) warnings.push(`fillColor "${params.fillColor || params.fill}" 不是合法十六进制颜色，未应用`);

      const lineHex = normalizeHexColor(params.lineColor);
      if (lineHex) { shape.lineFormat.color = lineHex; requested.lineColor = lineHex; }
      if (params.lineWeight !== undefined && params.lineWeight !== null) { shape.lineFormat.weight = params.lineWeight; requested.lineWeight = params.lineWeight; }

      await loadShapeDetailTyped(context, shape);
      await context.sync();
      const actual = shapeToJson(shape);
      const mismatches = verifyAgainstRequest(actual, requested);
      return {
        success: true,
        sheetName: sheet.name,
        kind,
        shape: actual,
        requested,
        verified: mismatches.length === 0,
        warnings: warnings.concat(mismatches.map(m => `读回与请求不一致：${m}`))
      };
    });
  }

  async function handleInsertImage(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      if (!params.base64Image) throw new Error("缺少必要参数：base64Image");
      const shape = sheet.shapes.addImage(params.base64Image);
      if (params.left) shape.left = params.left;
      if (params.top) shape.top = params.top;
      if (params.width) shape.width = params.width;
      if (params.height) shape.height = params.height;
      if (params.name || params.shapeName) shape.name = params.name || params.shapeName;

      await loadShapeDetailTyped(context, shape);
      await context.sync();
      return { success: true, sheetName: sheet.name, shape: shapeToJson(shape), verified: true };
    });
  }

  /** 读回：列表 + 位置尺寸 + 填充/线条 + 旋转 + 层级（CAP-08 的读回入口）。 */
  async function handleListShapes(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shapes = sheet.shapes;
      // 整批一次 sync，避免形状多时 N 次往返；宿主不认某个子属性时退化为不读该属性。
      shapes.load("items/name,items/id,items/type,items/left,items/top,items/width,items/height,items/rotation,items/zOrderPosition,items/visible");
      try { shapes.load("items/fill/type,items/fill/foregroundColor,items/fill/transparency"); } catch (e) {}
      try { shapes.load("items/lineFormat/color,items/lineFormat/weight"); } catch (e) {}
      try { shapes.load("items/textFrame/textRange/text"); } catch (e) {}
      await context.sync();

      const items = shapes.items.map(shapeToJson);
      const ordered = items.slice().sort((a, b) => Number(a.zOrderPosition) - Number(b.zOrderPosition));
      return {
        success: true,
        count: items.length,
        sheetName: sheet.name,
        shapes: items,
        zOrderBottomToTop: ordered.map(s => ({ name: s.name, zOrderPosition: s.zOrderPosition })),
        message: `工作表 [${sheet.name}] 共有 ${items.length} 个形状`
      };
    });
  }

  /** 改形状：位置/尺寸/旋转/缩放/填充/线条/文字/改名，或删除；同样带读回。 */
  async function handleUpdateShape(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const warnings = [];
      const requested = {};

      if (params.action === "delete") {
        const target = resolveShape(sheet, params);
        target.delete();
        await context.sync();
        const rest = sheet.shapes.load("items/name");
        await context.sync();
        const stillThere = rest.items.some(s => s.name === (params.name || params.shapeName));
        return {
          success: !stillThere,
          deleted: params.name || params.shapeName,
          remaining: rest.items.map(s => s.name),
          verified: !stillThere,
          warnings: stillThere ? ["删除后仍能读到该形状，删除未生效"] : []
        };
      }

      const shape = resolveShape(sheet, params);
      await loadShapeDetailTyped(context, shape);
      await context.sync();
      const beforeJson = shapeToJson(shape);

      for (const pair of [["left", params.left], ["top", params.top], ["width", params.width], ["height", params.height]]) {
        if (pair[1] === undefined || pair[1] === null) continue;
        shape[pair[0]] = pair[1]; requested[pair[0]] = pair[1];
      }
      if (params.rotation !== undefined && params.rotation !== null) { shape.rotation = params.rotation; requested.rotation = params.rotation; }
      if (params.newName || params.rename) { shape.name = params.newName || params.rename; requested.name = params.newName || params.rename; }
      if (params.text !== undefined && params.text !== null) { shape.textFrame.textRange.text = String(params.text); requested.text = String(params.text); }

      const fillHex = normalizeHexColor(params.fillColor || params.fill);
      if (fillHex) { shape.fill.setSolidColor(fillHex); requested.fillColor = fillHex; }
      else if (params.fillColor || params.fill) warnings.push(`fillColor "${params.fillColor || params.fill}" 不是合法十六进制颜色，未应用`);

      const lineHex = normalizeHexColor(params.lineColor);
      if (lineHex) { shape.lineFormat.color = lineHex; requested.lineColor = lineHex; }
      if (params.lineWeight !== undefined && params.lineWeight !== null) { shape.lineFormat.weight = params.lineWeight; requested.lineWeight = params.lineWeight; }

      if (params.rotationDelta !== undefined && params.rotationDelta !== null) { shape.incrementRotation(params.rotationDelta); requested.rotationDelta = params.rotationDelta; }
      if (params.scaleWidth !== undefined && params.scaleWidth !== null) {
        shape.scaleWidth(params.scaleWidth, params.scaleType === "originalSize" ? "OriginalSize" : "CurrentSize");
        requested.scaleWidth = params.scaleWidth;
      }
      if (params.scaleHeight !== undefined && params.scaleHeight !== null) {
        shape.scaleHeight(params.scaleHeight, params.scaleType === "originalSize" ? "OriginalSize" : "CurrentSize");
        requested.scaleHeight = params.scaleHeight;
      }

      if (Object.keys(requested).length === 0) {
        return {
          success: true,
          changed: false,
          shape: beforeJson,
          warnings: ["没有传入任何可修改字段（left/top/width/height/rotation/fillColor/lineColor/lineWeight/text/newName/scale*），未做修改"]
        };
      }

      await loadShapeDetailTyped(context, shape);
      await context.sync();
      const after = shapeToJson(shape);
      const mismatches = verifyAgainstRequest(after, requested);
      return {
        success: true,
        changed: true,
        sheetName: sheet.name,
        before: beforeJson,
        shape: after,
        requested,
        verified: mismatches.length === 0,
        warnings: warnings.concat(mismatches.map(m => `读回与请求不一致：${m}`))
      };
    });
  }

  async function handleGroupShapes(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const names = Array.isArray(params.shapeNames) ? params.shapeNames.filter(Boolean) : [];
      if (names.length < 2) throw new Error("group_shapes 需要至少 2 个形状（shapeNames 传形状名数组）");
      // 组合是"活动表"范围的操作：真机实测，对非活动表调 addGroup 会报
      // 「当前对象不允许此操作」（子代理的自检能过，是因为它新建表时顺带激活了那张表）。
      sheet.activate();
      await context.sync();   // 先让激活单独生效，再发起组合
      const proxies = names.map(n => sheet.shapes.getItem(n));
      proxies.forEach(p => p.load("id,name"));
      await context.sync();
      const ids = proxies.map(p => p.id);
      const group = sheet.shapes.addGroup(ids);
      if (params.groupName || params.name) group.name = params.groupName || params.name;
      await loadShapeDetailTyped(context, group);
      await context.sync();

      const groupJson = shapeToJson(group);
      let members = [];
      try {
        const memberShapes = group.group.shapes.load("items/name,items/type");
        await context.sync();
        members = memberShapes.items.map(s => ({ name: s.name, type: String(s.type) }));
      } catch (e) {
        groupJson.memberReadError = String(e.message || e);
      }
      return {
        success: true,
        sheetName: sheet.name,
        group: groupJson,
        memberCount: members.length,
        members,
        verified: members.length === ids.length,
        warnings: members.length === ids.length ? [] : [`请求分组 ${ids.length} 个形状，读回组内 ${members.length} 个`]
      };
    });
  }

  async function handleUngroupShapes(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const name = params.groupName || params.name;
      const group = sheet.shapes.getItem(name);
      group.load("id,type");
      await context.sync();
      // 先确认这确实是一个组合（ShapeType.group），再谈解组：对普通形状调 ungroup
      // 会报出难懂的宿主错误，提前拒绝对调用方更有用。
      if (String(group.type).toLowerCase().indexOf("group") < 0) {
        throw new Error(`"${name}" 不是组合（读到 type=${group.type}），无法解组。先用 list_shapes 确认组合名。`);
      }
      if (!group.group || !group.group.shapes) {
        throw new Error(`当前 Excel 版本读不到组合成员（Shape.group 不可用），无法解组 "${name}"。`);
      }
      const childShapes = group.group.shapes.load("items/name");
      await context.sync();
      const childNames = childShapes.items.map(s => s.name);
      if (typeof group.group.ungroup !== "function") {
        throw new Error("当前 Excel 版本不支持 ShapeGroup.ungroup()：只能删除整组，或改用 update_shape 单独移动成员");
      }
      group.group.ungroup();
      await context.sync();
      const rest = sheet.shapes.load("items/name");
      await context.sync();
      const remaining = rest.items.map(s => s.name);
      const childrenBack = childNames.filter(n => remaining.indexOf(n) >= 0);
      return {
        success: true,
        sheetName: sheet.name,
        ungrouped: name,
        releasedChildren: childNames,
        childrenStillPresent: childrenBack,
        shapesAfter: remaining,
        verified: childrenBack.length === childNames.length
      };
    });
  }

  /** 层级调整：bringToFront / sendToBack / bringForward / sendBackward，返回整表层级读回。 */
  async function handleSetShapeZOrder(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shape = resolveShape(sheet, params);
      const raw = String(params.zOrder || params.action || "bringToFront").toLowerCase().replace(/[\s_-]/g, "");
      const map = {
        bringtofront: "BringToFront", front: "BringToFront",
        sendtoback: "SendToBack", back: "SendToBack",
        bringforward: "BringForward", forward: "BringForward",
        sendbackward: "SendBackward", backward: "SendBackward"
      };
      const target = map[raw];
      if (target === undefined) throw new Error(`不认识的层级动作 "${params.zOrder || params.action}"，可用：bringToFront / sendToBack / bringForward / sendBackward`);
      shape.setZOrder(target);
      await context.sync();

      const all = sheet.shapes.load("items/name,items/zOrderPosition");
      await context.sync();
      const ordered = all.items
        .map(s => ({ name: s.name, zOrderPosition: s.zOrderPosition === undefined || s.zOrderPosition === null ? null : String(s.zOrderPosition) }))
        .sort((a, b) => Number(a.zOrderPosition) - Number(b.zOrderPosition));
      const self = ordered.filter(s => s.name === (params.name || params.shapeName))[0];
      const expectedEdge = target === "BringToFront" ? ordered.length - 1 : target === "SendToBack" ? 0 : null;
      return {
        success: true,
        sheetName: sheet.name,
        applied: target,
        shape: (params.name || params.shapeName) || params.shapeId,
        shapeZOrderPosition: self ? self.zOrderPosition : null,
        zOrderBottomToTop: ordered,
        // 端点动作可以强校验；bringForward/sendBackward 是相对移动，只如实给读回
        verified: expectedEdge === null ? null : Boolean(self && Number(self.zOrderPosition) === expectedEdge)
      };
    });
  }

  /** 读回：当前选中的形状（用户手上正在编辑哪一个）。 */
  async function handleGetActiveShape(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shape = sheet.shapes.getActiveShape();
      await loadShapeDetailTyped(context, shape);
      await context.sync();
      return { success: true, sheetName: sheet.name, shape: shapeToJson(shape) };
    });
  }

  /** 把形状导出成 PNG/JPEG（base64），用于视觉自检。 */
  async function handleExportShapeImage(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shape = resolveShape(sheet, params);
      const format = resolvePictureFormat(params.format);
      const scale = params.scale || 1;
      const image = shape.getAsImage(format, scale);
      await context.sync();
      const base64 = image.value;
      if (!base64) throw new Error("宿主未返回图像数据（getAsImage 返回空）");
      return {
        success: true,
        sheetName: sheet.name,
        name: params.name || params.shapeName,
        format: String(params.format || "png").toLowerCase(),
        scale,
        imageBase64: base64,
        byteEstimate: Math.round(base64.length * 0.75)
      };
    });
  }

  // ── CAP-08 能力自检 ──
  // 为什么要在加载项里自带自检：本机 MCP 工具面没有通道能把任意 Office.js 代码送进任务窗格
  // （excel_* 路由表不含形状方法；office_execute_script 走 macOS JXA / Windows COM），
  // 只靠桥接侧无法回答"这台 Excel 到底支持到哪一步"。自检在真实窗格里跑、写回一张工作表，
  // 结果可以被任意 host=microsoft 的读工具读回，也能被下一次会话直接复用。
  // 自检**逐项报可用/不可用**，任何一项失败都不影响其它项；不支持的项如实记录错误原文。
  let CAP08_SELF_TEST_RESULT = null;

  async function handleShapeSelfTest(params) {
    const sheetName = params.sheetName || "_cap08_shapetest";
    const keep = params.keep === true;           // 默认清理探测表，不留垃圾
    const started = new Date().toISOString();
    const checks = [];
    const record = async (name, fn) => {
      const t0 = Date.now();
      try {
        const detail = await fn();
        checks.push({ check: name, ok: true, ms: Date.now() - t0, detail: detail === undefined ? null : detail });
        return detail;
      } catch (e) {
        checks.push({ check: name, ok: false, ms: Date.now() - t0, error: String((e && e.message) || e), code: e && e.code });
        return null;
      }
    };

    const runOn = async (fn) => await Excel.run(fn);
    const sheetOf = (context) => context.workbook.worksheets.getItem(sheetName);

    /**
     * 让检查之间**互不依赖**。
     *
     * 原实现里后面的检查直接引用前面创建的形状（cap08_line / cap08_svg 等），
     * 一旦前面失败，后面会报"请求的资源不存在"——看上去像"分组/删除也不支持"，
     * 实际只是前置对象不存在。**级联失败会把"没做出来"误报成"宿主不支持"**，
     * 所以这里统一改成"需要什么就先确保它存在"。
     */
    const ensureShape = async (name, make) => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const existing = sheet.shapes.load("items/name");
      await context.sync();
      if (!existing.items.some(s => s.name === name)) {
        const sh = make(sheet);
        if (sh) sh.name = name;
      }
      await context.sync();
      return name;
    });
    const ensureRect = () => ensureShape("cap08_rect", (sheet) =>
      sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle, { left: 20, top: 20, width: 160, height: 90 }));
    const ensureText = () => ensureShape("cap08_text", (sheet) =>
      sheet.shapes.addTextBox("CAP-08 文本框"));

    try {
      await runOn(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        if (sheets.items.some(s => s.name === sheetName)) sheets.getItem(sheetName).delete();
        await context.sync();
        sheets.add(sheetName);
        await context.sync();
      });
    } catch (e) {
      return { success: false, stage: "createSheet", error: String(e.message || e), checks, started };
    }

    // 1. 几何形状
    const rect = await record("addGeometricShape(rectangle) + 填充/线条/旋转读回", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const sh = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle, { left: 20, top: 20, width: 160, height: 90 });
      sh.name = "cap08_rect";
      sh.fill.setSolidColor("#2F6FEB");
      sh.lineFormat.color = "#12305E";
      sh.lineFormat.weight = 2;
      sh.rotation = 15;
      sh.textFrame.textRange.text = "CAP-08";
      await loadShapeDetailTyped(context, sh);
      await context.sync();
      return shapeToJson(sh);
    }));

    // 2. 直线
    const line = await record("addLine（连接符/直线）", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const sh = sheet.shapes.addLine(200, 20, 360, 110);
      sh.name = "cap08_line";
      sh.lineFormat.color = "#E4572E";
      sh.lineFormat.weight = 3;
      await loadShapeDetailTyped(context, sh);
      await context.sync();
      return shapeToJson(sh);
    }));

    // 3. SVG
    const svgOk = await record("addSvg", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="80" viewBox="0 0 120 80"><rect width="120" height="80" fill="#0F9D58"/><circle cx="60" cy="40" r="26" fill="#ffffff"/></svg>';
      const b64 = typeof btoa === "function" ? btoa(svg) : null;
      if (!b64) throw new Error("当前环境没有 btoa，无法把 SVG 转 base64");
      const sh = sheet.shapes.addSvg(b64);
      sh.name = "cap08_svg";
      await loadShapeDetailTyped(context, sh);
      await context.sync();
      return shapeToJson(sh);
    }));

    // 4. 文本框
    const textBox = await record("addTextBox + 字体读回", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const sh = sheet.shapes.addTextBox("CAP-08 文本框");
      sh.name = "cap08_text";
      sh.left = 20; sh.top = 140; sh.width = 200; sh.height = 50;
      sh.textFrame.textRange.font.bold = true;
      sh.textFrame.textRange.font.size = 14;
      await loadShapeDetailTyped(context, sh);
      await context.sync();
      return shapeToJson(sh);
    }));

    // 5. 列表读回
    const list = await record("shapes 列表读回（几何属性 + zOrder）", async () => {
      await ensureRect();
      await ensureText();
      return await runOn(async (context) => {
        const sheet = sheetOf(context);
        // 只 load 标量属性：嵌套的 fill/lineFormat 需要逐个 load 才能读，
        // 混在一起 load 会让整批失败（真机实测报"当前对象不允许此操作"）。
        const shapes = sheet.shapes.load("items/name,items/type,items/left,items/top,items/width,items/height,items/rotation,items/zOrderPosition");
        await context.sync();
        return shapes.items.map(s => ({
          name: s.name, type: String(s.type), left: s.left, top: s.top,
          width: s.width, height: s.height, rotation: s.rotation, z: String(s.zOrderPosition)
        }));
      });
    });
    const listFill = await record("shapes 填充/线条读回（逐个 load）", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const sh = sheet.shapes.getItem("cap08_rect");
      sh.fill.load("foregroundColor");
      sh.lineFormat.load("color,weight");
      await context.sync();
      return { fill: sh.fill.foregroundColor, lineColor: sh.lineFormat.color, lineWeight: sh.lineFormat.weight };
    }));

    // 6. 层级
    const zorder = await record("setZOrder(bringToFront/sendToBack)", async () => {
      await ensureRect(); await ensureText();
      return await runOn(async (context) => {
      const sheet = sheetOf(context);
      // 不引用 cap08_line（它可能没创建成功）；用两个确定存在的形状对比层级
      sheet.shapes.getItem("cap08_rect").setZOrder(Excel.ShapeZOrder.sendToBack);
      sheet.shapes.getItem("cap08_text").setZOrder(Excel.ShapeZOrder.bringToFront);
      const all = sheet.shapes.load("items/name,items/zOrderPosition");
      await context.sync();
      return all.items.map(s => ({ name: s.name, z: String(s.zOrderPosition) })).sort((a, b) => Number(a.z) - Number(b.z));
      });
    });

    // 7. 激活形状
    const active = await record("getActiveShape", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const sh = sheet.shapes.getActiveShape();
      sh.load("name,type,left,top,width,height");
      await context.sync();
      return { name: sh.name, type: String(sh.type), left: sh.left, top: sh.top };
    }));

    // 8. 分组 / 解组
    const group = await record("addGroup + 组合成员读回", async () => {
      await ensureRect(); await ensureText();
      return await runOn(async (context) => {
      const sheet = sheetOf(context);
      // 只组合"确实创建成功"的形状：cap08_line / cap08_svg 在部分宿主上建不出来，
      // 引用它们会把分组本身的能力误判成不支持。
      const names = ["cap08_rect", "cap08_text"];
      const proxies = names.map(n => sheet.shapes.getItem(n));
      proxies.forEach(p => p.load("id"));
      await context.sync();
      const g = sheet.shapes.addGroup(proxies.map(p => p.id));
      g.name = "cap08_group";
      g.load("name");
      const members = g.group.shapes.load("items/name,items/type");
      await context.sync();
      return { group: g.name, memberCount: members.items.length, members: members.items.map(m => m.name) };
      });
    });

    const ungroup = await record("ShapeGroup.ungroup()", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const g = sheet.shapes.getItem("cap08_group");
      g.load("name");
      const kids = g.group.shapes.load("items/name");
      await context.sync();
      const kidNames = kids.items.map(k => k.name);
      if (typeof g.group.ungroup !== "function") throw new Error("该版本没有 ShapeGroup.ungroup()");
      g.group.ungroup();
      await context.sync();
      const rest = sheet.shapes.load("items/name");
      await context.sync();
      const remaining = rest.items.map(s => s.name);
      return { released: kidNames, backInSheet: kidNames.filter(n => remaining.indexOf(n) >= 0) };
    }));

    // 9. 形状导出图片
    const image = await record("getAsImage(png)", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const img = sheet.shapes.getItem("cap08_text").getAsImage(resolvePictureFormat("png"), 1);
      await context.sync();
      return { length: img.value ? img.value.length : 0, prefix: img.value ? String(img.value).slice(0, 16) : null };
    }));

    // 10. 缩放
    const scale = await record("scaleWidth/scaleHeight", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      const sh = sheet.shapes.getItem("cap08_rect");
      sh.scaleWidth(1.5, "CurrentSize");
      sh.scaleHeight(1.5, "CurrentSize");
      sh.load("width,height");
      await context.sync();
      return { width: sh.width, height: sh.height };
    }));

    // 11. 删除
    const del = await record("delete + 读回确认", async () => await runOn(async (context) => {
      const sheet = sheetOf(context);
      sheet.shapes.getItem("cap08_text").delete();
      await context.sync();
      const rest = sheet.shapes.load("items/name");
      await context.sync();
      return { remaining: rest.items.map(s => s.name) };
    }));

    const failed = checks.filter(c => !c.ok).map(c => c.check);
    const result = {
      success: failed.length === 0,
      sheetName,
      started,
      finished: new Date().toISOString(),
      addonVersion: ADDON_VERSION,
      total: checks.length,
      passed: checks.length - failed.length,
      failed,
      checks,
      readback: { rect, line, svg: svgOk, textBox, list, zorder, active, group, ungroup, image, scale, deleteRemaining: del ? del.remaining : null }
    };

    // 把结果写回探测表：换窗格/换会话后仍能用 host=microsoft 的读工具读回来。
    try {
      await runOn(async (context) => {
        const sheet = sheetOf(context);
        const rows = [["check", "ok", "ms", "detail/error"], ["ADDON_VERSION", ADDON_VERSION, "", ""]];
        for (const c of checks) rows.push([c.check, c.ok ? "OK" : "FAIL", String(c.ms), c.ok ? JSON.stringify(c.detail) : c.error]);
        sheet.getRange("A1:D" + rows.length).values = rows;
        sheet.getRange("A1:D1").format.font.bold = true;
        sheet.getRange("A:D").format.columnWidth = 160;
        await context.sync();
      });
    } catch (e) {
      result.writebackError = String(e.message || e);
    }

    if (!keep) {
      try {
        await runOn(async (context) => {
          const sheets = context.workbook.worksheets;
          sheets.getItem(sheetName).delete();
          await context.sync();
        });
        result.cleanedUp = true;
      } catch (e) {
        result.cleanupError = String(e.message || e);
        result.cleanedUp = false;
      }
    }

    CAP08_SELF_TEST_RESULT = result;
    // 返回体保持轻量（不含工作表明细），明细可从写回的表或下次 list_shapes 读
    return {
      success: result.success,
      sheetName,
      total: result.total,
      passed: result.passed,
      failed: result.failed,
      checks: checks.map(c => ({ check: c.check, ok: c.ok, ms: c.ms, error: c.error || undefined })),
      readback: result.readback,
      cleanedUp: result.cleanedUp === true,
      finished: result.finished
    };
  }

// ── 模块: src/excel/comment.js — 审阅与批注 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1732-1762 行（原样搬迁，未改写）。
  // 11. 审阅与批注
  async function handleManageComments(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const action = params.action || "add";

      if (action === "add") {
        const targetRange = sheet.getRange(params.address || params.cellAddress || "A1");
        const comment = sheet.comments.add(targetRange, params.content || params.text || "");
        comment.load("id, content");
        await context.sync();
        return { success: true, id: comment.id, content: comment.content };
      } else if (action === "list" || action === "read") {
        const comments = sheet.comments.load("items/id, items/content, items/authorName, items/creationDate");
        await context.sync();
        return {
          success: true,
          comments: comments.items.map(c => ({ id: c.id, content: c.content, author: c.authorName, created: c.creationDate }))
        };
      } else if (action === "delete") {
        if (!params.id && !params.commentId) {
          throw new Error("删除批注必须提供 id 或 commentId：请先 list 取得批注 id");
        }
        const comment = sheet.comments.getItem(params.id || params.commentId);
        comment.delete();
        await context.sync();
        return { success: true, message: "批注已删除" };
      }
      // 不再对未知 action 静默返回成功（问题台账 ISS-92）：
      // 原来 `update_comment` 就是掉进这里——什么都没做却报 success。
      throw new Error(
        `未支持的批注操作: ${action}（可用: add / list / delete）。` +
        `若要修改批注，请先 list 取 id、delete 后再 add。`
      );
    });
  }

// ── 模块: src/excel/script.js — 动态脚本执行 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1763-1773 行（原样搬迁，未改写）。
  // 12. 动态运行任意脚本
  async function handleRunScript(params) {
    return await Excel.run(async (context) => {
      const script = params.code || params.script;
      const fn = new Function("context", "Excel", script);
      const res = await fn(context, Excel);
      await context.sync();
      return { success: true, result: res };
    });
  }

// ── 模块: src/bootstrap.js — DOM 事件绑定与启动引导（必须最后拼接） ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1774-1808 行。

  /**
   * 真正重新加载任务窗格。
   *
   * 原实现用 `window.location.reload(true)`：`true` 是已废弃的 forceGet 参数，
   * Excel for Mac 的 WKWebView 里不保证重新拉取子资源（实测点了「刷新连接」后
   * 窗格仍跑旧代码，见 docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/16-cap-ms-shapes.md）。
   * 改成"清 query 再 reload"：地址变了就必然重新请求 taskpane.html 与 taskpane.js，
   * 不依赖 forceGet 的实现细节。
   */
  function reloadTaskPane() {
    try {
      const url = new URL(window.location.href);
      url.searchParams.set("_reload", String(Date.now()));
      window.location.replace(url.toString());
    } catch (e) {
      window.location.reload();
    }
  }

  function bindDomEvents() {
    const btnReconnect = document.getElementById("btnReconnect");
    if (btnReconnect) {
      btnReconnect.addEventListener("click", () => {
        reloadTaskPane();
      });
    }

    const btnToggleLog = document.getElementById("btnToggleLog");
    const logDrawer = document.getElementById("logDrawer");
    const logArrow = document.getElementById("logArrow");
    if (btnToggleLog && logDrawer) {
      btnToggleLog.addEventListener("click", () => {
        logDrawer.classList.toggle("collapsed");
        if (logArrow) {
          logArrow.style.transform = logDrawer.classList.contains("collapsed") ? "rotate(0deg)" : "rotate(180deg)";
        }
      });
    }

    const btnClearLog = document.getElementById("btnClearLog");
    const logArea = document.getElementById("logArea");
    if (btnClearLog && logArea) {
      btnClearLog.addEventListener("click", () => {
        logArea.innerText = "";
      });
    }

    // CAP-08 形状能力自检：隐藏入口（地址带 #selftest 时自动跑，或在控制台调 window.cap08ShapeSelfTest()）。
    const btnSelfTest = document.getElementById("cap08SelfTest");
    if (btnSelfTest) {
      btnSelfTest.addEventListener("click", () => {
        window.cap08ShapeSelfTest({});
      });
    }
    window.cap08ShapeSelfTest = async function (options) {
      try {
        const result = await handleShapeSelfTest(options || {});
        console.log("[CAP-08] 形状自检结果", result);
        if (typeof log === "function") log(`CAP-08 形状自检：${result.passed}/${result.total} 通过${result.failed && result.failed.length ? "，失败：" + result.failed.join("；") : ""}`);
        return result;
      } catch (e) {
        console.error("[CAP-08] 形状自检执行失败", e);
        throw e;
      }
    };
    if (window.location.hash === "#selftest") {
      window.cap08ShapeSelfTest({});
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bindDomEvents);
  } else {
    bindDomEvents();
  }

})();
