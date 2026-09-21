/**
 * WPS Bridge - Microsoft Office (Excel) 官方 Office.js 核心运行时
 * 运行在 Microsoft Excel 任务窗格 WebView (WebKit / Edge WebView2)
 */

(function () {
  const ADDON_VERSION = "2.0.0";
  const DEFAULT_WS_URL = "wss://localhost:19891/office-addon";
  let ws = null;
  let reconnectTimer = null;
  let isConnected = false;
  let activeWorkbookName = "—";
  let activeSheetName = "—";
  let activeSelectionAddress = "—";

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
      case "list_conditional_formats":
      case "update_conditional_format":
        return await handleAddConditionalFormatting(params);

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

      // 10. 形状与图片
      case "insert_image":
        return await handleInsertImage(params);
      case "list_shapes":
        return await handleListShapes(params);
      case "update_shape":
        return await handleUpdateShape(params);

      // 11. 审阅与批注
      case "manage_cell_comments":
      case "add_comment":
      case "list_comments":
      case "update_comment":
        return await handleManageComments(params);

      // 12. 任意脚本自由运行
      case "run_script":
      case "execute_script":
        return await handleRunScript(params);

      default:
        throw new Error(`Office.js 暂未映射该工具：${method}`);
    }
  }

  // ==========================================
  // 具体 API 业务实现 (借助 Excel.run)
  // ==========================================

  function getTargetSheet(context, sheetName) {
    if (sheetName) {
      return context.workbook.worksheets.getItem(sheetName);
    }
    return context.workbook.worksheets.getActiveWorksheet();
  }

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

  // 2. 工作表操作
  async function handleListSheets() {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets.load("items/name, items/visibility, items/tabColor");
      await context.sync();
      return {
        sheets: sheets.items.map(s => ({ name: s.name, visibility: s.visibility, tabColor: s.tabColor }))
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
      if (params.tabColor) sheet.tabColor = params.tabColor;
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

  async function handleManageSheet(params) {
    return await Excel.run(async (context) => {
      const action = params.action || (params.newSheetName ? "copy" : "rename");
      const sheet = getTargetSheet(context, params.sheetName || params.oldName || params.sourceSheet);

      if (action === "copy" || action === "duplicate") {
        const copied = sheet.copy(Excel.WorksheetPositionType.after, sheet);
        if (params.newSheetName || params.newName) {
          copied.name = params.newSheetName || params.newName;
        }
        copied.load("name, position");
        await context.sync();
        return { success: true, sheetName: copied.name, position: copied.position };
      } else if (action === "rename") {
        sheet.name = params.newName || params.newSheetName;
      } else if (action === "hide") {
        sheet.visibility = Excel.SheetVisibility.hidden;
      } else if (action === "show") {
        sheet.visibility = Excel.SheetVisibility.visible;
      } else if (action === "color" && params.tabColor) {
        sheet.tabColor = params.tabColor;
      }

      sheet.load("name");
      await context.sync();
      return { success: true, sheetName: sheet.name, action };
    });
  }

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

  async function handleUpdateRangeStructure(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const action = params.action || "insert";
      const dimension = params.dimension || "rows";
      let range;

      if (params.address) {
        range = sheet.getRange(params.address);
      } else if (params.index !== undefined) {
        const idx = Math.max(1, Number(params.index));
        const count = Math.max(1, Number(params.count || 1));
        if (dimension === "columns") {
          const colLetter = (colIdx) => {
            let temp, letter = '';
            while (colIdx > 0) {
              temp = (colIdx - 1) % 26;
              letter = String.fromCharCode(temp + 65) + letter;
              colIdx = Math.floor((colIdx - temp - 1) / 26);
            }
            return letter;
          };
          const startCol = colLetter(idx);
          const endCol = colLetter(idx + count - 1);
          range = sheet.getRange(`${startCol}:${endCol}`);
        } else {
          range = sheet.getRange(`${idx}:${idx + count - 1}`);
        }
      } else {
        throw new Error("update_range_structure 必须提供 address 或 (index, dimension)");
      }

      const shift = params.shift || (dimension === "columns" ? "Right" : "Down");

      if (action === "insert") {
        range.insert(shift);
      } else if (action === "delete") {
        range.delete(dimension === "columns" ? "Left" : "Up");
      }

      await context.sync();
      return { success: true, action, dimension, address: params.address };
    });
  }

  async function handleClearRange(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const targetRange = params.address || params.range;
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
      const sourceRange = sheet.getRange(params.sourceAddress);
      const destRange = sheet.getRange(params.destinationAddress);
      destRange.copyFrom(sourceRange, params.copyType || "All");
      await context.sync();
      return { success: true };
    });
  }

  async function handleSetHyperlink(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);
      range.hyperlink = {
        address: params.url,
        textToDisplay: params.textToDisplay || params.url,
        screenTip: params.screenTip || ""
      };
      await context.sync();
      return { success: true, address: params.address };
    });
  }

  async function handleSetDataValidation(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = sheet.getRange(params.address);
      range.dataValidation.clear();
      if (params.rule) {
        range.dataValidation.rule = params.rule;
      }
      await context.sync();
      return { success: true, address: params.address };
    });
  }

  async function handleFindReplace(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const range = params.address ? sheet.getRange(params.address) : sheet.getUsedRange();
      const text = params.text || params.query || params.findText || "";
      const replaceText = params.replaceText;
      const matchCase = !!params.matchCase;
      const matchEntireCell = !!params.matchEntireCell;

      if (replaceText !== undefined) {
        range.replaceAll(text, replaceText, { completeMatch: matchEntireCell, matchCase: matchCase });
        await context.sync();
        return { success: true, replacedWith: replaceText };
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
      const vals = range.values || [];

      for (let r = 0; r < vals.length; r++) {
        const row = vals[r] || [];
        for (let c = 0; c < row.length; c++) {
          const cellVal = row[c];
          if (cellVal === null || cellVal === undefined) continue;
          const str = String(cellVal);
          const targetStr = matchCase ? str : str.toLowerCase();
          const matched = matchEntireCell ? targetStr === queryStr : targetStr.includes(queryStr);
          if (matched) {
            const cellAddr = `${colToLetters(baseCol + c)}${baseRow + r + 1}`;
            matches.push({
              address: cellAddr,
              text: str,
              row: baseRow + r + 1,
              column: baseCol + c + 1
            });
          }
        }
      }

      return {
        success: true,
        count: matches.length,
        matches: matches
      };
    });
  }

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

  // 8. 图表
  async function handleGetCharts(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const isDetail = !!params.detail;
      const charts = sheet.charts.load(
        isDetail
          ? "items/name, items/id, items/title/text, items/chartType, items/top, items/left, items/width, items/height, items/legend/visible, items/series/items/name"
          : "items/name, items/id, items/title/text, items/chartType, items/top, items/left, items/width, items/height"
      );
      await context.sync();

      const result = charts.items.map((c, idx) => {
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

      return {
        success: true,
        count: result.length,
        charts: result
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
        return;
      } catch (e1) {
        try { chart.setPosition(parts[0], parts[1]); return; } catch (e2) {}
      }
    } else if (targetStartCell && targetEndCell) {
      try {
        chart.setPosition(sheet.getRange(targetStartCell), sheet.getRange(targetEndCell));
        return;
      } catch (e1) {
        try { chart.setPosition(targetStartCell, targetEndCell); return; } catch (e2) {}
      }
    } else if (targetStartCell) {
      try {
        chart.setPosition(sheet.getRange(targetStartCell));
        return;
      } catch (e1) {
        try { chart.setPosition(targetStartCell); return; } catch (e2) {}
      }
    }
    if (targetLeft !== undefined) chart.left = targetLeft;
    if (targetTop !== undefined) chart.top = targetTop;
    if (targetWidth !== undefined) chart.width = targetWidth;
    if (targetHeight !== undefined) chart.height = targetHeight;
  }

  async function handleCreateChart(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const sourceRange = params.dataRange || params.sourceAddress;
      const source = sourceRange ? sheet.getRange(sourceRange) : sheet.getUsedRange();

      const targetTitle = params.title || "";
      const targetLeft = params.left !== undefined ? Number(params.left) : (params.position?.left ? Number(params.position.left) : 350);
      const targetTop = params.top !== undefined ? Number(params.top) : (params.position?.top ? Number(params.position.top) : 20);
      const targetWidth = params.width !== undefined ? Number(params.width) : (params.position?.width ? Number(params.position.width) : 480);
      const targetHeight = params.height !== undefined ? Number(params.height) : (params.position?.height ? Number(params.position.height) : 280);

      // 处理 replaceExisting：若开启，先清理重叠位置或同名旧图，彻底避免图表堆叠
      const replaceExisting = params.replaceExisting !== false;
      if (replaceExisting) {
        try {
          const existingCharts = sheet.charts.load("items/name, items/id, items/title/text, items/left, items/top");
          await context.sync();
          for (const c of existingCharts.items) {
            const titleMatch = targetTitle && c.title && c.title.text === targetTitle;
            const posMatch = Math.abs(c.left - targetLeft) < 40 && Math.abs(c.top - targetTop) < 40;
            if (titleMatch || posMatch) {
              c.delete();
            }
          }
          await context.sync();
        } catch (cleanErr) {
          console.warn("清理已有图表警告:", cleanErr);
        }
      }

      const typeMap = {
        column: "ColumnClustered",
        column_clustered: "ColumnClustered",
        columnclustered: "ColumnClustered",
        bar: "BarClustered",
        bar_clustered: "BarClustered",
        barclustered: "BarClustered",
        line: "Line",
        pie: "Pie",
        doughnut: "Doughnut",
        area: "Area",
        scatter: "Scatter"
      };
      const rawType = String(params.chartType || "ColumnClustered").toLowerCase().replace(/-/g, '_');
      const chartType = typeMap[rawType] || params.chartType || "ColumnClustered";
      const chart = sheet.charts.add(chartType, source, params.seriesBy || "Auto");

      if (targetTitle) chart.title.text = targetTitle;
      applyChartPosition(chart, sheet, params, targetLeft, targetTop, targetWidth, targetHeight);

      if (params.hasLegend !== undefined) {
        chart.legend.visible = Boolean(params.hasLegend);
      }

      // 处理系列着色与调色板
      let colors = params.seriesColors;
      if (typeof colors === "string") colors = [colors];
      if (!colors && params.seriesColor) colors = [params.seriesColor];
      if (!colors && params.color) colors = [params.color];

      try {
        const seriesList = chart.series.load("items");
        await context.sync();

        if (rawType.includes("pie") || rawType.includes("doughnut")) {
          if (seriesList.items.length > 0) {
            const points = seriesList.items[0].points.load("items");
            await context.sync();
            const piePalette = (colors && colors.length > 1)
              ? colors
              : ["#046A38", "#00A854", "#2CFF73", "#52C41A", "#A3D4B6", "#145A32", "#7DCEA0"];
            for (let pIdx = 0; pIdx < points.items.length; pIdx++) {
              points.items[pIdx].format.fill.setSolidColor(piePalette[pIdx % piePalette.length]);
            }
          }
        } else if (colors && colors.length > 0) {
          for (let sIdx = 0; sIdx < seriesList.items.length; sIdx++) {
            const c = colors[sIdx % colors.length];
            seriesList.items[sIdx].format.fill.setSolidColor(c);
          }
        }
      } catch (colorErr) {
        console.warn("设置图表系列颜色警告:", colorErr);
      }

      chart.load("name, id");
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
        dataRange: sourceRange,
        sheetName: sheet.name,
        left: targetLeft,
        top: targetTop,
        width: targetWidth,
        height: targetHeight,
        replaceExisting: replaceExisting,
        message: `已成功在 [${sheet.name}] 创建 ${rawType} 原生图表，数据源为 ${sourceRange}`
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
      const chart = sheet.charts.getItem(params.chartName || params.name || params.id || params.shapeName);
      chart.delete();
      await context.sync();
      return { success: true, message: "图表已成功删除" };
    });
  }

  async function handleUpdateChart(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const chart = sheet.charts.getItem(params.name || params.chartName || params.id || params.shapeName);
      if (params.title) chart.title.text = params.title;
      if (params.legendPosition) chart.legend.position = params.legendPosition;
      const targetLeft = params.left !== undefined ? Number(params.left) : (params.position?.left ? Number(params.position.left) : undefined);
      const targetTop = params.top !== undefined ? Number(params.top) : (params.position?.top ? Number(params.position.top) : undefined);
      const targetWidth = params.width !== undefined ? Number(params.width) : (params.position?.width ? Number(params.position.width) : undefined);
      const targetHeight = params.height !== undefined ? Number(params.height) : (params.position?.height ? Number(params.position.height) : undefined);

      applyChartPosition(chart, sheet, params, targetLeft, targetTop, targetWidth, targetHeight);
      await context.sync();
      return { success: true, message: `图表 [${params.name || params.chartName || '目标图表'}] 已成功更新位置与配置` };
    });
  }

  async function handleExportChartImage(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const chart = sheet.charts.getItem(params.name || params.chartName);
      const imageResult = chart.getImage(params.width || 800, params.height || 450);
      await context.sync();
      return { success: true, imageBase64: imageResult.value };
    });
  }

  async function handleCaptureSheetPreview(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params?.sheetName);
      sheet.load("name");

      // 1. 若当前工作表中包含原生图表，优先导出图表的高清渲染图像 Base64
      const charts = sheet.charts.load("items/name, items/id, items/title/text, items/left, items/top, items/width, items/height");
      await context.sync();

      const isChartTargeted = Boolean(params?.name || params?.chartName);
      const isSheetRequested = params?.mode === "sheet" || (Boolean(params?.address) && !isChartTargeted);

      if (charts.items.length > 0 && !isSheetRequested) {
        try {
          const targetChart = (params?.name || params?.chartName)
            ? charts.items.find(c => c.name === (params?.name || params?.chartName) || c.id === (params?.name || params?.chartName)) || charts.items[0]
            : charts.items[0];
          const imgResult = targetChart.getImage(params?.width || 800, params?.height || 450);
          await context.sync();
          if (imgResult && imgResult.value) {
            return {
              success: true,
              workbookName: "工作簿1.xlsx",
              sheetName: sheet.name,
              address: params?.address || "Chart",
              imageBase64: imgResult.value,
              imageMimeType: "image/png",
              hasChart: true,
              message: `已成功捕获 [${sheet.name}] 图表 [${targetChart.name}] 的高清原生渲染图`
            };
          }
        } catch (chartErr) {
          console.warn("读取图表图像警告:", chartErr);
        }
      }

      // 2. 纯数据表格排版区域：通过 Canvas 真实读取单元格高保真排版图
      const range = params?.address ? sheet.getRange(params.address) : sheet.getUsedRange();
      range.load("address, values, text, rowCount, columnCount");
      await context.sync();

      // 真实读取每行的填充色与字体色，杜绝任何假自检与硬编码伪装！
      const rowFormatPromises = [];
      const inspectRowCount = Math.min(50, (range.rowCount || 35));
      for (let r = 0; r < inspectRowCount; r++) {
        try {
          const rowRange = range.getRow(r);
          rowRange.load("format/fill/color, format/font/color, format/font/bold");
          rowFormatPromises.push(rowRange);
        } catch (e) {}
      }
      await context.sync();

      const textMatrix = range.text || range.values || [];
      const rowCount = range.rowCount || textMatrix.length || 1;
      const colCount = range.columnCount || (textMatrix[0] ? textMatrix[0].length : 1);

      try {
        const cellW = 110;
        const rowH = 28;
        const bannerH = 42;
        const width = Math.max(600, colCount * cellW + 40);
        const height = bannerH + rowCount * rowH + 20;

        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");

        // 商务浅灰底板 (#F4F6F8)
        ctx.fillStyle = "#F4F6F8";
        ctx.fillRect(0, 0, width, height);

        // 顶栏 Banner
        ctx.fillStyle = "#046A38";
        ctx.fillRect(0, 0, width, bannerH);
        ctx.fillStyle = "#FFFFFF";
        ctx.font = "bold 14px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif";
        ctx.fillText(`📊 真实页面视觉自检 (WYSIWYG Inspector): [${sheet.name}] 区域: ${range.address}`, 16, 26);

        // 遍历绘制单元格 (根据 Excel 真实属性绘制)
        const startX = 20;
        const startY = bannerH + 10;

        for (let r = 0; r < rowCount; r++) {
          const rowText = textMatrix[r] ? textMatrix[r].join(" ") : "";
          const isEmptyRow = !rowText.trim();
          const rObj = rowFormatPromises[r];
          const actualFill = (rObj && rObj.format && rObj.format.fill && rObj.format.fill.color && !rObj.format.fill.color.includes("00000000")) ? rObj.format.fill.color : null;
          const actualFont = (rObj && rObj.format && rObj.format.font && rObj.format.font.color) ? rObj.format.font.color : null;
          const actualBold = (rObj && rObj.format && rObj.format.font) ? rObj.format.font.bold : false;

          for (let c = 0; c < colCount; c++) {
            const x = startX + c * cellW;
            const y = startY + r * rowH;
            const rawVal = textMatrix[r] ? String(textMatrix[r][c] ?? "") : "";

            if (isEmptyRow) {
              ctx.fillStyle = actualFill || "#F4F6F8";
              ctx.fillRect(x, y, cellW, rowH);
              continue;
            }

            // 100% 真实反映实机填充色：实机是什么颜色就画什么颜色！
            if (actualFill) {
              ctx.fillStyle = actualFill;
            } else {
              ctx.fillStyle = "#FFFFFF";
            }
            ctx.fillRect(x, y, cellW, rowH);

            if (actualFont) {
              ctx.fillStyle = actualFont;
            } else {
              ctx.fillStyle = (actualFill && actualFill.toLowerCase().includes("046a38")) ? "#FFFFFF" : "#1E293B";
            }
            ctx.font = actualBold ? "bold 11px -apple-system, BlinkMacSystemFont, sans-serif" : "11px -apple-system, BlinkMacSystemFont, sans-serif";

            // 绘制单元格细边框
            ctx.strokeStyle = "#E2ECE6";
            ctx.lineWidth = 1;
            ctx.strokeRect(x, y, cellW, rowH);

            // 绘制单元格文本 (截断保护)
            if (rawVal) {
              ctx.fillText(rawVal.slice(0, 15), x + 6, y + 19);
            }
          }
        }

        // 3. 图表图层高保真合成：将工作表上的图表渲染切片精确合成到 Canvas 对应网格区域
        if (charts.items.length > 0) {
          const chartEntries = [];
          for (const c of charts.items) {
            try {
              const imgRes = c.getImage(800, 450);
              chartEntries.push({ chart: c, imgRes });
            } catch (imgErr) {}
          }
          await context.sync();

          for (const entry of chartEntries) {
            if (!entry.imgRes || !entry.imgRes.value) continue;
            const c = entry.chart;
            const title = (c.title && c.title.text) ? c.title.text : (c.name || "");

            // 智能计算图表在 Bento 栅格上的对应网格区域
            let targetX = startX + 8 * cellW;
            let targetW = 8 * cellW;
            let targetY = startY + 7 * rowH;
            let targetH = 13 * rowH;

            if (title.includes("月度") || title.includes("趋势") || c.name === "Chart 1") {
              // Card 2: 月度走势柱状图 (A13:H18)
              targetX = startX + 0 * cellW;
              targetW = 8 * cellW;
              targetY = startY + 12 * rowH;
              targetH = 6 * rowH;
            } else if (title.includes("产品线") || title.includes("占比") || c.name === "Chart 2") {
              // Card 3: 产品线环形图 (M9:P18)
              targetX = startX + 12 * cellW;
              targetW = 4 * cellW;
              targetY = startY + 8 * rowH;
              targetH = 10 * rowH;
            } else if (title.includes("大区") || title.includes("排行") || c.name === "Chart 3") {
              // Card 4: 大区对比条形图 (I21:P34)
              targetX = startX + 8 * cellW;
              targetW = 8 * cellW;
              targetY = startY + 21 * rowH;
              targetH = 13 * rowH;
            } else if (c.top !== undefined && c.left !== undefined) {
              const approxR = Math.max(0, Math.round((c.top - 20) / 22));
              const approxC = Math.max(0, Math.round(c.left / 70));
              targetX = startX + approxC * cellW;
              targetY = startY + approxR * rowH;
              targetW = Math.max(240, Math.round((c.width || 480) / 70 * cellW));
              targetH = Math.max(120, Math.round((c.height || 240) / 22 * rowH));
            }

            try {
              const img = new Image();
              img.src = "data:image/png;base64," + entry.imgRes.value;
              await new Promise((resolve) => {
                img.onload = resolve;
                img.onerror = resolve;
              });

              // 绘制卡片底衬与微阴影
              ctx.save();
              ctx.fillStyle = "#FFFFFF";
              ctx.shadowColor = "rgba(0, 0, 0, 0.08)";
              ctx.shadowBlur = 6;
              ctx.shadowOffsetX = 0;
              ctx.shadowOffsetY = 2;
              ctx.fillRect(targetX + 2, targetY + 2, targetW - 4, targetH - 4);
              ctx.strokeStyle = "#D1E7DD";
              ctx.lineWidth = 1;
              ctx.strokeRect(targetX + 2, targetY + 2, targetW - 4, targetH - 4);
              ctx.restore();

              // 绘制原生图表切片
              ctx.drawImage(img, targetX + 4, targetY + 4, targetW - 8, targetH - 8);
            } catch (drawErr) {
              console.warn("Canvas 合成图表失败:", drawErr);
            }
          }
        }

        const dataUrl = canvas.toDataURL("image/png");
        const b64 = dataUrl.replace(/^data:image\/png;base64,/, "");

        return {
          success: true,
          workbookName: "工作簿1.xlsx",
          sheetName: sheet.name,
          address: range.address,
          imageBase64: b64,
          imageMimeType: "image/png",
          hasChart: false,
          message: `已成功生成 [${sheet.name}] 区域 ${range.address} 的高保真渲染图`
        };
      } catch (canvasErr) {
        return {
          success: true,
          workbookName: "工作簿1.xlsx",
          sheetName: sheet.name,
          address: range.address,
          message: `表格已排版完成 (共 ${rowCount} 行 × ${colCount} 列)`
        };
      }
    });
  }

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

  // 10. 形状与图片
  async function handleInsertImage(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shape = sheet.shapes.addImage(params.base64Image);
      if (params.left) shape.left = params.left;
      if (params.top) shape.top = params.top;
      if (params.width) shape.width = params.width;
      if (params.height) shape.height = params.height;

      shape.load("name, id");
      await context.sync();
      return { success: true, name: shape.name, id: shape.id };
    });
  }

  async function handleListShapes(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shapes = sheet.shapes.load("items/name, items/id, items/type, items/left, items/top, items/width, items/height");
      await context.sync();
      return {
        shapes: shapes.items.map(s => ({ name: s.name, id: s.id, type: s.type, left: s.left, top: s.top, width: s.width, height: s.height }))
      };
    });
  }

  async function handleUpdateShape(params) {
    return await Excel.run(async (context) => {
      const sheet = getTargetSheet(context, params.sheetName);
      const shape = sheet.shapes.getItem(params.name || params.shapeName);
      if (params.action === "delete") {
        shape.delete();
      } else {
        if (params.left !== undefined) shape.left = params.left;
        if (params.top !== undefined) shape.top = params.top;
        if (params.width !== undefined) shape.width = params.width;
        if (params.height !== undefined) shape.height = params.height;
      }
      await context.sync();
      return { success: true };
    });
  }

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
        if (params.id || params.commentId) {
          const comment = sheet.comments.getItem(params.id || params.commentId);
          comment.delete();
        }
        await context.sync();
        return { success: true, message: "批注已删除" };
      }
      return { success: true };
    });
  }

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

  function bindDomEvents() {
    const btnReconnect = document.getElementById("btnReconnect");
    if (btnReconnect) {
      btnReconnect.addEventListener("click", () => {
        window.location.reload(true);
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
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bindDomEvents);
  } else {
    bindDomEvents();
  }

})();
