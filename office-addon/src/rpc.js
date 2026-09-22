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
    DIAG.rpcSeen = (DIAG.rpcSeen || 0) + 1;
    DIAG.lastMethod = String(message.method || "");
    DIAG.seen[DIAG.lastMethod] = (DIAG.seen[DIAG.lastMethod] || 0) + 1;
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

      // 10. 形状与图片
      case "insert_image":
        return await handleInsertImage(params);
      case "list_shapes":
        return await handleListShapes(params);
      case "update_shape":
        return await handleUpdateShape(params);

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

      default:
        throw new Error(`Office.js 暂未映射该工具：${method}`);
    }
  }

