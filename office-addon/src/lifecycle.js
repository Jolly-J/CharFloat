// ── 模块: src/lifecycle.js — Office.js 生命周期与宿主状态感知 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 57-116 行（原样搬迁，未改写）。
  // ==========================================
  // Office.js 初始化与状态感知
  // ==========================================

  Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
      if (officeRuntimeStarted) return;
      officeRuntimeStarted = true;
      log("Office.js 环境就绪 (Microsoft Excel)");
      initExcelEventHooks();
      connect();
      configureBackgroundStartup();
    } else {
      log("当前宿主不是 Excel: " + info.host);
      updateStatusUI(false, "非 Excel 环境");
    }
  });

  async function configureBackgroundStartup() {
    try {
      sharedRuntimeSupported = Boolean(Office.context.requirements &&
        Office.context.requirements.isSetSupported("SharedRuntime", "1.1") &&
        Office.addin && typeof Office.addin.setStartupBehavior === "function");
      if (!sharedRuntimeSupported) {
        backgroundStartupState = "unsupported";
      } else {
        const current = typeof Office.addin.getStartupBehavior === "function"
          ? await Office.addin.getStartupBehavior() : null;
        if (current !== Office.StartupBehavior.load) await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
        backgroundStartupState = "enabled";
        log("当前工作簿已启用后台连接，侧栏关闭后继续运行，下次打开自动启动");
      }
    } catch (error) {
      backgroundStartupState = "failed";
      log("当前工作簿的自动启动设置未完成：" + error.message);
    }
    updateBackgroundUI();
    if (isConnected) sendOfficeRegistration();
    if (sharedRuntimeSupported && typeof Office.addin.onVisibilityModeChanged === "function") {
      try {
        await Office.addin.onVisibilityModeChanged(async () => {
          await updateActiveSummary();
          updateBackgroundUI();
          if (isConnected) sendOfficeRegistration();
        });
      } catch (error) { log("侧栏显示状态监听未注册：" + error.message); }
    }
  }

  async function updateActiveSummary() {
    try {
      const docUrl = Office?.context?.document?.url;
      if (docUrl) {
        const decoded = decodeURIComponent(docUrl.split("/").pop().split("\\").pop());
        if (decoded) activeWorkbookName = decoded;
      }
      await Excel.run(async (context) => {
        const canLoadName = Office.context.requirements && Office.context.requirements.isSetSupported("ExcelApi", "1.7") && typeof context.workbook.load === "function";
        if (canLoadName) context.workbook.load("name");
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        const selection = context.workbook.getSelectedRange();
        selection.load("address");
        await context.sync();

        if (canLoadName && context.workbook.name) activeWorkbookName = context.workbook.name;
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
              documentUrl: Office.context.document.url || "",
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
