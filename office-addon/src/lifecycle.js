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

