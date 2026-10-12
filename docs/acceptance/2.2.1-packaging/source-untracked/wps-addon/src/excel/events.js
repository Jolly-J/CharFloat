  // events.js — WPS Excel 领域片段，函数与初始化从原 excel.js 搬迁。
  // 由 build-wps-addon.mjs 拼入共享 IIFE；业务函数名与声明内容保持不变。


  function hookWpsEvents() {
    if (eventsHooked) return;
    try {
      if (typeof wps !== "undefined" && wps.ApiEvent && wps.ApiEvent.AddApiEventListener) {
        wps.ApiEvent.AddApiEventListener("SheetSelectionChange", onSelectionChange);
        eventsHooked = true;
        log("已挂载 WPS 选区事件监听");
      }
    } catch (e) {
      log("挂载 ApiEventListener 提示: " + e.message);
    }
  }

  function onSelectionChange(sh, targetRange) {
    if (!isConnected) return;
    try {
      const addr = targetRange ? targetRange.Address() : "";
      if (addr && addr !== lastSelectionAddress) {
        lastSelectionAddress = addr;
        sendPacket({
          type: "event",
          event: "selection_change",
          data: {
            sheetName: sh ? sh.Name : "",
            address: addr,
            rowCount: targetRange.Rows.Count,
            columnCount: targetRange.Columns.Count
          }
        });
      }
    } catch (e) {}
  }
