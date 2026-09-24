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
        log("已成功连上字浮 CharFloat 服务端 (MS Office 通道)");
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

