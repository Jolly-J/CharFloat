  // ---------------------------------------------------------------------------
  // connection.js — WebSocket 建连与握手、断线重连、报文发送与 RPC 回包
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
  function initWebSocket() {
    if (ws && (ws.readyState === WebSocket.CONNECTING || ws.readyState === WebSocket.OPEN)) {
      return;
    }

    try {
      log("正在连接本机 Bridge...");
      ws = new WebSocket(BRIDGE_WS_URL);

      ws.onopen = function () {
        isConnected = true;
        reconnectAttempts = 0;
        lastDisconnectReason = "正常连通";
        log("已成功连上 WPS Bridge 服务端");
        updateStatusUI(true, `已与本地 WPS Bridge (v${currentVersion}) 建立长连接`);

        let clientType = "wps-office-addon";
        let initialSummary = null;
        try {
          const host = detectHostComponent();
          clientType = host === "word" ? "wps-word-addon" : (host === "ppt" ? "wps-ppt-addon" : "wps-et-addon");
          const app = getApp();
          if (app) {
            initialSummary = getWorkspaceSummary(app);
          }
        } catch (summaryErr) {
          log("获取初始工作区摘要异常(已安全兜底): " + summaryErr.message);
        }

        try {
          sendPacket({
            type: "register",
            client: clientType,
            version: currentVersion,
            // 构建指纹由 scripts/build-wps-addon.mjs 注入到产物里。桥接拿它判断
            // "WPS 进程里加载的字节"是否等于磁盘/已部署的最新构建（ISS-59：部署后没重载时，
            // 磁盘是新的、进程里跑的是旧的，此前没有任何指纹能识别）。
            buildFingerprint: typeof ADDON_BUILD_FINGERPRINT === "string" ? ADDON_BUILD_FINGERPRINT : null,
            summary: initialSummary
          });
          log(`已成功发送注册报文 [${clientType}] (版本: ${currentVersion}, 构建: ${typeof ADDON_BUILD_FINGERPRINT === "string" ? ADDON_BUILD_FINGERPRINT.slice(0, 12) : "未知"})`);
        } catch (regErr) {
          log("发送注册报文失败: " + regErr.message);
        }

        try {
          hookWpsEvents();
        } catch (hookErr) {
          log("挂载事件监听异常: " + hookErr.message);
        }
      };

      ws.onmessage = function (event) {
        handleIncomingMessage(event.data);
      };

      ws.onerror = function (err) {
        lastDisconnectReason = "连接错误 (可能守护进程未启动或端口被拦截)";
        log("WebSocket 异常", err);
      };

      ws.onclose = function (evt) {
        isConnected = false;
        reconnectAttempts++;
        lastDisconnectReason = evt && evt.code === 1006
          ? "连接异常中断 (HTTP 401 拦截或网络断开)"
          : (evt && evt.reason ? evt.reason : "网络已断开");
        log(`连接断开 (code: ${evt ? evt.code : 'unknown'}, reason: ${lastDisconnectReason})，2.5 秒后自动重试...`);
        updateStatusUI(false, `等待连接到本地 Bridge (端口 ${BRIDGE_PORT})...`);
        scheduleReconnect();
      };
    } catch (e) {
      log("创建 WebSocket 异常: " + e.message);
      scheduleReconnect();
    }
  }

  function scheduleReconnect() {
    if (reconnectTimer) clearTimeout(reconnectTimer);
    reconnectTimer = setTimeout(() => {
      initWebSocket();
    }, 2500);
  }

  function sendPacket(packet) {
    try {
      if (ws && (ws.readyState === WebSocket.OPEN || ws.readyState === 1)) {
        ws.send(JSON.stringify(packet));
      }
    } catch (err) {
      log("sendPacket 异常: " + err.message);
    }
  }

  function sendRpcResponse(id, result, error) {
    sendPacket({
      type: "rpc_response",
      id: id,
      result: result,
      error: error
    });
  }
