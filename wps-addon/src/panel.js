// 侧栏控制器：只读状态轮询与用户操作；RPC/WebSocket 始终由宿主入口持有。
  function panelComponent() {
    try {
      const value = new URL(window.location.href).searchParams.get("component");
      if (["excel", "word", "ppt"].includes(value)) return value;
    } catch (e) {}
    return detectHostComponent();
  }

  async function panelRequest(route, body) {
    const response = await fetch("http://127.0.0.1:" + BRIDGE_PORT + route, {
      method: body === undefined ? "GET" : "POST",
      headers: { "Authorization": "Bearer " + BRIDGE_TOKEN, "Content-Type": "application/json" },
      body: body === undefined ? undefined : JSON.stringify(body)
    });
    const result = await response.json();
    if (!response.ok || result.success === false) throw new Error(result.error || "后台请求失败");
    return result;
  }

  async function reconnectPanelCore() {
    try {
      const result = await panelRequest("/api/v1/tool/call", {
        name: "wps_reload_addon", arguments: {}, clientName: "WPS 字浮侧栏", sessionId: "wps-panel"
      });
      const data = result.data || result;
      if (data.success === false) throw new Error(data.message || "没有组件接受重连请求");
      log("已请求宿主加载项重新连接");
    } catch (error) { showNativeAlert("重新连接未完成：" + error.message); }
  }

  function initPanelUI() {
    if (window.CHARFLOAT_PANEL_STARTED) return;
    window.CHARFLOAT_PANEL_STARTED = true;
    let stopped = false, timer = null;
    const component = panelComponent();
    const setText = (id, value) => {
      const element = document.getElementById(id);
      if (element) element.innerText = value == null || value === "" ? "—" : String(value);
    };
    setText("hostText", { excel: "WPS 表格", word: "WPS 文字", ppt: "WPS 演示" }[component]);
    if (component !== "excel") {
      for (const id of ["sheetRow", "selectionRow"]) {
        const element = document.getElementById(id);
        if (element) element.hidden = true;
      }
    }
    const reconnect = document.getElementById("btnReconnect");
    if (reconnect) reconnect.addEventListener("click", window.OnActionForceReconnect);
    const close = document.getElementById("btnClose");
    if (close) close.addEventListener("click", () => {
      try {
        const api = typeof wps !== "undefined" ? wps : Application;
        let id = new URL(window.location.href).searchParams.get("paneId");
        if (!id && api.PluginStorage) id = api.PluginStorage.getItem("charfloat.taskpane." + component);
        if (!id) throw new Error("无法确定侧栏 ID，请使用侧栏右上角的关闭按钮。");
        const pane = api.GetTaskPane(Number(id));
        if (pane) pane.Visible = false;
      } catch (error) { showNativeAlert(error.message); }
    });
    window.addEventListener("beforeunload", () => { stopped = true; if (timer) clearTimeout(timer); });
    async function poll() {
      try {
        const status = await panelRequest("/api/v1/status");
        if (stopped) return;
        const host = status.components && status.components[component];
        const summary = host && host.summary || {};
        isConnected = Boolean(host && host.connected);
        updateStatusUI(isConnected, isConnected ? "后台连接保持中" : "宿主未连接，请启动字浮后台或重新连接");
        setText("docText", summary.workbookName || summary.documentName || summary.presentationName);
        setText("sheetText", summary.activeSheetName);
        setText("selectionText", summary.selection && summary.selection.address);
      } catch (error) {
        if (stopped) return;
        isConnected = false;
        updateStatusUI(false, "暂时无法访问字浮后台");
      } finally { if (!stopped) timer = setTimeout(poll, 2000); }
    }
    poll();
  }
