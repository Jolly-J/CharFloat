// 本文件由 scripts/build-wps-addon.mjs 生成，请勿手改；改动请改 wps-addon/src/**
(function () {
  // ---------------------------------------------------------------------------
  // shared.js — 配置常量与运行态变量、日志/状态 UI/原生弹窗、宿主组件探测与文档定位、颜色换算、工作区摘要
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
/**
 * WPS Bridge - WPS 内部加载项核心运行时 (专业美学与多功能版)
 * 运行在 WPS Office 进程内部 (JSA 环境)
 */

  const ADDON_VERSION = "2.1.0";
  const config = window.WPS_BRIDGE_CONFIG || {};
  const currentVersion = config.version || ADDON_VERSION;
  const BRIDGE_PORT = config.port || 19890;
  const BRIDGE_TOKEN = config.token || "";
  const BRIDGE_WS_URL = "ws://127.0.0.1:" + BRIDGE_PORT + "/addon?token=" + encodeURIComponent(BRIDGE_TOKEN);
  let ws = null;
  let reconnectTimer = null;
  let isConnected = false;
  let lastSelectionAddress = "";
  let eventsHooked = false;
  let reconnectAttempts = 0;
  let lastDisconnectReason = "待连接";
  let serverVersionNotice = null;

  function log(msg, data) {
    const timestamp = new Date().toLocaleTimeString();
    const formatted = `[${timestamp}] ${msg}` + (data ? ` ${JSON.stringify(data)}` : "");
    console.log(formatted);
    const logEl = document.getElementById("logArea");
    if (logEl) {
      logEl.innerText = (formatted + "\n" + logEl.innerText).slice(0, 3000);
    }
  }

  function updateStatusUI(connected, text) {
    const badge = document.getElementById("statusBadge");
    const detail = document.getElementById("detailInfo");
    const verEl = document.getElementById("addonVersionTag");
    if (verEl) {
      verEl.innerText = "v" + currentVersion;
    }
    if (badge) {
      badge.className = "badge " + (connected ? "connected" : "disconnected");
      badge.innerText = connected ? "已连接" : "未连接";
    }
    if (detail && text) {
      detail.innerText = text;
    }
  }

  function showNativeAlert(msg) {
    try {
      if (typeof wps !== "undefined") {
        if (typeof wps.alert === "function") {
          wps.alert(msg);
          return;
        }
        if (wps.Application && typeof wps.Application.Alert === "function") {
          wps.Application.Alert(msg);
          return;
        }
      }
    } catch (e) {}
    try {
      const app = getApp();
      if (app && typeof app.Alert === "function") {
        app.Alert(msg);
        return;
      }
    } catch (e) {}
    try {
      if (typeof alert === "function") {
        alert(msg);
        return;
      }
    } catch (e) {}
    try {
      if (typeof wps !== "undefined" && typeof wps.ShowDialog === "function") {
        wps.ShowDialog(msg);
        return;
      }
    } catch (e) {}
    log("[Alert]", msg);
  }


  function getWordApp() {
    try {
      if (typeof wps !== "undefined") {
        if (typeof wps.WpsApplication === "function") {
          const w = wps.WpsApplication();
          if (w) return w;
        }
        if (wps.Application && (wps.Application.Documents || wps.Application.ActiveDocument)) {
          return wps.Application;
        }
      }
    } catch (e) {}
    try {
      if (typeof Application !== "undefined" && (Application.Documents || Application.ActiveDocument)) {
        return Application;
      }
    } catch (e) {}
    return null;
  }

  function getPptApp() {
    try {
      if (typeof wps !== "undefined") {
        if (typeof wps.WppApplication === "function") {
          const p = wps.WppApplication();
          if (p) return p;
        }
        if (wps.Application && (wps.Application.Presentations || wps.Application.ActivePresentation)) {
          return wps.Application;
        }
      }
    } catch (e) {}
    try {
      if (typeof Application !== "undefined" && (Application.Presentations || Application.ActivePresentation)) {
        return Application;
      }
    } catch (e) {}
    return null;
  }

  function getEtApp() {
    try {
      if (typeof wps !== "undefined") {
        if (typeof wps.EtApplication === "function") {
          const e = wps.EtApplication();
          if (e) return e;
        }
        if (wps.Application && (wps.Application.Workbooks || wps.Application.ActiveWorkbook)) {
          return wps.Application;
        }
      }
    } catch (e) {}
    try {
      if (typeof Application !== "undefined" && (Application.Workbooks || Application.ActiveWorkbook)) {
        return Application;
      }
    } catch (e) {}
    return null;
  }

  function detectHostComponent() {
    try {
      const w = getWordApp();
      if (w && (w.Documents || w.ActiveDocument)) return "word";
    } catch (e) {}
    try {
      const p = getPptApp();
      if (p && (p.Presentations || p.ActivePresentation)) return "ppt";
    } catch (e) {}
    try {
      const e = getEtApp();
      if (e && (e.Workbooks || e.ActiveWorkbook)) return "excel";
    } catch (e) {}
    try {
      if (typeof Application !== "undefined") {
        if (Application.Documents || Application.ActiveDocument) return "word";
        if (Application.Presentations || Application.ActivePresentation) return "ppt";
        if (Application.Workbooks || Application.ActiveWorkbook) return "excel";
      }
    } catch (e) {}
    return "excel";
  }

  function getApp() {
    const host = detectHostComponent();
    if (host === "word") return getWordApp() || (typeof Application !== "undefined" ? Application : null);
    if (host === "ppt") return getPptApp() || (typeof Application !== "undefined" ? Application : null);
    return getEtApp() || (typeof Application !== "undefined" ? Application : null);
  }

  const lockedTargets = {
    word: null,
    excel: null,
    ppt: null
  };

  function getWordDocument(app, docName) {
    const wordApp = getWordApp() || getWpsApp() || app || (typeof Application !== "undefined" ? Application : null);
    if (!wordApp) throw new Error("WPS 文字 (Word) 未就绪或未打开任何文档");

    const targetName = docName || lockedTargets.word;
    if (targetName) {
      try {
        const target = wordApp.Documents.Item(targetName);
        if (target) return target;
      } catch (e) {}
      try {
        const count = wordApp.Documents.Count;
        for (let i = 1; i <= count; i++) {
          const d = wordApp.Documents.Item(i);
          if (d.Name === targetName || d.FullName === targetName || false) return d;
        }
      } catch (e) {}
      const openList = [];
      try {
        for (let i = 1; i <= wordApp.Documents.Count; i++) openList.push(wordApp.Documents.Item(i).Name);
      } catch (e) {}
      throw new Error(`未在 WPS 中找到目标 Word 文档 [${targetName}]。当前已打开: ${openList.join(", ") || "无"}`);
    }

    if (wordApp.Documents && wordApp.Documents.Count === 1) {
      const single = wordApp.Documents.Item(1);
      lockedTargets.word = single.Name;
      return single;
    }

    if (wordApp.ActiveDocument) return wordApp.ActiveDocument;
    if (wordApp.Documents && wordApp.Documents.Count > 0) return wordApp.Documents.Item(1);
    throw new Error("当前未打开任何 Word 文档，请先在 WPS 文字中打开目标文档");
  }

  function getPptPresentation(app, presName) {
    const pptApp = getPptApp() || app || (typeof Application !== "undefined" ? Application : null);
    if (!pptApp) throw new Error("WPS 演示 (PowerPoint) 未就绪或未打开任何文稿");

    const targetName = presName || lockedTargets.ppt;
    if (targetName) {
      try {
        const target = pptApp.Presentations.Item(targetName);
        if (target) return target;
      } catch (e) {}
      try {
        const count = pptApp.Presentations.Count;
        for (let i = 1; i <= count; i++) {
          const p = pptApp.Presentations.Item(i);
          if (p.Name === targetName || p.FullName === targetName || false) return p;
        }
      } catch (e) {}
      const openList = [];
      try {
        for (let i = 1; i <= pptApp.Presentations.Count; i++) openList.push(pptApp.Presentations.Item(i).Name);
      } catch (e) {}
      throw new Error(`未在 WPS 中找到目标演示文稿 [${targetName}]。当前已打开: ${openList.join(", ") || "无"}`);
    }

    if (pptApp.Presentations && pptApp.Presentations.Count === 1) {
      const single = pptApp.Presentations.Item(1);
      lockedTargets.ppt = single.Name;
      return single;
    }

    if (pptApp.ActivePresentation) return pptApp.ActivePresentation;
    if (pptApp.Presentations && pptApp.Presentations.Count > 0) return pptApp.Presentations.Item(1);
    throw new Error("当前未打开任何 PowerPoint 演示文稿");
  }

  // 获取目标 Workbook (支持显式名称寻址或兜底 ActiveWorkbook)
  function getWorkbook(app, workbookName) {
    const excelApp = getEtApp() || app || (typeof Application !== "undefined" ? Application : null);
    if (!excelApp) throw new Error("WPS 表格 (Excel) 未就绪或未打开任何工作簿");

    const targetName = workbookName || lockedTargets.excel;
    if (targetName) {
      try {
        const targetWb = excelApp.Workbooks.Item(targetName);
        if (targetWb) return targetWb;
      } catch (e) {}
      try {
        const count = excelApp.Workbooks.Count;
        for (let i = 1; i <= count; i++) {
          const wb = excelApp.Workbooks.Item(i);
          if (wb.Name === targetName || wb.FullName === targetName || false) {
            return wb;
          }
        }
      } catch (e) {}
      const openList = [];
      try {
        for (let i = 1; i <= excelApp.Workbooks.Count; i++) openList.push(excelApp.Workbooks.Item(i).Name);
      } catch (e) {}
      throw new Error(`未在 WPS 中找到目标工作簿 [${targetName}]。当前已打开: ${openList.join(", ") || "无"}`);
    }

    if (excelApp.Workbooks && excelApp.Workbooks.Count === 1) {
      const single = excelApp.Workbooks.Item(1);
      lockedTargets.excel = single.Name;
      return single;
    }

    const wb = excelApp.ActiveWorkbook;
    if (!wb) throw new Error("当前 WPS 中没有打开的工作簿");
    return wb;
  }

  function getWorksheet(app, sheetName, workbookName) {
    const wb = getWorkbook(app, workbookName);
    if (!sheetName) return wb.ActiveSheet;
    try {
      const sh = wb.Worksheets.Item(sheetName);
      if (sh) return sh;
    } catch (e) {}

    let availableSheets = [];
    try {
      const count = wb.Worksheets.Count;
      for (let i = 1; i <= count; i++) {
        availableSheets.push(wb.Worksheets.Item(i).Name);
      }
    } catch (e) {}

    throw new Error(`在工作簿 [${wb.Name}] 中找不到工作表: "${sheetName}"。当前可用的工作表为: [${availableSheets.join(", ")}]。若需要创建新表，请先调用 wps_create_sheet。`);
  }

  // 十六进制颜色转 Excel BGR 整数
  function hexToExcelColor(hex) {
    if (!hex) return null;
    const cleanHex = hex.replace("#", "");
    if (cleanHex.length !== 6) return null;
    const r = parseInt(cleanHex.substring(0, 2), 16);
    const g = parseInt(cleanHex.substring(2, 4), 16);
    const b = parseInt(cleanHex.substring(4, 6), 16);
    return (b << 16) | (g << 8) | r;
  }

  // Excel BGR 整数转十六进制
  function excelColorToHex(num) {
    if (num === undefined || num === null || num < 0) return null;
    const r = num & 0xff;
    const g = (num >> 8) & 0xff;
    const b = (num >> 16) & 0xff;
    return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
  }



  // 1. 获取工作区总览 (支持 Word、PPT、Excel 智能自适应)
  function getWorkspaceSummary(app, workbookName) {
    const hostType = detectHostComponent();

    // 1. 如果处于 Word (文字) 环境
    if (hostType === "word") {
      try {
        const doc = getWordDocument(app);
        return {
          hostType: "word",
          hasOpenDocument: Boolean(doc),
          hasOpenWorkbook: false,
          documentName: doc ? doc.Name : "未命名文档",
          fullName: doc ? (doc.FullName || doc.Name) : "",
          paragraphCount: doc && doc.Paragraphs ? doc.Paragraphs.Count : 0,
          tableCount: doc && doc.Tables ? doc.Tables.Count : 0,
          wordCount: doc && doc.Words ? doc.Words.Count : 0,
          message: doc ? `当前已打开 Word 文档: ${doc.Name}` : "当前没有打开的 Word 文档"
        };
      } catch (e) {
        return { hostType: "word", hasOpenDocument: false, hasOpenWorkbook: false, message: e.message || "当前没有打开的 Word 文档" };
      }
    }

    // 2. 如果处于 PPT (演示) 环境
    if (hostType === "ppt") {
      try {
        const pres = getPptPresentation(app);
        return {
          hostType: "ppt",
          hasOpenPresentation: Boolean(pres),
          hasOpenWorkbook: false,
          presentationName: pres ? pres.Name : "未命名演示文稿",
          fullName: pres ? (pres.FullName || pres.Name) : "",
          slideCount: pres && pres.Slides ? pres.Slides.Count : 0,
          message: pres ? `当前已打开 PPT 演示文稿: ${pres.Name}` : "当前没有打开的 PPT 演示文稿"
        };
      } catch (e) {
        return { hostType: "ppt", hasOpenPresentation: false, hasOpenWorkbook: false, message: e.message || "当前没有打开的 PPT 演示文稿" };
      }
    }

    // 3. 默认处于 Excel (表格) 环境
    try {
      const wb = getWorkbook(app, workbookName);
      if (!wb) {
        return {
          hostType: "excel",
          hasOpenWorkbook: false,
          message: "当前没有打开的 Excel 表格"
        };
      }

      const openWorkbooks = [];
      try {
        const wbCount = app.Workbooks.Count;
        for (let i = 1; i <= wbCount; i++) {
          const w = app.Workbooks.Item(i);
          openWorkbooks.push({
            name: w.Name,
            fullName: w.FullName || w.Name,
            isActive: app.ActiveWorkbook ? app.ActiveWorkbook.Name === w.Name : false
          });
        }
      } catch (e) {}

      const sheets = [];
      const count = wb.Worksheets.Count;
      for (let i = 1; i <= count; i++) {
        const sheet = wb.Worksheets.Item(i);
        sheets.push({
          index: i,
          name: sheet.Name,
          visible: sheet.Visible === -1
        });
      }

      const activeSheet = wb.ActiveSheet;
      let selectionInfo = null;
      try {
        const sel = app.Selection;
        if (sel && sel.Address) {
          selectionInfo = {
            address: sel.Address(),
            rowCount: sel.Rows.Count,
            columnCount: sel.Columns.Count
          };
        }
      } catch (e) {}

      return {
        hostType: "excel",
        hasOpenWorkbook: true,
        workbookName: wb.Name,
        fullName: wb.FullName || wb.Name,
        openWorkbooks: openWorkbooks,
        activeSheetName: activeSheet ? activeSheet.Name : "",
        sheetCount: count,
        sheets: sheets,
        selection: selectionInfo
      };
    } catch (e) {
      return {
        hostType: "excel",
        hasOpenWorkbook: false,
        message: e.message || "当前没有打开的 Excel 表格"
      };
    }
  }

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
            summary: initialSummary
          });
          log(`已成功发送注册报文 [${clientType}] (版本: ${currentVersion})`);
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

  // ---------------------------------------------------------------------------
  // ribbon.js — Ribbon 全局回调 window.OnAction* 与剪贴板工具
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
  // ==========================================
  // Ribbon 全局回调顶置与原生状态呈现
  // ==========================================

  window.OnActionBridgeStatus = function () {
    try {
      const app = getApp();
      const host = detectHostComponent();
      const summary = app ? getWorkspaceSummary(app) : null;
      let msg = "【Office Agent Bridge 运行状态】\n\n";
      msg += `插件版本: v${currentVersion} (协议 v2)\n`;
      msg += `通信状态: ${isConnected ? `已连接 (127.0.0.1:${BRIDGE_PORT})` : `未连接 (${lastDisconnectReason})`}\n`;
      msg += `网关配置: 127.0.0.1:${BRIDGE_PORT} (Token: ${BRIDGE_TOKEN ? "已配置" : "缺失"})\n`;
      if (serverVersionNotice) {
        msg += `版本提示: 发现新版本 v${serverVersionNotice.latestVersion} (建议从客户端升级)\n`;
      }
      msg += `当前组件: ${host === "word" ? "WPS 文字 (Word)" : (host === "ppt" ? "WPS 演示 (PowerPoint)" : "WPS 表格 (Excel)")}\n\n`;

      if (host === "word") {
        if (summary && summary.hasOpenDocument) {
          msg += `当前文档: ${summary.documentName}\n`;
          msg += `段落总数: ${summary.paragraphCount} 段\n`;
          msg += `表格总数: ${summary.tableCount} 个\n`;
          msg += `字数统计: ${summary.wordCount} 字\n`;
        } else {
          msg += "当前未检测到打开的 Word 文档\n";
        }
      } else if (host === "ppt") {
        if (summary && summary.hasOpenPresentation) {
          msg += `当前演示文稿: ${summary.presentationName}\n`;
          msg += `幻灯片总数: ${summary.slideCount} 页\n`;
        } else {
          msg += "当前未检测到打开的 PPT 演示文稿\n";
        }
      } else {
        if (summary && summary.hasOpenWorkbook) {
          msg += `当前工作簿: ${summary.workbookName}\n`;
          msg += `活动工作表: ${summary.activeSheetName}\n`;
          if (summary.selection) {
            msg += `鼠标光标选区: ${summary.selection.address} (${summary.selection.rowCount}行 × ${summary.selection.columnCount}列)\n`;
          }
        } else {
          msg += "当前未检测到打开的 Excel 表格文件\n";
        }
      }

      if (!isConnected) {
        msg += "\n排查建议:\n1. 确认 Office Agent Bridge 桌面客户端已启动后台服务；\n2. 若长期无法连通，可在客户端点击【安装 / 升级加载项】。";
      }
      showNativeAlert(msg);
    } catch (e) {
      showNativeAlert("获取 Bridge 状态异常: " + (e.message || String(e)));
    }
  };

  window.OnActionForceReconnect = function () {
    try {
      log("用户点击重新连接...");
      if (ws) {
        try { ws.close(); } catch (e) {}
      }
      initWebSocket();
      window.location.reload();
    } catch (e) {
      showNativeAlert("重连异常: " + (e.message || String(e)));
    }
  };

  window.OnRibbonLoaded = function (ribbonUI) {
    try {
      log("WPS 功能区载入，自动激活连接...");
      initWebSocket();
    } catch (e) {}
  };

  window.OnGetImage = function (control) {
    try {
      const id = typeof control === "object" && control ? (control.Id || control.id) : String(control);
      if (id === "btnBrandHero") {
        return "logo.png";
      }
      if (id === "btnAutoFitFormat") {
        return "table-format.png";
      }
    } catch (e) {
      log("[OnGetImage Error]", e.message);
    }
    return "";
  };

  window.OnActionUndoAiAction = async function () {
    try {
      if (detectHostComponent() !== "excel") throw new Error("审计撤销仅覆盖表格单元格的值与公式。");
      const summary = getWorkspaceSummary(getApp());
      if (!summary.workbookName) throw new Error("请先打开目标工作簿。");
      const call = async (name, args) => {
        const response = await fetch("http://127.0.0.1:" + (config.port || 19890) + "/api/v1/tool/call", {
          method: "POST", headers: { "Content-Type": "application/json", "Authorization": "Bearer " + config.token },
          body: JSON.stringify({ name: name, arguments: args, clientName: "WPS Ribbon", sessionId: "wps-ribbon" })
        });
        const result = await response.json();
        if (!response.ok || result.success === false) throw new Error(result.error || "Bridge 请求失败");
        return result.data;
      };
      const records = await call("wps_get_audit_history", { workbookName: summary.workbookName, sheetName: summary.activeSheetName, status: "applied", limit: 20 });
      const record = records.find(item => !item.host || item.host === "wps");
      if (!record) { showNativeAlert("当前工作表没有可撤销的单元格修改记录。"); return; }
      if (!confirm("撤销此修改？\n" + record.description + "\n" + record.workbookName + " / " + record.sheetName + " / " + record.address + "\n有后续修改时会拒绝覆盖。")) return;
      await call("wps_rollback", { auditId: record.id });
      showNativeAlert("已恢复记录中的单元格值与公式。");
    } catch (error) { showNativeAlert("撤销未完成：" + error.message); }
  };

  function copyTextToClipboard(text) {
    try {
      if (typeof navigator !== "undefined" && navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text);
        return true;
      }
    } catch (e) {}
    try {
      if (typeof document !== "undefined") {
        const textArea = document.createElement("textarea");
        textArea.value = text;
        textArea.style.position = "fixed";
        textArea.style.opacity = "0";
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand("copy");
        document.body.removeChild(textArea);
        return true;
      }
    } catch (e) {}
    return false;
  }

  window.OnActionOpenDesktopApp = function () {
    try {
      showNativeAlert("【Office Agent Bridge 控制中心】\n\n请在屏幕顶部菜单栏或程序坞中切换至「Office Agent Bridge」桌面管理窗口。\n\n您可以在管理中心中配置 AI 助手（豆包、Kimi、Claude、WorkBuddy 等）、查看单元格修改快照以及进行系统诊断。");
    } catch (e) {
      showNativeAlert("打开控制中心提示: " + (e.message || String(e)));
    }
  };

  window.OnActionShowAuditHistory = async function () {
    try {
      if (detectHostComponent() !== "excel") {
        showNativeAlert("修改历史清单功能目前主要面向表格（Excel/WPS表格）单元格变更。");
        return;
      }
      const summary = getWorkspaceSummary(getApp());
      if (!summary.workbookName) throw new Error("请先打开目标工作簿。");
      const response = await fetch("http://127.0.0.1:" + (config.port || 19890) + "/api/v1/tool/call", {
        method: "POST", headers: { "Content-Type": "application/json", "Authorization": "Bearer " + config.token },
        body: JSON.stringify({ name: "wps_get_audit_history", arguments: { workbookName: summary.workbookName, limit: 10 }, clientName: "WPS Ribbon", sessionId: "wps-ribbon" })
      });
      const result = await response.json();
      const records = result?.data || [];
      if (!records.length) {
        showNativeAlert("【安全时光机 · 历史清单】\n\n当前工作簿尚未产生 AI 修改记录。\n当 AI 智能体执行表格写入时，系统将自动记录前后快照。");
        return;
      }
      let msg = `【安全时光机 · 最近 ${records.length} 条修改快照】\n\n`;
      records.forEach((r, idx) => {
        const time = r.timestamp ? new Date(r.timestamp).toLocaleTimeString() : "刚刚";
        const statusText = r.status === "rolled_back" ? "[已撤销]" : "[有效]";
        msg += `${idx + 1}. [${time}] ${statusText} ${r.sheetName || "Sheet1"}!${r.address || "区域"}\n   说明: ${r.description || "单元格写入"}\n`;
      });
      msg += "\n如需撤销最近修改，可直接点击功能区【撤销 AI 修改】。";
      showNativeAlert(msg);
    } catch (e) {
      showNativeAlert("获取修改历史异常: " + (e.message || String(e)));
    }
  };

  window.OnActionToggleSheetLock = function () {
    try {
      const app = getApp();
      if (!app) throw new Error("未检测到活动办公应用");
      const host = detectHostComponent();
      if (host === "excel") {
        const sheet = app.ActiveSheet;
        if (!sheet) throw new Error("未检测到活动工作表");
        if (sheet.ProtectContents) {
          sheet.Unprotect();
          showNativeAlert("【工作表安全锁已解除】\n\n当前工作表已允许编辑与 AI 智能体写入。");
        } else {
          sheet.Protect();
          showNativeAlert("【工作表已锁定保护】\n\n已开启工作表防误改保护！AI 智能体尝试写入时将受到安全防护拒绝。");
        }
      } else {
        showNativeAlert("锁定保护功能目前优先适配表格工作表。");
      }
    } catch (e) {
      showNativeAlert("锁定操作异常: " + (e.message || String(e)));
    }
  };

  window.OnActionAutoFitFormat = function () {
    try {
      const app = getApp();
      if (!app) throw new Error("未检测到活动办公应用");
      const host = detectHostComponent();
      if (host === "excel") {
        const sheet = app.ActiveSheet;
        if (!sheet) throw new Error("未检测到活动工作表");
        if (sheet.UsedRange && sheet.UsedRange.Columns) {
          sheet.UsedRange.Columns.AutoFit();
          showNativeAlert("【自适应排版完成】\n\n已自动根据内容长度优化并自适应整张工作表的所有列宽！");
        } else {
          showNativeAlert("当前工作表没有可用数据区域。");
        }
      } else {
        showNativeAlert("自适应排版功能优先适配表格（Excel）排版。");
      }
    } catch (e) {
      showNativeAlert("自适应排版失败: " + (e.message || String(e)));
    }
  };

  window.OnActionCopyAsMarkdown = function () {
    try {
      const app = getApp();
      if (!app) throw new Error("未检测到活动办公应用");
      const host = detectHostComponent();
      if (host === "excel") {
        const sel = app.Selection;
        if (!sel) throw new Error("请先用鼠标框选需要复制的单元格区域。");
        const rowCount = sel.Rows.Count;
        const colCount = sel.Columns.Count;
        if (rowCount === 0 || colCount === 0) throw new Error("选区为空");
        
        let md = "";
        for (let r = 1; r <= rowCount; r++) {
          let rowCells = [];
          for (let c = 1; c <= colCount; c++) {
            const cell = sel.Cells.Item(r, c);
            const val = cell.Text || cell.Value2 || "";
            rowCells.push(String(val).replace(/\|/g, "\\|").replace(/\n/g, " "));
          }
          md += "| " + rowCells.join(" | ") + " |\n";
          if (r === 1) {
            md += "| " + rowCells.map(() => "---").join(" | ") + " |\n";
          }
        }
        copyTextToClipboard(md);
        showNativeAlert(`【选区已复制为 Markdown】\n\n已成功将 ${rowCount} 行 × ${colCount} 列数据格式化为 Markdown 表格并写入剪贴板！\n可直接在任意 AI 助手（豆包、Kimi、ChatGPT 等）对话框中按 Ctrl+V / Cmd+V 粘贴提问。`);
      } else {
        showNativeAlert("复制为 Markdown 功能目前主要面向表格数据选区。");
      }
    } catch (e) {
      showNativeAlert("复制选区失败: " + (e.message || String(e)));
    }
  };

  window.OnActionClearEmptyRows = function () {
    try {
      showNativeAlert("【清除冗余空白】\n\n建议直接选中需要清理的行或列，按键盘 Delete 清除，或在 AI 助手中输入：“帮我检查并清理本表中的空行空列”。");
    } catch (e) {}
  };

  window.OnActionPromptFinance = function () {
    const prompt = "请分析当前打开的财务/业务数据表格，从核心营收、同比环比、毛利率以及异常波动点进行深度专业洞察，并指出潜在的经营风险与优化建议。";
    copyTextToClipboard(prompt);
    showNativeAlert("【已复制财务分析提示词】\n\n「" + prompt + "」\n\n提示词已复制到剪贴板，可直接粘贴发送给您的 AI 助手！");
  };

  window.OnActionPromptFormula = function () {
    const prompt = "请帮我检查当前表格中计算公式的逻辑与引用范围，指出潜在的 #N/A 或循环引用错误，并给出最简洁优雅的修复公式建议。";
    copyTextToClipboard(prompt);
    showNativeAlert("【已复制公式排错提示词】\n\n「" + prompt + "」\n\n提示词已复制到剪贴板，可直接粘贴发送给您的 AI 助手！");
  };

  window.OnActionPromptCleaning = function () {
    const prompt = "请帮我清洗当前表格数据：规范日期与手机号格式，剔除前后不可见空格，识别并标记重复项与缺失值。";
    copyTextToClipboard(prompt);
    showNativeAlert("【已复制数据清洗提示词】\n\n「" + prompt + "」\n\n提示词已复制到剪贴板，可直接粘贴发送给您的 AI 助手！");
  };

  window.OnActionPromptSummary = function () {
    const prompt = "请快速提炼当前文档/表格的核心关键数据与结论，按照高管汇报要点梳理出 3 条核心要点与下一步行动建议。";
    copyTextToClipboard(prompt);
    showNativeAlert("【已复制提炼摘要提示词】\n\n「" + prompt + "」\n\n提示词已复制到剪贴板，可直接粘贴发送给您的 AI 助手！");
  };

  window.OnActionShowGuide = function () {
    showNativeAlert("【Office Agent (AI) 快速使用指南】\n\n1. 确保 Office Agent Bridge 桌面客户端处于「正常运行」状态；\n2. 在桌面端「AI 助手授权中心」一键绑定您的常用客户端（豆包 / Kimi / Claude / WorkBuddy 等）；\n3. 在 AI 客户端中直接对话即可实时读取、分析并修改当前打开的表格与文档！\n4. 任何时候均可点击上方【撤销 AI 修改】秒级恢复数据。");
  };

  // ---------------------------------------------------------------------------
  // dispatch.js — RPC 报文解析与方法分发（handleIncomingMessage）
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
  async function handleIncomingMessage(raw) {
    let packet;
    try {
      packet = JSON.parse(raw);
    } catch (e) {
      log("无法解析报文: " + raw);
      return;
    }

    if (packet && packet.type === "version_notice") {
      serverVersionNotice = packet;
      log(`收到服务端版本协商提醒: ${packet.message} (最新可用: v${packet.latestVersion})`);
      const updateTipEl = document.getElementById("updateTip");
      if (updateTipEl) {
        updateTipEl.style.display = "block";
        updateTipEl.innerText = `💡 提示: 服务端检测到加载项有新版本 v${packet.latestVersion}，可从客户端升级`;
      }
      return;
    }

    const { id, method, params } = packet;
    log(`收到 RPC 调用 [${method}], id: ${id}`);

    const app = getApp();
    if (!app) {
      sendRpcResponse(id, null, "WPS 宿主对象未就绪");
      return;
    }

    try {
      let result = null;
      switch (method) {
        case "ping":
          result = { pong: true, time: Date.now() };
          break;
        case "get_workspace_summary":
          result = getWorkspaceSummary(app, params?.workbookName);
          break;
        case "get_sheet_outline":
          result = getSheetOutline(app, params);
          break;
        case "get_style_token":
          result = getStyleToken(app, params);
          break;
        case "create_sheet":
          result = createWorksheet(app, params);
          break;
        case "delete_sheet":
          result = deleteWorksheet(app, params);
          break;
        case "clear_range":
          result = clearRange(app, params);
          break;
        case "read_range":
          result = readRangeData(app, params);
          break;
        case "get_range_styles":
          result = getRangeStyles(app, params);
          break;
        case "search_cells":
          result = searchCells(app, params);
          break;
        case "patch_cells":
          result = patchCells(app, params);
          break;
        case "format_cells":
          result = formatCells(app, params);
          break;
        case "add_conditional_formatting":
          result = addConditionalFormatting(app, params);
          break;
        case "freeze_panes":
          result = freezePanes(app, params);
          break;
        case "modify_rows_columns":
          result = modifyRowsColumns(app, params);
          break;
        case "auto_fit_columns":
          result = autoFitColumns(app, params);
          break;
        case "rollback_cells":
          result = rollbackCells(app, params);
          break;
        case "insert_dimension":
          result = insertDimension(app, params);
          break;
        case "capture_sheet_preview":
          result = captureSheetPreview(app, params);
          break;
        case "add_chart":
          result = addChart(app, params);
          break;
        case "get_charts":
          result = getCharts(app, params);
          break;
        case "delete_chart":
          result = deleteChart(app, params);
          break;
        case "create_pivot_table":
          result = createPivotTable(app, params);
          break;
        case "set_filter_and_sort":
          result = setFilterAndSort(app, params);
          break;
        case "set_data_validation":
          result = setDataValidation(app, params);
          break;
        case "manage_sheet":
          result = manageSheet(app, params);
          break;
        case "manage_rows_and_columns":
          result = modifyRowsColumns(app, params);
          break;
        case "manage_cell_comments":
          result = manageCellComments(app, params);
          break;
        case "find_and_replace":
          result = findAndReplace(app, params);
          break;
        case "duplicate_sheet":
          result = duplicateSheet(app, params);
          break;
        case "save_workbook":
          result = saveWorkbook(app, params);
          break;

        // Word (文字) RPC 分发
        case "word_create_document":
          result = wordCreateDocument(app, params);
          break;
        case "word_save_document":
          result = wordSaveDocument(app, params);
          break;
        case "word_close_document":
          result = wordCloseDocument(app, params);
          break;
        case "word_manage_content":
          result = wordManageContent(app, params);
          break;
        case "word_read_document":
          result = wordReadDocument(app, params);
          break;
        case "word_write_content":
          result = wordWriteContent(app, params);
          break;
        case "word_format_document":
          result = wordFormatDocument(app, params);
          break;
        case "word_insert_table_of_contents":
          result = wordInsertTableOfContents(app, params);
          break;
        case "word_manage_table":
          result = wordManageTable(app, params);
          break;
        case "word_review_and_comments":
          result = wordReviewAndComments(app, params);
          break;
        case "word_page_layout_and_watermark":
          result = wordPageLayoutAndWatermark(app, params);
          break;
        case "word_find_and_replace":
          result = wordFindAndReplace(app, params);
          break;
        case "word_capture_preview":
          result = wordCapturePreview(app, params);
          break;

        // PowerPoint (演示) RPC 分发
        case "ppt_read_presentation":
          result = pptReadPresentation(app, params);
          break;
        case "ppt_get_slide_shapes":
          result = pptGetSlideShapes(app, params);
          break;
        case "ppt_generate_deck":
          result = pptGenerateDeck(app, params);
          break;
        case "ppt_manage_slides":
          result = pptManageSlides(app, params);
          break;
        case "ppt_manage_table":
          result = pptManageTable(app, params);
          break;
        case "ppt_add_business_cards":
          result = pptAddBusinessCards(app, params);
          break;
        case "ppt_insert_native_chart":
          result = pptInsertNativeChart(app, params);
          break;
        case "ppt_manage_shapes_and_media":
          result = pptManageShapesAndMedia(app, params);
          break;
        case "ppt_capture_slide_preview":
          result = pptCaptureSlidePreview(app, params);
          break;

        // 文档强隔离锁定 RPC 分发
        case "lock_target_document": {
          const { component, targetName } = params || {};
          if (component && lockedTargets.hasOwnProperty(component)) {
            lockedTargets[component] = targetName ? String(targetName).trim() : null;
          }
          result = {
            success: true,
            component,
            lockedTarget: lockedTargets[component],
            allLocks: lockedTargets,
            message: `WPS 内部已锁定 ${component.toUpperCase()} 文档为 [${lockedTargets[component]}]`
          };
          break;
        }
        case "unlock_target_document": {
          const { component } = params || {};
          if (component && lockedTargets.hasOwnProperty(component)) {
            lockedTargets[component] = null;
          } else {
            lockedTargets.word = null;
            lockedTargets.excel = null;
            lockedTargets.ppt = null;
          }
          result = { success: true, allLocks: lockedTargets, message: "WPS 内部已释放锁定" };
          break;
        }
        case "get_locked_status": {
          result = {
            locks: lockedTargets,
            openPptPresentations: (() => {
              try {
                const pptApp = getPptApp();
                const list = [];
                if (pptApp && pptApp.Presentations) {
                  for (let i = 1; i <= pptApp.Presentations.Count; i++) {
                    list.push(pptApp.Presentations.Item(i).Name);
                  }
                }
                return list;
              } catch (e) { return []; }
            })(),
            openWordDocuments: (() => {
              try {
                const wordApp = getWpsApp() || getWordApp();
                const list = [];
                if (wordApp && wordApp.Documents) {
                  for (let i = 1; i <= wordApp.Documents.Count; i++) {
                    list.push(wordApp.Documents.Item(i).Name);
                  }
                }
                return list;
              } catch (e) { return []; }
            })(),
            openExcelWorkbooks: (() => {
              try {
                const etApp = getEtApp();
                const list = [];
                if (etApp && etApp.Workbooks) {
                  for (let i = 1; i <= etApp.Workbooks.Count; i++) {
                    list.push(etApp.Workbooks.Item(i).Name);
                  }
                }
                return list;
              } catch (e) { return []; }
            })()
          };
          break;
        }

        case "eval":
        case "eval_code":
        case "execute_script": {
          const startTime = Date.now();
          const logs = [];
          const customConsole = {
            log: (...a) => logs.push(a.map(x => (typeof x === "object" ? JSON.stringify(x) : String(x))).join(" ")),
            warn: (...a) => logs.push("[WARN] " + a.map(x => (typeof x === "object" ? JSON.stringify(x) : String(x))).join(" ")),
            error: (...a) => logs.push("[ERROR] " + a.map(x => (typeof x === "object" ? JSON.stringify(x) : String(x))).join(" "))
          };

          let targetDoc = null;
          let targetWb = null;
          let targetPres = null;
          try { targetDoc = getWordDocument(app, params?.documentName); } catch (e) {}
          try { targetWb = getWorkbook(app, params?.workbookName); } catch (e) {}
          try { targetPres = getPptPresentation(app, params?.presentationName); } catch (e) {}

          const scriptCode = params?.code || params?.script || "";
          if (!scriptCode || typeof scriptCode !== "string") {
            throw new Error("缺少要执行的脚本代码: code");
          }

          // 封装具有返回值的异步/同步图灵执行沙箱
          const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;
          const runner = new AsyncFunction(
            "app",
            "doc",
            "wb",
            "pres",
            "wps",
            "params",
            "console",
            `"use strict";\n${scriptCode}`
          );

          let evalResult;
          try {
            evalResult = await runner(
              app,
              targetDoc,
              targetWb,
              targetPres,
              typeof wps !== "undefined" ? wps : undefined,
              params?.params || {},
              customConsole
            );
          } catch (execErr) {
            throw new Error(`原生脚本执行异常: ${execErr.message}\n堆栈: ${execErr.stack || "无"}\n控制台输出: ${logs.join("\n")}`);
          }

          // 安全序列化返回值（防止 Office 原生对象循环引用）
          function safeSerialize(val, depth = 0) {
            if (val === null || val === undefined) return val;
            if (typeof val !== "object") return val;
            if (depth > 2) return String(val);
            if (Array.isArray(val)) return val.map(item => safeSerialize(item, depth + 1));
            const out = {};
            for (const k in val) {
              try {
                const v = val[k];
                if (typeof v === "function") continue;
                out[k] = safeSerialize(v, depth + 1);
              } catch (e) {}
            }
            return Object.keys(out).length > 0 ? out : String(val);
          }

          result = {
            success: true,
            executionTimeMs: Date.now() - startTime,
            returnValue: safeSerialize(evalResult),
            logs,
            message: `WPS 原生图灵脚本执行完毕（耗时 ${Date.now() - startTime}ms）`
          };
          break;
        }

        case "inspect_api": {
          const targetExpr = params?.expression || params?.path || "app";
          let targetDoc = null;
          let targetWb = null;
          let targetPres = null;
          try { targetDoc = getWordDocument(app, params?.documentName); } catch (e) {}
          try { targetWb = getWorkbook(app, params?.workbookName); } catch (e) {}
          try { targetPres = getPptPresentation(app, params?.presentationName); } catch (e) {}

          const resolver = new Function(
            "app", "doc", "wb", "pres", "wps",
            `try { return ${targetExpr}; } catch (e) { return { __inspect_error: e.message }; }`
          );

          const obj = resolver(
            app,
            targetDoc,
            targetWb,
            targetPres,
            typeof wps !== "undefined" ? wps : undefined
          );

          if (!obj || obj.__inspect_error) {
            result = {
              success: false,
              expression: targetExpr,
              error: obj ? obj.__inspect_error : "目标对象为空 (null/undefined)"
            };
            break;
          }

          const propList = [];
          const methodList = [];
          const memberMap = new Set();

          // 遍历对象属性与原型链
          let curr = obj;
          let depth = 0;
          while (curr && depth < 3) {
            try {
              const names = Object.getOwnPropertyNames(curr);
              for (const name of names) {
                if (memberMap.has(name) || name.startsWith("__")) continue;
                memberMap.add(name);
                try {
                  const val = obj[name];
                  const type = typeof val;
                  if (type === "function") {
                    methodList.push(name);
                  } else {
                    propList.push({
                      name,
                      type,
                      valueSample: type === "object" ? (val ? "[Object]" : "null") : String(val).slice(0, 80)
                    });
                  }
                } catch (e) {
                  propList.push({ name, type: "unknown", error: "无法读取" });
                }
              }
            } catch (e) {}
            try { curr = Object.getPrototypeOf(curr); } catch (e) { break; }
            depth++;
          }

          result = {
            success: true,
            expression: targetExpr,
            typeName: typeof obj,
            constructorName: obj.constructor ? obj.constructor.name : "Object",
            propertyCount: propList.length,
            methodCount: methodList.length,
            properties: propList.slice(0, 100),
            methods: methodList.sort(),
            message: `成功完成对 [${targetExpr}] 的运行时 API 反射探测`
          };
          break;
        }

        case "reload":
          window.location.reload();
          result = { reloading: true };
          break;
        default:
          throw new Error(`未知的 RPC 方法: ${method}`);
      }
      sendRpcResponse(id, result, null);
    } catch (err) {
      log(`执行方法 [${method}] 失败: ${err.message}`);
      sendRpcResponse(id, null, err.message);
    }
  }

  // ---------------------------------------------------------------------------
  // sheet-sort.js — 表格取值比较与"是否已排序"判定的纯函数（不调用任何宿主 API）
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 抽出来的原因：排序写完必须**读回校验**才能避免"返回 success 但顺序没变"（问题台账 ISS-38），
  // 而校验逻辑必须可单元测试——否则"校验恒真"这种缺陷没人发现得了。
  // 例外：文件末尾的 @build-strip 导出块只为 Node 单元测试直接 import 使用，构建拼接时整段移除。
  // ---------------------------------------------------------------------------

  /** 单元格比较：空值最小，数字按数值，其余按字符串。 */
  function compareCellValues(a, b) {
    if (a === b) return 0;
    const emptyA = a === null || a === undefined || a === "";
    const emptyB = b === null || b === undefined || b === "";
    if (emptyA && emptyB) return 0;
    if (emptyA) return -1;
    if (emptyB) return 1;
    const na = Number(a), nb = Number(b);
    if (Number.isFinite(na) && Number.isFinite(nb)) return na === nb ? 0 : (na < nb ? -1 : 1);
    const sa = String(a), sb = String(b);
    return sa === sb ? 0 : (sa < sb ? -1 : 1);
  }

  /**
   * 校验二维数据是否已按 rules 指定的列升/降序。
   *
   * - `rules[].colIndex` 是**区域内的相对列号**（与实现里 `targetRange.Columns.Item` 一致）；
   * - 第 0 行按表头处理（排序时传 `header=1`），因此相邻对从 (1,2) 开始；
   * - 相邻对覆盖到最后一对 `(n-2, n-1)`，**不得漏最后一行**；
   * - 少于 3 行（表头 + 1 行数据）时无需比较，直接视为已排序。
   */
  function isSortedByRules(matrix, rules) {
    if (!Array.isArray(matrix) || !Array.isArray(rules) || rules.length === 0) return true;
    if (matrix.length < 3) return true;
    for (let i = 2; i < matrix.length; i++) {
      for (let r = 0; r < rules.length; r++) {
        const col = Number(rules[r].colIndex) - 1;
        if (!Number.isFinite(col) || col < 0) return false;
        const desc = rules[r].order === "desc";
        const cmp = compareCellValues(matrix[i - 1] ? matrix[i - 1][col] : null, matrix[i] ? matrix[i][col] : null);
        if (cmp === 0) continue;
        if (desc ? cmp < 0 : cmp > 0) return false;
        break;
      }
    }
    return true;
  }

  // ---------------------------------------------------------------------------
  // excel.js — WPS 表格（Excel/ET）全部 RPC 实现
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
  // 2. 提取原表设计语言 (Design Token)
  function getStyleToken(app, params) {
    const { sheetName, sampleAddress = "A3", workbookName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(sampleAddress);

    const fontName = range.Font.Name || "微软雅黑";
    const fontSize = range.Font.Size || 11;
    const fontBold = !!range.Font.Bold;
    const fontColor = excelColorToHex(range.Font.Color);
    const headerBg = excelColorToHex(range.Interior.Color);

    // 尝试探测大标题
    let titleFontName = fontName;
    try {
      const titleCell = sheet.Range("A1");
      if (titleCell.Font.Name) titleFontName = titleCell.Font.Name;
    } catch (e) {}

    return {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      fontName: fontName,
      titleFontName: titleFontName,
      sampleFontSize: fontSize,
      sampleBold: fontBold,
      fontColor: fontColor,
      headerBackgroundColor: headerBg || "#1E3A8A"
    };
  }

  // 3. 新建或激活工作表 (显式锁定 workbookName)
  function createWorksheet(app, params) {
    const { sheetName, workbookName } = params;
    if (!sheetName) throw new Error("缺少 sheetName 参数");

    const wb = getWorkbook(app, workbookName);

    let targetSheet = null;
    try {
      targetSheet = wb.Worksheets.Item(sheetName);
    } catch (e) {}

    let isNew = false;
    if (!targetSheet) {
      const count = wb.Worksheets.Count;
      targetSheet = wb.Worksheets.Add(null, wb.Worksheets.Item(count));
      targetSheet.Name = sheetName;
      isNew = true;
    }

    try {
      wb.Activate();
      targetSheet.Activate();
    } catch (e) {}

    return {
      success: true,
      workbookName: wb.Name,
      sheetName: targetSheet.Name,
      isNew: isNew,
      sheetCount: wb.Worksheets.Count
    };
  }

  // 4. 删除指定工作表
  function deleteWorksheet(app, params) {
    const { sheetName, workbookName } = params;
    if (!sheetName) throw new Error("缺少 sheetName 参数");

    const wb = getWorkbook(app, workbookName);
    let targetSheet = null;
    try {
      targetSheet = wb.Worksheets.Item(sheetName);
    } catch (e) {
      return { success: true, workbookName: wb.Name, message: `工作表 ${sheetName} 已不存在，无需删除` };
    }

    if (wb.Worksheets.Count <= 1) {
      throw new Error(`工作簿 [${wb.Name}] 仅剩 1 个工作表，不能删除最后一个工作表`);
    }

    try { app.DisplayAlerts = false; } catch (e) {}
    try {
      targetSheet.Delete();
    } finally {
      try { app.DisplayAlerts = true; } catch (e) {}
    }

    return {
      success: true,
      workbookName: wb.Name,
      message: `已成功删除工作簿 [${wb.Name}] 中的工作表 [${sheetName}]`,
      sheetCount: wb.Worksheets.Count
    };
  }

  // 5. 清理指定区域
  function clearRange(app, params) {
    const { sheetName, address, workbookName } = params;
    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);
    range.Clear();
    return { success: true, workbookName: sheet.Parent.Name, clearedAddress: range.Address() };
  }

  // 6. 获取大纲
  function getSheetOutline(app, params) {
    const sheetName = typeof params === "string" ? params : params?.sheetName;
    const workbookName = typeof params === "object" ? params?.workbookName : null;
    const sheet = getWorksheet(app, sheetName, workbookName);
    const usedRange = sheet.UsedRange;

    if (!usedRange || usedRange.Rows.Count === 0 || usedRange.Columns.Count === 0) {
      return {
        sheetName: sheet.Name,
        isEmpty: true,
        usedRangeAddress: "",
        rowCount: 0,
        columnCount: 0,
        headerPreview: []
      };
    }

    const address = usedRange.Address();
    const rowCount = usedRange.Rows.Count;
    const colCount = usedRange.Columns.Count;

    const sampleRows = Math.min(rowCount, 3);
    const previewRange = sheet.Range(sheet.Cells.Item(usedRange.Row, usedRange.Column), sheet.Cells.Item(usedRange.Row + sampleRows - 1, usedRange.Column + colCount - 1));
    const rawValues = previewRange.Value2;

    let headerPreview = [];
    if (Array.isArray(rawValues)) {
      headerPreview = rawValues;
    } else if (rawValues !== undefined && rawValues !== null) {
      headerPreview = [[rawValues]];
    }

    return {
      sheetName: sheet.Name,
      isEmpty: false,
      usedRangeAddress: address,
      startRow: usedRange.Row,
      startColumn: usedRange.Column,
      rowCount: rowCount,
      columnCount: colCount,
      headerPreview: headerPreview
    };
  }

  // 7. 切片读取数据
  function readRangeData(app, params) {
    const { sheetName, address, workbookName, includeFormulas = true, includeNumberFormats = false } = params;
    if (!address) throw new Error("缺少必要参数: address (例如 'A1:C10')");

    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);

    if (range.Rows.Count * range.Columns.Count > 10000) throw new Error("单次读取上限为 10000 个单元格，请分片读取");
    const values = range.Value2;
    let formulas = null;
    let numberFormats = null;

    if (includeFormulas) {
      formulas = range.Formula;
    }
    if (includeNumberFormats) {
      numberFormats = [];
      for (let r = 1; r <= range.Rows.Count; r++) {
        const row = [];
        for (let c = 1; c <= range.Columns.Count; c++) row.push(safeRead(() => range.Cells.Item(r, c).NumberFormat, null));
        numberFormats.push(row);
      }
    }

    return {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      rowCount: range.Rows.Count,
      columnCount: range.Columns.Count,
      values: normalize2DArray(values, range.Rows.Count, range.Columns.Count),
      formulas: formulas !== null && formulas !== undefined ? normalize2DArray(formulas, range.Rows.Count, range.Columns.Count) : null,
      numberFormats: numberFormats ? normalize2DArray(numberFormats, range.Rows.Count, range.Columns.Count) : null
    };
  }

  function safeRead(reader, fallback) {
    try {
      const value = reader();
      return value === undefined ? fallback : value;
    } catch (e) {
      return fallback;
    }
  }

  function alignmentName(value, vertical) {
    const map = vertical
      ? { "-4160": "top", "-4108": "center", "-4107": "bottom" }
      : { "-4131": "left", "-4108": "center", "-4152": "right", "1": "general" };
    return map[String(value)] || value;
  }

  function readBorders(range) {
    const borderIds = { left: 7, top: 8, bottom: 9, right: 10, insideHorizontal: 11, insideVertical: 12 };
    const result = {};
    Object.keys(borderIds).forEach((name) => {
      const border = safeRead(() => range.Borders.Item(borderIds[name]), null);
      if (!border) return;
      result[name] = {
        lineStyle: safeRead(() => border.LineStyle, null),
        weight: safeRead(() => border.Weight, null),
        color: excelColorToHex(safeRead(() => border.Color, null))
      };
    });
    return result;
  }

  function readStyleFields(range, include) {
    const result = {};
    const wants = (field) => include.indexOf(field) >= 0;
    if (wants("fontName")) result.fontName = safeRead(() => range.Font.Name, null);
    if (wants("fontSize")) result.fontSize = safeRead(() => range.Font.Size, null);
    if (wants("bold")) result.bold = safeRead(() => range.Font.Bold, null);
    if (wants("fontColor")) result.fontColor = excelColorToHex(safeRead(() => range.Font.Color, null));
    if (wants("backgroundColor")) result.backgroundColor = excelColorToHex(safeRead(() => range.Interior.Color, null));
    if (wants("numberFormat")) result.numberFormat = safeRead(() => range.NumberFormat, null);
    if (wants("horizontalAlignment")) result.horizontalAlignment = alignmentName(safeRead(() => range.HorizontalAlignment, null), false);
    if (wants("verticalAlignment")) result.verticalAlignment = alignmentName(safeRead(() => range.VerticalAlignment, null), true);
    if (wants("wrapText")) result.wrapText = safeRead(() => range.WrapText, null);
    if (wants("rowHeight")) result.rowHeight = safeRead(() => range.RowHeight, null);
    if (wants("columnWidth")) result.columnWidth = safeRead(() => range.ColumnWidth, null);
    if (wants("merged")) result.merged = safeRead(() => range.MergeCells, null);
    if (wants("mergeArea")) {
      const firstCell = safeRead(() => range.Cells.Item(1, 1), null);
      const isMerged = firstCell ? safeRead(() => !!firstCell.MergeCells, false) : false;
      result.mergeArea = isMerged ? safeRead(() => firstCell.MergeArea.Address(), null) : null;
    }
    if (wants("borders")) result.borders = readBorders(range);
    return result;
  }

  function getRangeStyles(app, params) {
    const config = params || {};
    const { sheetName, workbookName, address, mode = "summary" } = config;
    if (!address) throw new Error("缺少必要参数: address (例如 'A1:C10')");
    const allowed = ["fontName", "fontSize", "bold", "fontColor", "backgroundColor", "numberFormat", "horizontalAlignment", "verticalAlignment", "wrapText", "rowHeight", "columnWidth", "merged", "mergeArea", "borders"];
    const defaults = ["fontName", "fontSize", "bold", "fontColor", "backgroundColor", "numberFormat", "horizontalAlignment", "verticalAlignment", "wrapText", "rowHeight", "columnWidth", "merged", "mergeArea"];
    const include = Array.isArray(config.include) ? config.include.filter((field) => allowed.indexOf(field) >= 0) : defaults;
    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);
    const base = {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      rowCount: range.Rows.Count,
      columnCount: range.Columns.Count,
      mode
    };

    if (mode !== "cells") {
      const styles = readStyleFields(range, include);
      const mixedOrUnavailableFields = Object.keys(styles).filter((field) => {
        if (field === "mergeArea" && styles.merged === false) return false;
        return styles[field] === null;
      });
      return Object.assign(base, { styles, mixedOrUnavailableFields });
    }

    const maxCells = Math.max(1, Math.min(500, Math.trunc(Number(config.maxCells) || 100)));
    const cells = [];
    const totalCells = range.Rows.Count * range.Columns.Count;
    for (let r = 1; r <= range.Rows.Count && cells.length < maxCells; r++) {
      for (let c = 1; c <= range.Columns.Count && cells.length < maxCells; c++) {
        const cell = range.Cells.Item(r, c);
        cells.push(Object.assign({ address: cell.Address() }, readStyleFields(cell, include)));
      }
    }
    return Object.assign(base, { totalCells, returnedCells: cells.length, truncated: totalCells > cells.length, cells });
  }

  // 8. 单元格搜索
  function searchCells(app, params) {
    const { sheetName, query, maxResults = 50, workbookName, address } = params;
    if (!query) throw new Error("缺少搜索关键字: query");

    const sheet = getWorksheet(app, sheetName, workbookName);
    // 可选 address：原实现只能全表检索，报错却让调用方"缩小检索范围"（问题台账 ISS-29）。
    const usedRange = address ? sheet.Range(address) : sheet.UsedRange;
    if (!usedRange) return { matches: [] };

    if (usedRange.Rows.Count * usedRange.Columns.Count > 100000) {
      throw new Error("检索范围超过 100000 单元格，请用 address 指定更小的区域");
    }
    const matches = [];
    const values = normalize2DArray(usedRange.Value2, usedRange.Rows.Count, usedRange.Columns.Count);
    // 说明承诺可搜"文本或公式"，但原实现只读 Value2 —— 按公式搜必然 0 命中且返回 success
    // （问题台账 ISS-29，比报错更危险）。这里同时检索公式串。
    const formulas = normalize2DArray(usedRange.Formula, usedRange.Rows.Count, usedRange.Columns.Count);
    const startRow = usedRange.Row;
    const startCol = usedRange.Column;

    const lowerQuery = String(query).toLowerCase();

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const val = values[r][c];
        const formula = formulas[r] ? formulas[r][c] : null;
        const hitValue = val !== null && val !== undefined && String(val).toLowerCase().includes(lowerQuery);
        const hitFormula = formula !== null && formula !== undefined && String(formula).toLowerCase().includes(lowerQuery);
        if (hitValue || hitFormula) {
          const cellRow = startRow + r;
          const cellCol = startCol + c;
          matches.push({
            address: sheet.Cells.Item(cellRow, cellCol).Address(),
            row: cellRow,
            column: cellCol,
            value: val,
            formula: formula === null || formula === undefined ? "" : String(formula),
            matchedIn: hitValue ? "value" : "formula"
          });
          if (matches.length >= maxResults) break;
        }
      }
      if (matches.length >= maxResults) break;
    }

    return {
      sheetName: sheet.Name,
      query: query,
      searchedRange: usedRange.Address(),
      totalFound: matches.length,
      matches: matches
    };
  }

  // 9. 批量修改数据 (带快照)
  function patchCells(app, params) {
    const { sheetName, address, values, formulas, workbookName } = params;
    if (!address) throw new Error("缺少必要参数: address (例如 'B2:C5')");
    if (!values && !formulas) throw new Error("必须提供 values 或 formulas 进行更新");

    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);
    const rowCount = range.Rows.Count;
    const colCount = range.Columns.Count;
    if (rowCount * colCount > 10000) throw new Error("单次写入上限为 10000 个单元格，请分批处理");
    [values, formulas].forEach(matrix => {
      if (matrix !== undefined && matrix !== null && (!Array.isArray(matrix) || matrix.length !== rowCount || matrix.some(row => !Array.isArray(row) || row.length !== colCount))) throw new Error("写入矩阵必须与目标区域尺寸一致");
    });

    const oldValues = normalize2DArray(range.Value2, rowCount, colCount);
    const oldFormulas = normalize2DArray(range.Formula, rowCount, colCount);

    const beforeSnapshot = {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      rowCount: rowCount,
      columnCount: colCount,
      values: oldValues,
      formulas: oldFormulas
    };

    try {
      if (values !== undefined && values !== null) range.Value2 = values;
      if (formulas !== undefined && formulas !== null) {
        // 空项（''/null/undefined）表示"该单元格只写值、不改公式"，必须**跳过**而不是写入：
        // 给 range.Formula 赋空字符串等于清空单元格，会把上一步刚写入的值一起抹掉，
        // 而且仍返回 success + modifiedCount，属于静默数据丢失（见问题台账 ISS-01 / DP7）。
        // 只有在整张矩阵都没有空项时才用整批赋值，保住批量性能。
        const hasBlank = formulas.some(row => Array.isArray(row) && row.some(cell => cell === '' || cell === null || cell === undefined));
        if (hasBlank) {
          for (let r = 0; r < rowCount; r++) {
            const row = formulas[r];
            if (!Array.isArray(row)) continue;
            for (let c = 0; c < colCount; c++) {
              const cellFormula = row[c];
              if (cellFormula === '' || cellFormula === null || cellFormula === undefined) continue;
              range.Cells.Item(r + 1, c + 1).Formula = cellFormula;
            }
          }
        } else {
          range.Formula = formulas;
        }
      }
    } catch (writeError) {
      try { range.Formula = oldFormulas; }
      catch (restoreError) { throw new Error("写入部分失败且恢复失败，请检查目标区域：" + writeError.message); }
      throw new Error("写入失败，已恢复原值与公式：" + writeError.message);
    }

    const newValues = normalize2DArray(range.Value2, rowCount, colCount);
    const newFormulas = normalize2DArray(range.Formula, rowCount, colCount);

    const diff = [];
    for (let r = 0; r < rowCount; r++) {
      for (let c = 0; c < colCount; c++) {
        const oVal = oldValues[r] ? oldValues[r][c] : null;
        const nVal = newValues[r] ? newValues[r][c] : null;
        const oForm = oldFormulas[r] ? oldFormulas[r][c] : null;
        const nForm = newFormulas[r] ? newFormulas[r][c] : null;

        if (oVal !== nVal || oForm !== nForm) {
          diff.push({
            cell: range.Cells.Item(r + 1, c + 1).Address(),
            oldValue: oVal,
            newValue: nVal,
            oldFormula: oForm,
            newFormula: nForm
          });
        }
      }
    }

    const afterSnapshot = {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      rowCount: rowCount,
      columnCount: colCount,
      values: newValues,
      formulas: newFormulas
    };

    return {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      modifiedCount: diff.length,
      diff: diff,
      beforeSnapshot: beforeSnapshot,
      afterSnapshot: afterSnapshot
    };
  }

  // 10. 高级格式与排版引擎 (美学升级)
  function formatCells(app, params) {
    const {
      sheetName,
      address,
      workbookName,
      fontName,
      fontSize,
      bold,
      fontColor,
      backgroundColor,
      horizontalAlignment,
      verticalAlignment = "center",
      numberFormat,
      rowHeight,
      wrapText,
      borders
    } = params;

    if (!address) throw new Error("缺少 address 参数");

    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);

    // 1. 字体与大小
    if (fontName) range.Font.Name = fontName;
    if (fontSize) range.Font.Size = fontSize;
    if (bold !== undefined) range.Font.Bold = !!bold;

    // 2. 颜色
    if (backgroundColor) {
      const bgr = hexToExcelColor(backgroundColor);
      if (bgr !== null) range.Interior.Color = bgr;
    }
    if (fontColor) {
      const bgr = hexToExcelColor(fontColor);
      if (bgr !== null) range.Font.Color = bgr;
    }

    // 3. 对齐
    if (horizontalAlignment) {
      if (horizontalAlignment === "center") range.HorizontalAlignment = -4108; // xlCenter
      else if (horizontalAlignment === "left") range.HorizontalAlignment = -4131; // xlLeft
      else if (horizontalAlignment === "right") range.HorizontalAlignment = -4152; // xlRight
    }
    if (verticalAlignment === "center") {
      range.VerticalAlignment = -4108; // xlCenter
    }

    // 4. 换行与行高
    if (wrapText !== undefined) {
      range.WrapText = !!wrapText;
    }
    if (rowHeight) {
      range.RowHeight = rowHeight;
    }

    // 5. 数字格式
    if (numberFormat) {
      range.NumberFormat = numberFormat;
    }

    // 6. 精细浅灰边框
    if (borders) {
      try {
        const borderBgr = hexToExcelColor(typeof borders === "string" ? borders : "#CBD5E1");
        // 边框: xlEdgeLeft=7, xlEdgeTop=8, xlEdgeBottom=9, xlEdgeRight=10, xlInsideHorizontal=11, xlInsideVertical=12
        [7, 8, 9, 10, 11, 12].forEach((bId) => {
          try {
            const b = range.Borders.Item(bId);
            b.LineStyle = 1; // xlContinuous
            b.Weight = 2; // xlThin
            if (borderBgr !== null) b.Color = borderBgr;
          } catch (e) {}
        });
      } catch (e) {}
    }

    // 7. 合并 / 拆分单元格
    // 执行后必须读回校验：原来 catch 里只 log，失败也会返回 success —— 调用方以为合并成功（问题台账 ISS-42）。
    if (params.merge === true || params.unmerge === true) {
      const wantMerged = params.merge === true;
      const actionLabel = wantMerged ? "合并" : "取消合并";
      let failure = null;
      try {
        if (wantMerged) range.Merge(); else range.UnMerge();
      } catch (e) {
        failure = e.message;
      }
      let mergedNow = null;
      try { mergedNow = range.MergeCells === true; } catch (e) { mergedNow = null; }
      if (failure) throw new Error(`${actionLabel}失败：${failure}`);
      if (mergedNow !== null && mergedNow !== wantMerged) {
        throw new Error(`${actionLabel}未生效（读回 merged=${mergedNow}）；若目标区域与已有合并区部分重叠，请先取消原有合并`);
      }
    }

    return {
      success: true,
      address: range.Address(),
      appliedStyles: params
    };
  }

  // 10.1 条件格式与数据条/色阶
  function addConditionalFormatting(app, params) {
    const {
      sheetName,
      address,
      workbookName,
      ruleType = "cell_value",
      operator = "less_than",
      formula1,
      formula2,
      backgroundColor,
      fontColor,
      barColor,
      colorScaleMin,
      colorScaleMax,
      clearExisting = false
    } = params || {};

    if (!address) throw new Error("缺少 address 参数");
    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);

    if (clearExisting) {
      try { range.FormatConditions.Delete(); } catch (e) {}
    }

    if (ruleType === "cell_value") {
      let xlOp = 6; // xlLess
      if (operator === "greater_than") xlOp = 5;
      else if (operator === "less_than") xlOp = 6;
      else if (operator === "equal") xlOp = 3;
      else if (operator === "between") xlOp = 1;

      const f1 = formula1 !== undefined ? String(formula1) : "0";
      const f2 = formula2 !== undefined ? String(formula2) : undefined;
      const fc = range.FormatConditions.Add(1, xlOp, f1, f2);

      if (backgroundColor) {
        const bg = hexToExcelColor(backgroundColor);
        if (bg !== null) fc.Interior.Color = bg;
      }
      if (fontColor) {
        const fg = hexToExcelColor(fontColor);
        if (fg !== null) fc.Font.Color = fg;
      }
    } else if (ruleType === "data_bar") {
      const db = range.FormatConditions.AddDatabar();
      if (barColor && db.BarColor) {
        const bc = hexToExcelColor(barColor);
        if (bc !== null) db.BarColor.Color = bc;
      }
    } else if (ruleType === "color_scale") {
      const cs = range.FormatConditions.AddColorScale(2);
      if (colorScaleMin && cs.ColorScaleCriteria) {
        const c1 = hexToExcelColor(colorScaleMin);
        if (c1 !== null) cs.ColorScaleCriteria.Item(1).FormatColor.Color = c1;
      }
      if (colorScaleMax && cs.ColorScaleCriteria) {
        const c2 = hexToExcelColor(colorScaleMax);
        if (c2 !== null) cs.ColorScaleCriteria.Item(2).FormatColor.Color = c2;
      }
    } else {
      throw new Error(`未知的条件格式类型: ${ruleType} (支持 cell_value, data_bar, color_scale)`);
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      ruleType: ruleType,
      message: `已成功在 ${range.Address()} 应用 ${ruleType} 条件格式`
    };
  }

  // 10.2 冻结窗格吸顶
  function freezePanes(app, params) {
    const { sheetName, workbookName, freezeRowIndex, freezeColumnIndex, unfreeze } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}

    const win = app.ActiveWindow;
    if (!win) throw new Error("无法获取 WPS 活动窗口句柄");

    if (unfreeze) {
      win.FreezePanes = false;
      win.SplitRow = 0;
      win.SplitColumn = 0;
      return { success: true, message: "已成功解除当前工作表的窗格冻结" };
    }

    win.FreezePanes = false;
    if (freezeRowIndex && freezeRowIndex > 1) {
      win.SplitRow = freezeRowIndex - 1;
    }
    if (freezeColumnIndex && freezeColumnIndex > 1) {
      win.SplitColumn = freezeColumnIndex - 1;
    }
    win.FreezePanes = true;

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      freezeRowIndex: freezeRowIndex || 1,
      freezeColumnIndex: freezeColumnIndex || 1,
      message: `已成功锁定第 ${freezeRowIndex ? freezeRowIndex - 1 : 0} 行之上的表头吸顶显示`
    };
  }

  // 10.3 整行整列增删、隐藏与高度宽度控制
  function modifyRowsColumns(app, params) {
    const { sheetName, workbookName, targetType, action, index, count = 1, size } = params || {};
    if (!targetType || !action || index === undefined) {
      throw new Error("缺少必要参数: targetType ('row'|'column'), action ('insert'|'delete'|'hide'|'unhide'|'set_size'), index (起始行号或列号/列字母)");
    }
    const sheet = getWorksheet(app, sheetName, workbookName);

    // 解析列号（支持数字如 2 或字母如 "B"）
    function parseColIndex(val) {
      if (typeof val === "number") return val;
      const str = String(val).toUpperCase().trim();
      if (/^\d+$/.test(str)) return parseInt(str, 10);
      let num = 0;
      for (let i = 0; i < str.length; i++) {
        num = num * 26 + (str.charCodeAt(i) - 64);
      }
      return num > 0 ? num : 1;
    }

    if (targetType === "row") {
      const startRow = Number(index);
      const endRow = startRow + Number(count) - 1;
      const rowRange = sheet.Range(`${startRow}:${endRow}`);

      if (action === "insert") {
        rowRange.Insert(-4121); // xlDown
      } else if (action === "delete") {
        rowRange.Delete(-4162); // xlUp
      } else if (action === "hide" || action === "unhide") {
        const want = action === "hide";
        // 本机 WPS 上给 `Range("5:6").Hidden` 赋值是**静默 no-op**（列走 Columns.Item 就正常），
        // 因此改为逐行写 Rows.Item(n).Hidden，并在写完后**读回校验**，未生效即报错（问题台账 ISS-60）。
        for (let r = startRow; r <= endRow; r++) {
          sheet.Rows.Item(r).Hidden = want;
        }
        const readBack = [];
        for (let r = startRow; r <= endRow; r++) {
          readBack.push(sheet.Rows.Item(r).Hidden === true);
        }
        if (readBack.some(applied => applied !== want)) {
          throw new Error(
            `${want ? "隐藏" : "取消隐藏"}第 ${startRow}-${endRow} 行未生效（逐行读回 ${JSON.stringify(readBack)}）。` +
            `已改用 Rows.Item(n).Hidden 逐行写入，若仍失败请检查该行是否被工作表保护。`
          );
        }
      } else if (action === "set_size") {
        if (size === undefined) throw new Error("设置行高时必须传入 size 参数 (单位: 磅值)");
        rowRange.RowHeight = Number(size);
      } else {
        throw new Error(`未知的行操作动作: ${action}`);
      }
    } else if (targetType === "column") {
      const startCol = parseColIndex(index);
      for (let i = 0; i < count; i++) {
        const col = sheet.Columns.Item(startCol + i);
        if (action === "insert") {
          col.Insert(-4161); // xlToRight
        } else if (action === "delete") {
          col.Delete(-4159); // xlToLeft
        } else if (action === "hide") {
          col.Hidden = true;
        } else if (action === "unhide") {
          col.Hidden = false;
        } else if (action === "set_size") {
          if (size === undefined) throw new Error("设置列宽时必须传入 size 参数 (字符宽度)");
          col.ColumnWidth = Number(size);
        } else {
          throw new Error(`未知的列操作动作: ${action}`);
        }
      }
    } else {
      throw new Error(`targetType 必须是 'row' 或 'column'`);
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      targetType,
      action,
      index,
      count,
      size,
      message: `已成功对工作表 [${sheet.Name}] 的第 ${index} ${targetType === "row" ? "行" : "列"}（共 ${count} 项）执行 ${action} 操作`
    };
  }

  // 10.4 单元格批注与审阅备注管理
  function manageCellComments(app, params) {
    const { sheetName, workbookName, address, action = "read", text, author } = params || {};
    if (!address && action !== "clear_all") {
      throw new Error("缺少必要参数: address (如 'C5' 或 'D10:D20')");
    }
    const sheet = getWorksheet(app, sheetName, workbookName);
    const targetRange = address ? sheet.Range(address) : sheet.UsedRange;

    if (action === "add") {
      if (!text) throw new Error("添加批注时必须提供 text 参数");
      const cell = targetRange.Cells.Item(1, 1);
      try {
        if (cell.Comment) cell.Comment.Delete();
      } catch (e) {}
      const commentContent = author ? `${author}:\n${text}` : text;
      const comment = cell.AddComment(commentContent);
      comment.Visible = false;
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        address: cell.Address,
        action: "add",
        commentText: text,
        author: author || "AI 智能审核",
        message: `已在单元格 ${cell.Address} 成功添加批注`
      };
    } else if (action === "read") {
      const comments = [];
      const rowCount = targetRange.Rows.Count;
      const colCount = targetRange.Columns.Count;
      for (let r = 1; r <= rowCount; r++) {
        for (let c = 1; c <= colCount; c++) {
          const cell = targetRange.Cells.Item(r, c);
          try {
            if (cell.Comment) {
              comments.push({
                address: cell.Address,
                text: cell.Comment.Text(),
                author: cell.Comment.Author || undefined
              });
            }
          } catch (e) {}
        }
      }
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        address: targetRange.Address,
        comments,
        totalComments: comments.length
      };
    } else if (action === "delete") {
      const cell = targetRange.Cells.Item(1, 1);
      try {
        if (cell.Comment) cell.Comment.Delete();
      } catch (e) {}
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        address: cell.Address,
        action: "delete",
        message: `已成功删除单元格 ${cell.Address} 上的批注`
      };
    } else if (action === "clear_all") {
      try {
        sheet.Cells.ClearComments();
      } catch (e) {}
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        action: "clear_all",
        message: `已清空工作表 [${sheet.Name}] 中的所有批注`
      };
    }
    throw new Error(`未知的批注操作 action: ${action} (支持 'add' | 'read' | 'delete' | 'clear_all')`);
  }

  // 10.5 全局查找与定位替换
  function findAndReplace(app, params) {
    const { sheetName, workbookName, searchQuery, replaceText, matchCase = false, matchEntireCell = false, searchRange, maxResults = 50 } = params || {};
    if (searchQuery === undefined || searchQuery === null) {
      throw new Error("缺少必要参数: searchQuery (要查找的文本或数值)");
    }
    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = searchRange ? sheet.Range(searchRange) : sheet.UsedRange;
    const queryStr = String(searchQuery);
    const queryLower = queryStr.toLowerCase();

    const rowCount = range.Rows.Count;
    const colCount = range.Columns.Count;
    const values = range.Value2 || [];
    const formulas = range.Formula || [];

    const matches = [];
    let replacedCount = 0;

    for (let r = 0; r < rowCount; r++) {
      const rowVal = Array.isArray(values) ? values[r] : [values];
      const rowForm = Array.isArray(formulas) ? formulas[r] : [formulas];
      for (let c = 0; c < colCount; c++) {
        const val = rowVal ? rowVal[c] : null;
        const form = rowForm ? rowForm[c] : null;
        if (val === null || val === undefined) continue;

        const cellStr = String(val);
        const formStr = form === null || form === undefined ? "" : String(form);
        const testMatch = (text) => matchEntireCell
          ? (matchCase ? text === queryStr : text.toLowerCase() === queryLower)
          : (matchCase ? text.includes(queryStr) : text.toLowerCase().includes(queryLower));

        const hitValue = testMatch(cellStr);
        // 原实现读了 formulas 却只拿 Value2 匹配 —— 按公式片段搜必然 0 命中且返回 success
        // （问题台账 ISS-62）。这里把真正的公式串（以 = 开头）也纳入检索。
        const hitFormula = !hitValue && formStr.startsWith("=") && testMatch(formStr);
        const matched = hitValue || hitFormula;

        if (matched) {
          const cell = range.Cells.Item(r + 1, c + 1);
          const addr = cell.Address;
          let isReplaced = false;

          if (replaceText !== undefined) {
            const source = hitFormula ? formStr : cellStr;
            const newVal = matchEntireCell
              ? replaceText
              : source.replace(new RegExp(queryStr.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'g' : 'gi'), replaceText);
            // 命中在公式里就写回公式位，否则写值位
            if (hitFormula) cell.Formula = newVal; else cell.Value2 = newVal;
            isReplaced = true;
            replacedCount++;
          }

          matches.push({
            address: addr,
            row: r + 1,
            col: c + 1,
            value: val,
            formula: formStr.startsWith("=") ? formStr : null,
            matchedIn: hitFormula ? "formula" : "value",
            replaced: isReplaced
          });

          if (matches.length >= maxResults) break;
        }
      }
      if (matches.length >= maxResults) break;
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      searchQuery,
      replaceText,
      totalFound: matches.length,
      replacedCount,
      results: matches
    };
  }

  // 10.6 工作表复制与模板克隆
  function duplicateSheet(app, params) {
    const { sourceSheetName, newSheetName, position = "after", workbookName } = params || {};
    if (!sourceSheetName || !newSheetName) {
      throw new Error("缺少必要参数: sourceSheetName (源工作表) 和 newSheetName (新工作表名称)");
    }
    const wb = getWorkbook(app, workbookName);
    const srcSheet = getWorksheet(app, sourceSheetName, workbookName);

    try {
      const exist = wb.Worksheets.Item(newSheetName);
      if (exist) throw new Error(`工作表 [${newSheetName}] 已存在，请指定不同的新工作表名称`);
    } catch (e) {
      if (e.message && e.message.includes("已存在")) throw e;
    }

    if (position === "before") {
      srcSheet.Copy(srcSheet);
    } else {
      srcSheet.Copy(null, srcSheet);
    }

    const newSheet = wb.ActiveSheet;
    newSheet.Name = newSheetName;

    return {
      success: true,
      workbookName: wb.Name,
      sourceSheetName: srcSheet.Name,
      newSheetName: newSheet.Name,
      totalSheets: wb.Worksheets.Count,
      message: `已成功将工作表 [${sourceSheetName}] 完整克隆为 [${newSheetName}]（包含所有格式、公式与图表）`
    };
  }

  function saveWorkbook(app, params) {
    const { workbookName } = params || {};
    const wb = getWorkbook(app, workbookName);
    wb.Save();
    return {
      success: true,
      workbookName: wb.Name,
      fullName: wb.FullName,
      message: `工作簿 [${wb.Name}] 已成功保存到磁盘`
    };
  }

  // 11. 智能列宽排版 (AutoFit + 宽度上下限与自动换行)
  function autoFitColumns(app, params) {
    const { sheetName, columnRules, workbookName } = params;
    const sheet = getWorksheet(app, sheetName, workbookName);

    const results = [];

    if (!Array.isArray(columnRules) || columnRules.length === 0) {
      // 说明承诺"不传 columnRules 则自适应全表已用区域"，但原实现直接返回 success + results: []，
      // 列宽纹丝不动（静默 no-op，问题台账 ISS-39）。这里按已用区域的列范围逐列自适应。
      const used = sheet.UsedRange;
      const firstColumn = used && typeof used.Column === "number" ? used.Column : 1;
      const columnCount = used && used.Columns && used.Columns.Count ? used.Columns.Count : 0;
      for (let index = firstColumn; index < firstColumn + columnCount; index++) {
        const col = sheet.Columns.Item(index);
        col.AutoFit();
        results.push({ colIndex: index, finalWidth: col.ColumnWidth, wrapped: !!col.WrapText });
      }
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        mode: "usedRange",
        results,
        message: columnCount > 0
          ? `已按已用区域自适应 ${results.length} 列`
          : "已用区域为空，未调整任何列"
      };
    }

    for (const rule of columnRules) {
      const { colIndex, minWidth = 12, maxWidth = 30, wrapText = true } = rule;
      const col = sheet.Columns.Item(colIndex);

      // 先执行自适应计算
      col.AutoFit();
      let currentWidth = col.ColumnWidth;

      // 应用下限
      if (currentWidth < minWidth) {
        currentWidth = minWidth;
      }
      // 应用上限
      if (maxWidth && currentWidth > maxWidth) {
        currentWidth = maxWidth;
        if (wrapText) {
          col.WrapText = true;
        }
      }
      col.ColumnWidth = currentWidth;

      results.push({
        colIndex: colIndex,
        finalWidth: currentWidth,
        wrapped: col.WrapText
      });
    }

    return { success: true, workbookName: sheet.Parent.Name, results };
  }

  // 12. 原子回滚
  function rollbackCells(app, params) {
    const { sheetName, address, snapshot, workbookName } = params;
    if (!address || !snapshot) throw new Error("缺少必要回滚参数");

    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);

    if (snapshot.formulas) {
      range.Formula = snapshot.formulas;
    } else if (snapshot.values) {
      range.Value2 = snapshot.values;
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      address: range.Address(),
      restoredRows: range.Rows.Count,
      restoredCols: range.Columns.Count
    };
  }

  // 13. 插入行或列
  function insertDimension(app, params) {
    const { sheetName, type, index = 1, count = 1, workbookName } = params;
    const sheet = getWorksheet(app, sheetName, workbookName);

    if (type === "column") {
      const colRange = sheet.Columns.Item(index);
      for (let i = 0; i < count; i++) colRange.Insert();
      return { success: true, message: `在列 ${index} 插入 ${count} 列` };
    } else if (type === "row") {
      const rowRange = sheet.Rows.Item(index);
      for (let i = 0; i < count; i++) rowRange.Insert();
      return { success: true, message: `在行 ${index} 插入 ${count} 行` };
    }
    throw new Error("type 必须是 'row' 或 'column'");
  }

  function normalize2DArray(val, expectedRows, expectedCols) {
    if (val === undefined || val === null) {
      return Array(expectedRows).fill(null).map(() => Array(expectedCols).fill(null));
    }
    if (!Array.isArray(val)) return [[val]];
    if (!Array.isArray(val[0])) return [val];
    return val;
  }

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

  // 14. 捕获工作表或指定区域的渲染预览图 (原生 JSA 高保真推入剪贴板)
  function captureSheetPreview(app, params) {
    const { sheetName, address, range, workbookName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);

    try {
      sheet.Activate();
    } catch (e) {}

    const targetAddr = address || range;
    let targetRange;
    if (targetAddr) {
      targetRange = sheet.Range(targetAddr);
    } else {
      targetRange = sheet.UsedRange;
    }

    // 执行高保真选区图形渲染并推入系统剪贴板 (1 = xlScreen, -4147 = xlBitmap)
    targetRange.CopyPicture(1, -4147);

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: targetRange.Address()
    };
  }

  // 15. 原生矢量图表生成与管理 (第二梯队 - 健壮事务版)
  function addChart(app, params) {
    const {
      sheetName,
      workbookName,
      chartType = "column_clustered",
      dataRange,
      dataRanges,
      title,
      position,
      hasLegend = true,
      hasDataLabels = false,
      smoothLine = false,
      seriesColors,
      yAxis,
      seriesSettings,
      replaceExisting = true
    } = params || {};

    // 1. 强制公式全局重算（彻底根治跨表引用公式未计算导致分类轴塌陷的严重时钟竞争 Bug！）
    try {
      if (app && app.Calculate) {
        app.Calculate();
      }
    } catch (calcErr) {
      log("公式重算提示: " + calcErr.message);
    }

    // 2. 解析数据源（支持非连续区域联合数据源，如 ['A4:A19', 'E4:E19']）
    let targetDataRange = dataRange;
    if (!targetDataRange && Array.isArray(dataRanges) && dataRanges.length > 0) {
      targetDataRange = dataRanges.join(",");
    }
    if (!targetDataRange) {
      throw new Error("缺少必要参数: dataRange (例如 'A4:E19') 或 dataRanges (例如 ['A4:A19', 'E4:E19'])");
    }

    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}

    // 计算图表位置与尺寸
    let left = 360;
    let top = 40;
    let width = 480;
    let height = 280;

    if (position) {
      if (position.leftCell) {
        try {
          const anchorRange = sheet.Range(position.leftCell);
          left = anchorRange.Left;
          top = anchorRange.Top;
        } catch (e) {
          log("图表定位 leftCell 警告: " + e.message);
        }
      }
      if (position.width) width = Number(position.width);
      if (position.height) height = Number(position.height);
    }

    // 3. 覆盖模式（如果开启 replaceExisting，先清理该锚点处的重叠旧图，杜绝废图堆叠）
    if (replaceExisting && position && position.leftCell) {
      try {
        const shapes = sheet.Shapes;
        const total = shapes.Count;
        for (let i = total; i >= 1; i--) {
          const shp = shapes.Item(i);
          if (shp.HasChart && Math.abs(shp.Left - left) < 25 && Math.abs(shp.Top - top) < 25) {
            log(`自动清理锚点 [${position.leftCell}] 处的旧图表: ${shp.Name}`);
            shp.Delete();
          }
        }
      } catch (cleanErr) {
        log("清理旧图表提示: " + cleanErr.message);
      }
    }

    // 映射 Excel 图表类型常量
    // xlLine=4, xlColumnClustered=51, xlBarClustered=57, xlPie=5, xlDoughnut=-4120, xlPareto=122,
    // xlArea=1, xlXYScatter=-4169
    const chartTypeMap = {
      line: 4,
      column: 51,
      column_clustered: 51,
      bar: 57,
      bar_clustered: 57,
      pie: 5,
      doughnut: -4120,
      pareto: 122,
      area: 1,
      scatter: -4169
    };
    // 原来未知类型会**静默退回柱状图**并返回 success（问题台账 ISS-17）：调用方拿到的图不是要的类型却毫无察觉。
    // 现在改为显式报错，并带上可用类型清单。
    if (chartTypeMap[chartType] === undefined) {
      throw new Error(
        `不支持的图表类型: ${chartType}（可用: ${Object.keys(chartTypeMap).join(" / ")}）`
      );
    }
    const xlChartType = chartTypeMap[chartType];

    let shape = null;
    try {
      if (sheet.Shapes.AddChart2) {
        shape = sheet.Shapes.AddChart2(201, xlChartType, left, top, width, height);
      } else {
        shape = sheet.Shapes.AddChart(xlChartType, left, top, width, height);
      }
    } catch (err) {
      // 降级为默认簇状柱状图
      shape = sheet.Shapes.AddChart(51, left, top, width, height);
    }

    if (!shape || !shape.Chart) {
      throw new Error("无法在当前工作表中创建原生图表对象");
    }

    // 4. 事务性配置（一旦发生异常，自动自毁半成品 Shape，绝不在工作表留下空白废图！）
    try {
      const chart = shape.Chart;
      const srcRange = sheet.Range(targetDataRange);
      chart.SetSourceData(srcRange);

      // 安全设置标题（防范 ChartTitle 空指针崩溃）
      if (title) {
        try {
          chart.HasTitle = true;
          if (chart.ChartTitle) {
            chart.ChartTitle.Text = title;
          }
        } catch (tErr) {
          try {
            if (chart.ChartTitle && chart.ChartTitle.Characters) {
              chart.ChartTitle.Characters.Text = title;
            }
          } catch (tErr2) {
            log("设置图表标题警告: " + tErr2.message);
          }
        }
      }

      if (hasLegend !== undefined) {
        chart.HasLegend = !!hasLegend;
      }

      if (hasDataLabels) {
        try {
          chart.ApplyDataLabels();
        } catch (e) {}
      }

      // 4.1 平滑线设置与系列颜色注入
      try {
        const seriesCol = chart.SeriesCollection();
        const seriesCount = seriesCol ? seriesCol.Count : 0;

        // 全局平滑线设置 (主要针对折线图)
        if (smoothLine) {
          for (let sIdx = 1; sIdx <= seriesCount; sIdx++) {
            try {
              seriesCol.Item(sIdx).Smooth = true;
            } catch (e) {}
          }
        }

        // 系列颜色注入 (seriesColors 数组)
        if (Array.isArray(seriesColors)) {
          seriesColors.forEach((colorHex, idx) => {
            const bgr = hexToExcelColor(colorHex);
            const sIdx = idx + 1;
            if (bgr !== null && sIdx <= seriesCount) {
              try {
                const series = seriesCol.Item(sIdx);
                if (series.Format && series.Format.Line) {
                  series.Format.Line.ForeColor.RGB = bgr;
                } else if (series.Border) {
                  series.Border.Color = bgr;
                }
                if (series.Format && series.Format.Fill) {
                  series.Format.Fill.ForeColor.RGB = bgr;
                } else if (series.Interior) {
                  series.Interior.Color = bgr;
                }
              } catch (e) {
                log(`系列 ${sIdx} 颜色注入警告: ${e.message}`);
              }
            }
          });
        }

        // 高阶单系列独立设置 (seriesSettings)
        if (Array.isArray(seriesSettings)) {
          seriesSettings.forEach((ss) => {
            const sIdx = Number(ss.seriesIndex);
            if (sIdx >= 1 && sIdx <= seriesCount) {
              try {
                const series = seriesCol.Item(sIdx);
                if (ss.smooth !== undefined) {
                  series.Smooth = !!ss.smooth;
                }
                if (ss.color) {
                  const bgr = hexToExcelColor(ss.color);
                  if (bgr !== null) {
                    if (series.Format && series.Format.Line) series.Format.Line.ForeColor.RGB = bgr;
                    else if (series.Border) series.Border.Color = bgr;
                    if (series.Format && series.Format.Fill) series.Format.Fill.ForeColor.RGB = bgr;
                    else if (series.Interior) series.Interior.Color = bgr;
                  }
                }
              } catch (e) {}
            }
          });
        }
      } catch (e) {
        log("系列格式化提示: " + e.message);
      }

      // 4.2 Y 轴数值范围、刻度与数字格式设置 (核心排版防呆)
      if (yAxis) {
        try {
          // xlValue = 2 (数值坐标轴), xlPrimary = 1 (主坐标轴)
          const valAxis = chart.Axes(2, 1);
          if (valAxis) {
            if (yAxis.min !== undefined && yAxis.min !== null) {
              valAxis.MinimumScale = Number(yAxis.min);
            }
            if (yAxis.max !== undefined && yAxis.max !== null) {
              valAxis.MaximumScale = Number(yAxis.max);
            }
            if (yAxis.step !== undefined && yAxis.step !== null) {
              valAxis.MajorUnit = Number(yAxis.step);
            }
            if (yAxis.numberFormat) {
              try {
                valAxis.TickLabels.NumberFormat = yAxis.numberFormat;
              } catch (e) {}
            }
            if (yAxis.title) {
              try {
                valAxis.HasTitle = true;
                valAxis.AxisTitle.Text = yAxis.title;
              } catch (e) {}
            }
          }
        } catch (e) {
          log("数值坐标轴设置提示: " + e.message);
        }
      }

      // 建图后读回真实 ChartType：宿主可能（在任何版本）把请求的类型落成别的类型，
      // 原实现对此毫无察觉（ISS-17）。这里把实际类型带回，不一致就给出 warning。
      let actualChartType = null;
      try { actualChartType = Number(shape.Chart.ChartType); } catch (e) {}
      const typeWarnings = [];
      if (actualChartType !== null && Number.isFinite(actualChartType) && actualChartType !== xlChartType) {
        typeWarnings.push(
          `请求的 chartType=${chartType}（xlChartType=${xlChartType}）实际落成 ChartType=${actualChartType}；` +
          `请用 wps_get_charts 复核，必要时改用 wps_execute_script 直接指定常量。`
        );
      }

      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        chartType,
        requestedChartType: xlChartType,
        actualChartType,
        warnings: typeWarnings,
        dataRange: targetDataRange,
        title: title || "",
        left,
        top,
        width,
        height,
        shapeName: shape.Name,
        chartIndex: getChartOrdinal(sheet, shape.Name),
        message: `已成功在 [${sheet.Name}] 创建 ${chartType} 原生图表，数据源为 ${targetDataRange}`
      };
    } catch (transactionErr) {
      // 事务回滚：立即自毁残缺半成品 Shape，绝不在工作表留下空白废图！
      try {
        if (shape) shape.Delete();
      } catch (delErr) {}
      throw new Error(`图表创建异常并已自动销毁半成品对象: ${transactionErr.message}`);
    }
  }

  function getChartOrdinal(sheet, shapeName) {
    let ordinal = 0;
    for (let i = 1; i <= sheet.Shapes.Count; i++) {
      const shape = sheet.Shapes.Item(i);
      if (!shape.HasChart) continue;
      ordinal++;
      if (shape.Name === shapeName) return ordinal;
    }
    return null;
  }

  function getCharts(app, params) {
    const { sheetName, workbookName, shapeName, chartIndex, chartTitle, detail = false } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const charts = [];
    let ordinal = 0;

    for (let i = 1; i <= sheet.Shapes.Count; i++) {
      const shape = sheet.Shapes.Item(i);
      if (!shape.HasChart) continue;
      ordinal++;
      const chart = shape.Chart;
      const title = safeRead(() => chart.HasTitle && chart.ChartTitle ? chart.ChartTitle.Text : "", "");
      if (shapeName && shape.Name !== shapeName) continue;
      if (chartIndex && ordinal !== Number(chartIndex)) continue;
      if (chartTitle && String(title).indexOf(chartTitle) < 0) continue;

      const item = {
        chartIndex: ordinal,
        shapeName: shape.Name,
        title,
        chartType: safeRead(() => chart.ChartType, null),
        hasLegend: !!safeRead(() => chart.HasLegend, false),
        left: safeRead(() => shape.Left, null),
        top: safeRead(() => shape.Top, null),
        width: safeRead(() => shape.Width, null),
        height: safeRead(() => shape.Height, null),
        leftCell: safeRead(() => shape.TopLeftCell.Address(), null),
        seriesCount: safeRead(() => chart.SeriesCollection().Count, 0)
      };

      if (detail) {
        const series = [];
        const seriesCollection = safeRead(() => chart.SeriesCollection(), null);
        const seriesCount = seriesCollection ? safeRead(() => seriesCollection.Count, 0) : 0;
        for (let s = 1; s <= seriesCount; s++) {
          const current = seriesCollection.Item(s);
          series.push({
            seriesIndex: s,
            name: safeRead(() => current.Name, null),
            formula: safeRead(() => current.Formula, null),
            smooth: safeRead(() => current.Smooth, null),
            lineColor: excelColorToHex(safeRead(() => current.Format.Line.ForeColor.RGB, null)),
            fillColor: excelColorToHex(safeRead(() => current.Format.Fill.ForeColor.RGB, null))
          });
        }
        item.series = series;
        item.yAxis = safeRead(() => {
          const axis = chart.Axes(2, 1);
          return {
            minimumScale: axis.MinimumScale,
            maximumScale: axis.MaximumScale,
            majorUnit: axis.MajorUnit,
            numberFormat: axis.TickLabels.NumberFormat,
            title: axis.HasTitle && axis.AxisTitle ? axis.AxisTitle.Text : ""
          };
        }, null);
      }
      charts.push(item);
    }

    return {
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      totalChartCount: ordinal,
      chartCount: charts.length,
      detail: !!detail,
      charts
    };
  }

  // 15.1 原生矢量图表删除与清理引擎
  function deleteChart(app, params) {
    const {
      sheetName,
      workbookName,
      chartTitle,
      shapeName,
      leftCell,
      chartIndex,
      clearAll = false
    } = params || {};

    const sheet = getWorksheet(app, sheetName, workbookName);
    const shapes = sheet.Shapes;
    const count = shapes.Count;
    const chartOrdinals = {};
    let chartOrdinal = 0;
    for (let i = 1; i <= count; i++) {
      if (shapes.Item(i).HasChart) {
        chartOrdinal++;
        chartOrdinals[i] = chartOrdinal;
      }
    }

    let targetLeft = null;
    let targetTop = null;
    if (leftCell) {
      try {
        const anchor = sheet.Range(leftCell);
        targetLeft = anchor.Left;
        targetTop = anchor.Top;
      } catch (e) {}
    }

    // 先收集候选、再决定删不删。
    // `leftCell` 是按**像素邻近（±30px）**匹配的，多张图叠在同一位置时会**全部命中**——
    // 实测传 leftCell:"H2" 一次删掉了 14 张图（问题台账 ISS-19）。破坏性操作必须先设卡。
    const candidates = [];
    for (let i = 1; i <= count; i++) {
      const shp = shapes.Item(i);
      if (!shp.HasChart) continue;
      let reason = null;
      if (clearAll) {
        reason = "clearAll";
      } else if (shapeName && shp.Name === shapeName) {
        reason = "shapeName";
      } else if (chartIndex && chartOrdinals[i] === Number(chartIndex)) {
        reason = "chartIndex";
      } else if (targetLeft !== null && targetTop !== null) {
        if (Math.abs(shp.Left - targetLeft) < 30 && Math.abs(shp.Top - targetTop) < 30) {
          reason = "leftCell(±30px)";
        }
      } else if (chartTitle) {
        try {
          if (shp.Chart.HasTitle && shp.Chart.ChartTitle.Text.includes(chartTitle)) {
            reason = "chartTitle";
          }
        } catch (e) {}
      }
      if (reason) {
        candidates.push({ index: i, name: shp.Name, reason: reason, left: Math.round(shp.Left), top: Math.round(shp.Top) });
      }
    }

    if (candidates.length === 0) {
      throw new Error(
        `未匹配到任何图表：工作表 [${sheet.Name}] 上没有符合条件的图表。` +
        `请先用 wps_get_charts 读回图表清单（shapeName / 序号 / 标题 / 位置），再指定要删除的那一张。`
      );
    }

    if (!clearAll && candidates[0].reason === "leftCell(±30px)" && candidates.length > 1) {
      throw new Error(
        `leftCell 按像素邻近（±30px）匹配，本次命中 ${candidates.length} 张图表，**已拒绝批量删除**以免误伤。` +
        `命中清单：${candidates.map(c => `${c.name}(左${c.left},上${c.top})`).join("；")}。` +
        `请改用 shapeName 或 chartIndex 精确指定；确认要全删请显式传 clearAll: true。`
      );
    }

    const deletedNames = [];
    // 从后往前删，避免索引位移
    for (let k = candidates.length - 1; k >= 0; k--) {
      const shp = shapes.Item(candidates[k].index);
      deletedNames.push(shp.Name);
      shp.Delete();
    }
    const deletedCount = deletedNames.length;

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      deletedCount,
      deletedNames,
      matchedBy: candidates[0].reason,
      message: `已成功在工作表 [${sheet.Name}] 中删除 ${deletedCount} 张图表（匹配方式：${candidates[0].reason}）`
    };
  }

  // 16. 数据透视表一键生成 (第三梯队)
  function createPivotTable(app, params) {
    const {
      workbookName,
      sourceSheetName,
      sourceRange,
      destSheetName,
      destCell,
      rowFields = [],
      columnFields = [],
      dataFields = []
    } = params || {};

    if (!sourceRange) throw new Error("缺少必要参数: sourceRange (例如 '明细!A1:K5422')");
    if (!destCell) throw new Error("缺少必要参数: destCell (例如 'B4')");

    const wb = getWorkbook(app, workbookName);
    const srcSheet = getWorksheet(app, sourceSheetName, workbookName);
    const srcRange = srcSheet.Range(sourceRange);

    const destSheet = getWorksheet(app, destSheetName, workbookName);
    try { destSheet.Activate(); } catch (e) {}
    const destRange = destSheet.Range(destCell);

    // 1 = xlDatabase
    const pivotCache = wb.PivotCaches().Create(1, srcRange);
    const ptName = "PivotTable_" + Date.now();
    const pivotTable = pivotCache.CreatePivotTable(destRange, ptName);

    // 行字段: xlRowField = 1
    if (Array.isArray(rowFields)) {
      rowFields.forEach((rf) => {
        try {
          const pf = pivotTable.PivotFields(rf);
          if (pf) pf.Orientation = 1;
        } catch (e) {
          log("设置透视表行维度异常: " + rf + ", " + e.message);
        }
      });
    }

    // 列字段: xlColumnField = 2
    if (Array.isArray(columnFields)) {
      columnFields.forEach((cf) => {
        try {
          const pf = pivotTable.PivotFields(cf);
          if (pf) pf.Orientation = 2;
        } catch (e) {
          log("设置透视表列维度异常: " + cf + ", " + e.message);
        }
      });
    }

    // 汇总字段与函数: xlSum = 4, xlCount = 2, xlAverage = 5, xlMax = 6, xlMin = 7
    const summaryFuncMap = {
      sum: 4,
      count: 2,
      average: 5,
      max: 6,
      min: 7
    };

    if (Array.isArray(dataFields)) {
      dataFields.forEach((df) => {
        try {
          const pf = pivotTable.PivotFields(df.fieldName);
          const caption = df.caption || (`${df.summaryFunction || "求和"}:${df.fieldName}`);
          const func = summaryFuncMap[df.summaryFunction] || 4;
          pivotTable.AddDataField(pf, caption, func);
        } catch (e) {
          log("设置透视表数据字段异常: " + JSON.stringify(df) + ", " + e.message);
        }
      });
    }

    return {
      success: true,
      workbookName: wb.Name,
      pivotTableName: ptName,
      destSheetName: destSheet.Name,
      destCell,
      rowCount: rowFields.length,
      colCount: columnFields.length,
      dataCount: dataFields.length,
      message: `已成功在 [${destSheet.Name}] ${destCell} 生成数据透视表`
    };
  }

  // 17. 自动筛选与数据排序 (第三梯队)
  function setFilterAndSort(app, params) {
    const { sheetName, workbookName, range, enableAutoFilter, sortRules } = params || {};
    if (!range) throw new Error("缺少必要参数: range (例如 'A4:E20')");

    const sheet = getWorksheet(app, sheetName, workbookName);
    const targetRange = sheet.Range(range);
    const warnings = [];

    // 自动筛选控制
    let appliedFilterRange = null;
    if (enableAutoFilter !== undefined) {
      if (enableAutoFilter) {
        if (!sheet.AutoFilterMode) {
          targetRange.AutoFilter();
        }
      } else {
        if (sheet.AutoFilterMode) {
          sheet.AutoFilterMode = false;
        }
      }
      // 回读筛选实际覆盖范围：宿主会把筛选自动扩展到相邻的整块数据区，
      // 传入的 range 只是锚点而不是约束（问题台账 ISS-40）。这里把真实范围报出来，不再让调用方以为是自己传的那个。
      try {
        appliedFilterRange = sheet.AutoFilterMode && sheet.AutoFilter && sheet.AutoFilter.Range
          ? sheet.AutoFilter.Range.Address()
          : null;
      } catch (e) {
        appliedFilterRange = null;
      }
      if (appliedFilterRange && appliedFilterRange !== targetRange.Address()) {
        warnings.push(`筛选实际覆盖 ${appliedFilterRange}，与传入的 ${range} 不一致：宿主会把筛选扩展到相邻数据块。若需精确范围，请在目标区与其它数据之间留一个空行。`);
      }
    }

    // 数据排序：旧式 Range.Sort(...) 在本机 WPS 上会静默 no-op（问题台账 ISS-38），
    // 因此先走 SortFields，再读回校验，都无效时**报错而不是返回假成功**。
    let sortApplied = null;
    if (Array.isArray(sortRules) && sortRules.length > 0) {
      const rowCount = targetRange.Rows.Count;
      const colCount = targetRange.Columns.Count;
      const before = normalize2DArray(targetRange.Value2, rowCount, colCount);
      const attempts = [];
      let after = before;
      let sorted = false;

      const readCurrent = () => normalize2DArray(targetRange.Value2, rowCount, colCount);
      const sortKeyOf = (rule) => targetRange.Columns.Item(Number(rule.colIndex));
      const orderOf = (rule) => (rule.order === "desc" ? 2 : 1);

      // 本机 WPS 实测：`Range.Sort` 是**方法**而不是对象（访问 .SortFields 抛
      // "Cannot read properties of undefined (reading 'Clear')"），而 `sheet.Sort` 才是可用的 Sort 对象。
      // 因此按下面顺序逐条尝试，每一条都**读回校验**，全都不生效就报错（问题台账 ISS-38）。
      const sortPaths = [
        {
          name: "sheet.Sort.SortFields",
          run: () => {
            const s = sheet.Sort;
            s.SortFields.Clear();
            for (let i = 0; i < sortRules.length; i++) s.SortFields.Add(sortKeyOf(sortRules[i]), 0, orderOf(sortRules[i]));
            s.SetRange(targetRange);
            s.Header = 1; // 包含表头
            s.Apply();
          }
        },
        {
          name: "range.Sort.SortFields",
          run: () => {
            const s = targetRange.Sort;
            s.SortFields.Clear();
            for (let i = 0; i < sortRules.length; i++) s.SortFields.Add(sortKeyOf(sortRules[i]), 0, orderOf(sortRules[i]));
            s.SetRange(targetRange);
            s.Header = 1;
            s.Apply();
          }
        },
        {
          name: "range.Sort(旧式)",
          run: () => {
            const key1 = sortKeyOf(sortRules[0]), order1 = orderOf(sortRules[0]);
            let key2, order2;
            if (sortRules[1]) {
              key2 = sortKeyOf(sortRules[1]);
              order2 = orderOf(sortRules[1]);
            }
            targetRange.Sort(key1, order1, key2, order2, undefined, undefined, 1);
          }
        }
      ];

      for (let p = 0; p < sortPaths.length; p++) {
        const path = sortPaths[p];
        let error = null;
        try {
          path.run();
        } catch (e) {
          error = e.message;
        }
        // 即使抛错也读回：某条路径可能已经部分生效
        after = readCurrent();
        sorted = isSortedByRules(after, sortRules);
        attempts.push(path.name + (error ? ":异常(" + error + ")" : ":ok") + (sorted ? ":已生效" : ":未生效"));
        if (sorted) break;
      }

      if (!sorted) {
        throw new Error(
          `排序未生效：区域 ${range} 读回后的顺序不满足请求的排序规则。已尝试：${attempts.join("；")}。` +
          `请核对 colIndex 是否为**区域内相对列号**（1 = 区域第一列），或改用 wps_execute_script 的 sheet.Sort.SortFields 路径。`
        );
      }
      sortApplied = { changed: JSON.stringify(before) !== JSON.stringify(after), attempts };
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      range,
      enableAutoFilter: enableAutoFilter !== undefined ? !!enableAutoFilter : sheet.AutoFilterMode,
      appliedFilterRange,
      sortedRuleCount: Array.isArray(sortRules) ? sortRules.length : 0,
      sortApplied,
      warnings,
      message: `已成功在 [${sheet.Name}] ${range} 应用筛选与排序`
    };
  }

  // 18. 单元格下拉验证菜单 (第三梯队)
  function setDataValidation(app, params) {
    const {
      sheetName,
      workbookName,
      address,
      validationType = "list",
      listItems = [],
      operator = "between",
      minVal,
      maxVal,
      promptTitle,
      promptMessage,
      errorTitle,
      errorMessage
    } = params || {};

    if (!address) throw new Error("缺少必要参数: address (例如 'E5:E20')");

    const sheet = getWorksheet(app, sheetName, workbookName);
    const targetRange = sheet.Range(address);

    try {
      targetRange.Validation.Delete();
    } catch (e) {}

    if (validationType === "list") {
      const listStr = Array.isArray(listItems) ? listItems.join(",") : String(listItems || "");
      // Type: 3 (xlValidateList), AlertStyle: 1 (xlValidAlertStop), Operator: 1 (xlBetween)
      targetRange.Validation.Add(3, 1, 1, listStr);
      targetRange.Validation.InCellDropdown = true;
    } else if (validationType === "number_range") {
      let op = 1; // xlBetween
      if (operator === "greater_than") op = 5;
      else if (operator === "less_than") op = 6;
      else if (operator === "equal") op = 3;

      const f1 = minVal !== undefined ? String(minVal) : "0";
      const f2 = maxVal !== undefined ? String(maxVal) : undefined;
      // Type: 2 (xlValidateDecimal)
      targetRange.Validation.Add(2, 1, op, f1, f2);
    } else {
      throw new Error(`未知的 validationType: ${validationType} (支持 list, number_range)`);
    }

    if (promptMessage) {
      targetRange.Validation.InputTitle = promptTitle || "选择提示";
      targetRange.Validation.InputMessage = promptMessage;
      targetRange.Validation.ShowInput = true;
    }

    if (errorMessage) {
      targetRange.Validation.ErrorTitle = errorTitle || "输入无效";
      targetRange.Validation.ErrorMessage = errorMessage;
      targetRange.Validation.ShowError = true;
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address,
      validationType,
      message: `已成功在 [${sheet.Name}] ${address} 配置数据有效性验证`
    };
  }

  // 19. 工作表综合管理 (第四梯队)
  function manageSheet(app, params) {
    const { sheetName, workbookName, action, newName, targetIndex, color, password } = params || {};
    if (!sheetName) throw new Error("缺少必要参数: sheetName");
    if (!action) throw new Error("缺少必要参数: action (rename, move, tab_color, protect, unprotect)");

    const wb = getWorkbook(app, workbookName);
    const sheet = getWorksheet(app, sheetName, workbookName);

    switch (action) {
      case "rename": {
        if (!newName) throw new Error("rename 操作必须提供 newName");
        const oldName = sheet.Name;
        sheet.Name = newName;
        return { success: true, workbookName: wb.Name, oldName, newName, message: `工作表 [${oldName}] 已成功重命名为 [${newName}]` };
      }
      case "move": {
        if (!targetIndex) throw new Error("move 操作必须提供 targetIndex (1-indexed)");
        const total = wb.Worksheets.Count;
        const validIndex = Math.max(1, Math.min(Number(targetIndex), total));
        const refSheet = wb.Worksheets.Item(validIndex);
        sheet.Move(refSheet);
        return { success: true, workbookName: wb.Name, sheetName: sheet.Name, targetIndex: validIndex, message: `工作表 [${sheet.Name}] 已成功移动至位置 ${validIndex}` };
      }
      case "tab_color": {
        if (!color) throw new Error("tab_color 操作必须提供十六进制颜色 (例如 '#0F172A')");
        const bgr = hexToExcelColor(color);
        if (bgr !== null) {
          sheet.Tab.Color = bgr;
        }
        return { success: true, workbookName: wb.Name, sheetName: sheet.Name, tabColor: color, message: `工作表 [${sheet.Name}] 标签底色已设置为 ${color}` };
      }
      case "protect": {
        sheet.Protect(password || undefined);
        return { success: true, workbookName: wb.Name, sheetName: sheet.Name, protected: true, message: `工作表 [${sheet.Name}] 已锁定保护` };
      }
      case "unprotect": {
        sheet.Unprotect(password || undefined);
        return { success: true, workbookName: wb.Name, sheetName: sheet.Name, protected: false, message: `工作表 [${sheet.Name}] 已解除锁定保护` };
      }
      default:
        throw new Error(`未知的 Sheet 操作: ${action} (支持 rename, move, tab_color, protect, unprotect)`);
    }
  }

  // ---------------------------------------------------------------------------
  // word.js — WPS 文字（Word/WPS）全部 RPC 实现
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
  // ==========================================
  // Word (文字) 模块 工业级全套核心操作实现
  // ==========================================

  function wordCreateDocument(app, params) {
    const { templatePath, isVisible } = params || {};
    const wordApp = getWordApp() || app || (typeof Application !== "undefined" ? Application : null);
    if (!wordApp) throw new Error("WPS 文字 (Word) 未就绪");
    const doc = templatePath ? wordApp.Documents.Add(templatePath) : wordApp.Documents.Add();
    try {
      if (doc.Activate) doc.Activate();
    } catch (e) {}
    return {
      success: true,
      documentName: doc.Name,
      fullName: doc.FullName || doc.Name,
      message: `已成功创建新文档: ${doc.Name}`
    };
  }

  function wordSaveDocument(app, params) {
    const { documentName, filePath, format } = params || {};
    const doc = getWordDocument(app, documentName);
    const fmt = (format || "").toLowerCase();

    if (fmt === "pdf" || (filePath && filePath.toLowerCase().endsWith(".pdf"))) {
      if (!filePath) throw new Error("导出 PDF 必须指定 filePath 完整保存路径");
      // 17 = wdExportFormatPDF
      doc.ExportAsFixedFormat(filePath, 17);
      return {
        success: true,
        documentName: doc.Name,
        savedPath: filePath,
        format: "pdf",
        message: `文档 [${doc.Name}] 已成功导出为 PDF: ${filePath}`
      };
    }

    if (filePath) {
      try {
        doc.SaveAs2(filePath);
      } catch (e) {
        doc.SaveAs(filePath);
      }
      return {
        success: true,
        documentName: doc.Name,
        savedPath: filePath,
        format: "docx",
        message: `文档 [${doc.Name}] 已成功另存为: ${filePath}`
      };
    }

    doc.Save();
    return {
      success: true,
      documentName: doc.Name,
      savedPath: doc.FullName || doc.Name,
      format: "docx",
      message: `文档 [${doc.Name}] 已成功原地保存`
    };
  }

  function wordCloseDocument(app, params) {
    const { documentName, saveChanges } = params || {};
    const doc = getWordDocument(app, documentName);
    const docName = doc.Name;
    let saveFlag = 0; // 0 = wdDoNotSaveChanges
    if (saveChanges === true) saveFlag = -1; // -1 = wdSaveChanges
    doc.Close(saveFlag);
    return {
      success: true,
      documentName: docName,
      saveChanges: Boolean(saveChanges),
      message: `文档 [${docName}] 已成功关闭 (saveChanges: ${Boolean(saveChanges)})`
    };
  }

  function wordManageContent(app, params) {
    const { documentName, action, paragraphIndex, paragraphRange, tableIndex } = params || {};
    const doc = getWordDocument(app, documentName);

    switch (action) {
      case "delete_paragraph": {
        if (paragraphRange && Array.isArray(paragraphRange) && paragraphRange.length === 2) {
          const startP = Number(paragraphRange[0]);
          const endP = Number(paragraphRange[1]);
          const startPos = doc.Paragraphs.Item(startP).Range.Start;
          const endPos = doc.Paragraphs.Item(endP).Range.End;
          const r = doc.Range(startPos, endPos);
          r.Delete();
          return {
            success: true,
            documentName: doc.Name,
            action: "delete_paragraph",
            deletedRange: [startP, endP],
            message: `已成功删除段落 P${startP} ~ P${endP}`
          };
        }
        if (!paragraphIndex) throw new Error("delete_paragraph 操作必须提供 paragraphIndex 或 paragraphRange");
        const idx = Number(paragraphIndex);
        const p = doc.Paragraphs.Item(idx);
        p.Range.Delete();
        return {
          success: true,
          documentName: doc.Name,
          action: "delete_paragraph",
          deletedIndex: idx,
          message: `已成功删除第 ${idx} 个段落`
        };
      }
      case "delete_table": {
        const tIdx = Number(tableIndex) || 1;
        if (tIdx > doc.Tables.Count) throw new Error(`表格索引超出范围: 当前仅有 ${doc.Tables.Count} 个表格`);
        const table = doc.Tables.Item(tIdx);
        table.Delete();
        return {
          success: true,
          documentName: doc.Name,
          action: "delete_table",
          deletedTableIndex: tIdx,
          message: `已成功删除第 ${tIdx} 个表格`
        };
      }
      case "clear_all": {
        doc.Content.Delete();
        return {
          success: true,
          documentName: doc.Name,
          action: "clear_all",
          message: `已成功清空文档 [${doc.Name}] 全部内容`
        };
      }
      default:
        throw new Error(`未知的 Word 内容管理操作: ${action} (支持 delete_paragraph, delete_table, clear_all)`);
    }
  }

  function wordReadDocument(app, params) {
    const { documentName, scope, maxParagraphs, includeFormatting, includeTables } = params || {};
    const doc = getWordDocument(app, documentName);
    const outline = [];
    const maxP = Number(maxParagraphs) || 200;

    try {
      const paraCount = doc.Paragraphs.Count;
      for (let i = 1; i <= Math.min(paraCount, 200); i++) {
        const para = doc.Paragraphs.Item(i);
        const text = (para.Range.Text || "").trim();
        if (!text) continue;
        const level = para.OutlineLevel;
        if (level >= 1 && level <= 3) {
          outline.push({ level, text: text.replace(/[\r\n\x07]/g, ""), paragraphIndex: i });
        }
      }
    } catch (e) {}

    let previewText = "";
    const paragraphDetails = [];
    if (scope === "full" || scope === "paragraphs" || !scope) {
      try {
        const count = Math.min(doc.Paragraphs.Count, maxP);
        const snippets = [];
        for (let i = 1; i <= count; i++) {
          const para = doc.Paragraphs.Item(i);
          const t = (para.Range.Text || "").trim();
          const cleanText = t.replace(/[\r\n\x07]/g, "");
          snippets.push(`[P${i}] ${cleanText}`);

          if (includeFormatting) {
            let isBold = false;
            let isItalic = false;
            let fSize = 12;
            let fName = "";
            let align = 0;
            try {
              isBold = para.Range.Font.Bold === true || para.Range.Font.Bold === -1;
              isItalic = para.Range.Font.Italic === true || para.Range.Font.Italic === -1;
              fSize = para.Range.Font.Size;
              fName = para.Range.Font.NameFarEast || para.Range.Font.Name;
              align = para.Format.Alignment;
            } catch (fe) {}

            paragraphDetails.push({
              index: i,
              text: cleanText,
              bold: isBold,
              italic: isItalic,
              fontSize: fSize,
              fontName: fName,
              alignment: align,
              outlineLevel: para.OutlineLevel
            });
          }
        }
        previewText = snippets.join("\n");
      } catch (e) {}
    }

    const tablesSummary = [];
    if (includeTables !== false && doc.Tables) {
      try {
        const tCount = doc.Tables.Count;
        for (let t = 1; t <= tCount; t++) {
          const tbl = doc.Tables.Item(t);
          const rCount = tbl.Rows.Count;
          const cCount = tbl.Columns.Count;
          const previewRows = [];
          for (let r = 1; r <= Math.min(rCount, 3); r++) {
            const rowData = [];
            for (let c = 1; c <= Math.min(cCount, 10); c++) {
              try {
                rowData.push((tbl.Cell(r, c).Range.Text || "").trim().replace(/[\r\n\x07]/g, ""));
              } catch (ce) {
                rowData.push("");
              }
            }
            previewRows.push(rowData);
          }
          tablesSummary.push({
            tableIndex: t,
            rowCount: rCount,
            columnCount: cCount,
            preview: previewRows
          });
        }
      } catch (te) {}
    }

    let selectionText = "";
    try {
      const wordApp = getWordApp() || app;
      if (wordApp.Selection && wordApp.Selection.Range) {
        selectionText = (wordApp.Selection.Range.Text || "").trim().replace(/[\r\n\x07]/g, "");
      }
    } catch (e) {}

    return {
      documentName: doc.Name,
      fullName: doc.FullName || doc.Name,
      paragraphCount: doc.Paragraphs ? doc.Paragraphs.Count : 0,
      tableCount: doc.Tables ? doc.Tables.Count : 0,
      wordCount: doc.Words ? doc.Words.Count : 0,
      outline,
      tables: tablesSummary,
      paragraphDetails: includeFormatting ? paragraphDetails : undefined,
      previewText: scope === "outline" ? undefined : previewText,
      selectionText: selectionText || undefined
    };
  }

  function wordWriteContent(app, params) {
    const { documentName, location, targetBookmark, paragraphIndex, type, content, formatting } = params || {};
    const doc = getWordDocument(app, documentName);
    const wordApp = getWordApp() || app;
    let targetRange = null;

    if (location === "start") {
      targetRange = doc.Range(0, 0);
    } else if (location === "selection" && wordApp.Selection) {
      targetRange = wordApp.Selection.Range;
    } else if (location === "bookmark" && targetBookmark) {
      if (doc.Bookmarks.Exists(targetBookmark)) {
        targetRange = doc.Bookmarks.Item(targetBookmark).Range;
      } else {
        targetRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
      }
    } else if (location === "after_paragraph" && paragraphIndex) {
      const p = doc.Paragraphs.Item(Number(paragraphIndex));
      targetRange = doc.Range(p.Range.End, p.Range.End);
    } else {
      const endPos = doc.Content.End > 1 ? doc.Content.End - 1 : 0;
      targetRange = doc.Range(endPos, endPos);
    }

    const lines = Array.isArray(content) ? content : [String(content || "")];
    for (const text of lines) {
      if (!text) continue;
      const newPara = doc.Paragraphs.Add(targetRange);
      newPara.Range.Text = text + "\n";

      if (type === "heading1") {
        try { newPara.Range.Style = doc.Styles.Item(-2); } catch (e) { newPara.Range.Font.Bold = true; newPara.Range.Font.Size = 22; }
      } else if (type === "heading2") {
        try { newPara.Range.Style = doc.Styles.Item(-3); } catch (e) { newPara.Range.Font.Bold = true; newPara.Range.Font.Size = 16; }
      } else if (type === "heading3") {
        try { newPara.Range.Style = doc.Styles.Item(-4); } catch (e) { newPara.Range.Font.Bold = true; newPara.Range.Font.Size = 14; }
      } else if (type === "bullet_list") {
        try { newPara.Range.ListFormat.ApplyBulletDefault(); } catch (e) {}
      } else if (type === "quote") {
        try {
          newPara.Range.Font.Italic = true;
          newPara.Format.LeftIndent = 28;
        } catch (e) {}
      }

      // 精确控制排版并主动断开加粗继承污染
      if (formatting && typeof formatting === "object") {
        if (formatting.bold !== undefined) newPara.Range.Font.Bold = Boolean(formatting.bold);
        if (formatting.italic !== undefined) newPara.Range.Font.Italic = Boolean(formatting.italic);
        if (formatting.fontSizePt !== undefined) newPara.Range.Font.Size = Number(formatting.fontSizePt);
        if (formatting.fontName) {
          newPara.Range.Font.NameFarEast = formatting.fontName;
          newPara.Range.Font.NameAscii = formatting.fontName;
        }
        if (formatting.alignment !== undefined) newPara.Format.Alignment = Number(formatting.alignment);
        if (formatting.firstLineIndentChars !== undefined) newPara.Format.CharacterUnitFirstLineIndent = Number(formatting.firstLineIndentChars);
        if (formatting.lineSpacingPt !== undefined) {
          newPara.Format.LineSpacingRule = 4;
          newPara.Format.LineSpacing = Number(formatting.lineSpacingPt);
        }
        if (formatting.spaceBeforePt !== undefined) newPara.Format.SpaceBefore = Number(formatting.spaceBeforePt);
        if (formatting.spaceAfterPt !== undefined) newPara.Format.SpaceAfter = Number(formatting.spaceAfterPt);
      } else if (!type || type === "paragraph") {
        // 普通正文默认强制消除前文加粗继承
        try {
          newPara.Range.Font.Bold = false;
        } catch (e) {}
      }

      targetRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
    }

    return {
      success: true,
      documentName: doc.Name,
      insertedLines: lines.length,
      location: location || "end",
      type: type || "paragraph",
      message: `已成功向 [${doc.Name}] 写入 ${lines.length} 行内容`
    };
  }

  function wordFormatDocument(app, params) {
    const { documentName, target, paragraphIndex, paragraphRange, searchQuery, searchQueries, preset, fontName, fontSizePt, bold, italic, lineSpacingPt, firstLineIndentChars, spaceBeforePt, spaceAfterPt, margins, alignment } = params || {};
    const doc = getWordDocument(app, documentName);
    const wordApp = getWordApp() || app;

    try {
      if (preset === "gov_standard") {
        doc.PageSetup.TopMargin = 37 * 2.83465;
        doc.PageSetup.BottomMargin = 35 * 2.83465;
        doc.PageSetup.LeftMargin = 28 * 2.83465;
        doc.PageSetup.RightMargin = 26 * 2.83465;
      } else if (margins) {
        if (margins.topMm) doc.PageSetup.TopMargin = margins.topMm * 2.83465;
        if (margins.bottomMm) doc.PageSetup.BottomMargin = margins.bottomMm * 2.83465;
        if (margins.leftMm) doc.PageSetup.LeftMargin = margins.leftMm * 2.83465;
        if (margins.rightMm) doc.PageSetup.RightMargin = margins.rightMm * 2.83465;
      }
    } catch (e) {}

    function applyFormatToPara(para) {
      if (bold !== undefined) para.Range.Font.Bold = Boolean(bold);
      if (italic !== undefined) para.Range.Font.Italic = Boolean(italic);
      if (fontName) {
        para.Range.Font.NameFarEast = fontName;
        para.Range.Font.NameAscii = fontName;
      }
      if (fontSizePt !== undefined) para.Range.Font.Size = Number(fontSizePt);
      if (alignment !== undefined) para.Format.Alignment = Number(alignment);
      if (lineSpacingPt !== undefined) {
        para.Format.LineSpacingRule = 4;
        para.Format.LineSpacing = Number(lineSpacingPt);
      }
      if (firstLineIndentChars !== undefined) {
        para.Format.CharacterUnitFirstLineIndent = Number(firstLineIndentChars);
      }
      if (spaceBeforePt !== undefined) para.Format.SpaceBefore = Number(spaceBeforePt);
      if (spaceAfterPt !== undefined) para.Format.SpaceAfter = Number(spaceAfterPt);
    }

    // 1. 基于搜索词强力加粗/排版 (支持正文与全量表格穿透)
    const queries = Array.isArray(searchQueries) ? searchQueries : (searchQuery ? [searchQuery] : []);
    if (queries.length > 0) {
      const ranges = [doc.Content];
      if (doc.Tables) {
        for (let t = 1; t <= doc.Tables.Count; t++) {
          try { ranges.push(doc.Tables.Item(t).Range); } catch (te) {}
        }
      }
      let formattedCount = 0;
      for (const q of queries) {
        if (!q) continue;
        for (const rng of ranges) {
          try {
            const f = rng.Find;
            f.ClearFormatting();
            f.Text = q;
            f.MatchCase = false;
            f.MatchWholeWord = false;
            f.MatchWildcards = false;
            f.Forward = true;
            f.Wrap = 0; // wdFindStop
            while (f.Execute()) {
              formattedCount++;
              if (bold !== undefined) f.Parent.Font.Bold = Boolean(bold);
              if (italic !== undefined) f.Parent.Font.Italic = Boolean(italic);
              if (fontSizePt !== undefined) f.Parent.Font.Size = Number(fontSizePt);
              if (fontName) {
                f.Parent.Font.NameFarEast = fontName;
                f.Parent.Font.NameAscii = fontName;
              }
            }
          } catch (fe) {}
        }
      }
      return {
        success: true,
        documentName: doc.Name,
        target: "search_matches",
        queries,
        formattedMatches: formattedCount,
        message: `已成功为 ${queries.length} 个关键词匹配项 (${formattedCount} 处) 应用排版`
      };
    }

    if (target === "paragraph" || paragraphIndex) {
      const pIdx = Number(paragraphIndex || 1);
      const p = doc.Paragraphs.Item(pIdx);
      applyFormatToPara(p);
      return {
        success: true,
        documentName: doc.Name,
        target: `paragraph_${pIdx}`,
        message: `已成功格式化第 ${pIdx} 个段落`
      };
    }

    if (target === "range" || (paragraphRange && Array.isArray(paragraphRange))) {
      const startP = Number(paragraphRange[0]);
      const endP = Number(paragraphRange[1]);
      for (let i = startP; i <= endP; i++) {
        try {
          applyFormatToPara(doc.Paragraphs.Item(i));
        } catch (e) {}
      }
      return {
        success: true,
        documentName: doc.Name,
        target: `range_${startP}_${endP}`,
        message: `已成功格式化段落 P${startP} ~ P${endP}`
      };
    }

    if (target === "selection" && wordApp.Selection && wordApp.Selection.Range) {
      const rng = wordApp.Selection.Range;
      if (bold !== undefined) rng.Font.Bold = Boolean(bold);
      if (italic !== undefined) rng.Font.Italic = Boolean(italic);
      if (fontName) {
        rng.Font.NameFarEast = fontName;
        rng.Font.NameAscii = fontName;
      }
      if (fontSizePt !== undefined) rng.Font.Size = Number(fontSizePt);
      return {
        success: true,
        documentName: doc.Name,
        target: "selection",
        message: `已成功格式化当前选区文字`
      };
    }

    const paraCount = doc.Paragraphs.Count;
    for (let i = 1; i <= paraCount; i++) {
      try {
        const para = doc.Paragraphs.Item(i);
        const level = para.OutlineLevel;

        if (preset === "gov_standard") {
          if (level >= 10 || level === 0) {
            para.Range.Font.NameFarEast = "仿宋_GB2312";
            para.Range.Font.NameAscii = "仿宋_GB2312";
            para.Range.Font.Size = 16;
            para.Format.LineSpacingRule = 4;
            para.Format.LineSpacing = 28;
            para.Format.CharacterUnitFirstLineIndent = 2;
            para.Format.SpaceBefore = 0;
            para.Format.SpaceAfter = 0;
          } else if (level === 1) {
            para.Range.Font.NameFarEast = "黑体";
            para.Range.Font.Size = 16;
            para.Range.Font.Bold = false;
            para.Format.LineSpacingRule = 4;
            para.Format.LineSpacing = 28;
            para.Format.CharacterUnitFirstLineIndent = 2;
          } else if (level === 2) {
            para.Range.Font.NameFarEast = "楷体_GB2312";
            para.Range.Font.Size = 16;
            para.Range.Font.Bold = false;
            para.Format.LineSpacingRule = 4;
            para.Format.LineSpacing = 28;
            para.Format.CharacterUnitFirstLineIndent = 2;
          }
        } else if (preset === "business_modern") {
          para.Range.Font.NameFarEast = "微软雅黑";
          para.Range.Font.NameAscii = "Segoe UI";
          if (level >= 10 || level === 0) {
            para.Range.Font.Size = 11;
            para.Format.LineSpacingRule = 5;
            para.Format.SpaceAfter = 6;
          }
        } else {
          applyFormatToPara(para);
        }
      } catch (e) {}
    }

    return {
      success: true,
      documentName: doc.Name,
      preset: preset || "custom",
      formattedParagraphs: paraCount,
      message: `已成功对 [${doc.Name}] 应用 [${preset || "自定义"}] 排版规范`
    };
  }

  function wordInsertTableOfContents(app, params) {
    const { documentName, upperHeadingLevel, lowerHeadingLevel, insertLocation, includePageNumbers } = params || {};
    const doc = getWordDocument(app, documentName);
    let targetRange = null;

    if (insertLocation === "selection" && app.Selection) {
      targetRange = app.Selection.Range;
    } else {
      targetRange = doc.Range(0, 0);
    }

    const upper = Number(upperHeadingLevel) || 1;
    const lower = Number(lowerHeadingLevel) || 3;

    try {
      doc.TablesOfContents.Add(
        targetRange,
        true,
        upper,
        lower,
        true,
        undefined,
        true,
        includePageNumbers !== false
      );
    } catch (e) {
      const p = doc.Paragraphs.Add(targetRange);
      p.Range.Text = "【目录】\n";
      p.Range.Font.Bold = true;
    }

    return {
      success: true,
      documentName: doc.Name,
      headingLevels: `${upper}-${lower}`,
      message: `已成功在 [${doc.Name}] 插入标准目录`
    };
  }

  function wordManageTable(app, params) {
    const { documentName, action, tableIndex, rows, columns, data, stylePreset, repeatHeader, mergeRange, cellRow, cellColumn, cellFormat, rowIndex, columnIndex } = params || {};
    const doc = getWordDocument(app, documentName);

    if (action === "inspect") {
      const idx = Number(tableIndex) || 1;
      if (idx > doc.Tables.Count) throw new Error(`表格索引超出范围: 当前仅有 ${doc.Tables.Count} 个表格`);
      const table = doc.Tables.Item(idx);
      const rCount = table.Rows.Count;
      const cCount = table.Columns.Count;
      const matrix = [];
      for (let r = 1; r <= rCount; r++) {
        const rowData = [];
        for (let c = 1; c <= cCount; c++) {
          try {
            rowData.push((table.Cell(r, c).Range.Text || "").trim().replace(/[\r\n\x07]/g, ""));
          } catch (ce) {
            rowData.push("");
          }
        }
        matrix.push(rowData);
      }
      return {
        success: true,
        documentName: doc.Name,
        tableIndex: idx,
        rowCount: rCount,
        columnCount: cCount,
        data: matrix,
        message: `表格 ${idx} 结构透视: ${rCount} 行 × ${cCount} 列`
      };
    }

    if (action === "insert") {
      const dataRows = Array.isArray(data) && data.length > 0 ? data.length : (Number(rows) || 3);
      const dataCols = Array.isArray(data) && data[0] && Array.isArray(data[0]) ? data[0].length : (Number(columns) || 3);

      const targetRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
      const table = doc.Tables.Add(targetRange, dataRows, dataCols);

      if (Array.isArray(data)) {
        for (let r = 0; r < data.length; r++) {
          for (let c = 0; c < data[r].length; c++) {
            try {
              table.Cell(r + 1, c + 1).Range.Text = String(data[r][c] !== undefined ? data[r][c] : "");
            } catch (e) {}
          }
        }
      }

      if (stylePreset === "mckinsey_three_line" || !stylePreset) {
        try {
          table.Borders.Enable = false;
          table.Borders.Item(-1).LineStyle = 1;
          table.Borders.Item(-1).LineWidth = 12;
          table.Borders.Item(-3).LineStyle = 1;
          table.Borders.Item(-3).LineWidth = 12;
          if (table.Rows.Count > 1) {
            table.Rows.Item(1).Borders.Item(-3).LineStyle = 1;
            table.Rows.Item(1).Borders.Item(-3).LineWidth = 6;
            table.Rows.Item(1).Range.Font.Bold = true;
          }
        } catch (e) {}
      }

      if (repeatHeader !== false && table.Rows.Count > 0) {
        try {
          table.Rows.Item(1).HeadingFormat = true;
        } catch (e) {}
      }

      return {
        success: true,
        documentName: doc.Name,
        tableIndex: doc.Tables.Count,
        rows: dataRows,
        columns: dataCols,
        style: stylePreset || "mckinsey_three_line",
        message: `已成功在 [${doc.Name}] 插入 ${dataRows}行 × ${dataCols}列 专业三线表`
      };
    }

    if (action === "update_data" || action === "write_matrix") {
      const idx = Number(tableIndex) || 1;
      if (idx > doc.Tables.Count) throw new Error(`表格索引超出范围: 当前仅有 ${doc.Tables.Count} 个表格`);
      const table = doc.Tables.Item(idx);

      if (Array.isArray(data)) {
        const reqRows = data.length;
        // 动态扩容行数，绝不静默丢弃任何行
        while (table.Rows.Count < reqRows) {
          table.Rows.Add();
        }

        for (let r = 0; r < data.length; r++) {
          const rowArr = data[r];
          if (!Array.isArray(rowArr)) continue;
          for (let c = 0; c < rowArr.length; c++) {
            try {
              const cell = table.Cell(r + 1, c + 1);
              cell.Range.Text = String(rowArr[c] !== undefined && rowArr[c] !== null ? rowArr[c] : "");
              // 消除意外的粗体继承污染
              if (r > 0) {
                cell.Range.Font.Bold = false;
              }
            } catch (e) {}
          }
        }
      }
      return {
        success: true,
        documentName: doc.Name,
        tableIndex: idx,
        rowCount: table.Rows.Count,
        columnCount: table.Columns.Count,
        message: `表格 ${idx} 数据已安全更新 (${data ? data.length : 0} 行)`
      };
    }

    if (action === "format_cell") {
      const idx = Number(tableIndex) || 1;
      const table = doc.Tables.Item(idx);
      const r = Number(cellRow || rowIndex || 1);
      const c = Number(cellColumn || columnIndex || 1);
      const cell = table.Cell(r, c);
      if (cellFormat && typeof cellFormat === "object") {
        if (cellFormat.bold !== undefined) cell.Range.Font.Bold = Boolean(cellFormat.bold);
        if (cellFormat.fontSizePt !== undefined) cell.Range.Font.Size = Number(cellFormat.fontSizePt);
        if (cellFormat.fontName) cell.Range.Font.NameFarEast = cellFormat.fontName;
        if (cellFormat.backgroundColor) {
          const bgr = hexToExcelColor(cellFormat.backgroundColor);
          if (bgr !== null) cell.Shading.BackgroundPatternColor = bgr;
        }
      }
      return {
        success: true,
        documentName: doc.Name,
        tableIndex: idx,
        cell: [r, c],
        message: `表格 ${idx} 单元格 [${r}, ${c}] 格式已更新`
      };
    }

    if (action === "add_row") {
      const idx = Number(tableIndex) || 1;
      const table = doc.Tables.Item(idx);
      table.Rows.Add();
      return { success: true, documentName: doc.Name, tableIndex: idx, rowCount: table.Rows.Count, message: `表格 ${idx} 已追加新行` };
    }

    if (action === "delete_row") {
      const idx = Number(tableIndex) || 1;
      const table = doc.Tables.Item(idx);
      const r = Number(rowIndex || table.Rows.Count);
      table.Rows.Item(r).Delete();
      return { success: true, documentName: doc.Name, tableIndex: idx, rowCount: table.Rows.Count, message: `表格 ${idx} 第 ${r} 行已删除` };
    }

    if (action === "merge_cells" && mergeRange) {
      const idx = Number(tableIndex) || 1;
      const table = doc.Tables.Item(idx);
      const startCell = table.Cell(mergeRange.startRow, mergeRange.startCol);
      const endCell = table.Cell(mergeRange.endRow, mergeRange.endCol);
      startCell.Merge(endCell);
      return { success: true, documentName: doc.Name, tableIndex: idx, message: `表格 ${idx} 单元格已合并` };
    }

    throw new Error(`未知的 Word 表格操作: ${action} (支持 inspect, insert, update_data, write_matrix, format_cell, add_row, delete_row, merge_cells)`);
  }

  function wordReviewAndComments(app, params) {
    const { documentName, action, commentText, author } = params || {};
    const doc = getWordDocument(app, documentName);
    const wordApp = getWordApp() || app;

    switch (action) {
      case "enable_track_changes": {
        doc.TrackRevisions = true;
        return { success: true, documentName: doc.Name, trackRevisions: true, message: `文档 [${doc.Name}] 已开启修订记录模式` };
      }
      case "disable_track_changes": {
        doc.TrackRevisions = false;
        return { success: true, documentName: doc.Name, trackRevisions: false, message: `文档 [${doc.Name}] 已关闭修订记录模式` };
      }
      case "accept_all_revisions": {
        doc.AcceptAllRevisions();
        return { success: true, documentName: doc.Name, message: `已接受文档 [${doc.Name}] 中的全部修订` };
      }
      case "reject_all_revisions": {
        doc.RejectAllRevisions();
        return { success: true, documentName: doc.Name, message: `已拒绝文档 [${doc.Name}] 中的全部修订` };
      }
      case "add_comment": {
        if (!commentText) throw new Error("add_comment 操作必须提供 commentText");
        const range = (wordApp.Selection && wordApp.Selection.Range.Text) ? wordApp.Selection.Range : doc.Range(0, 0);
        const comment = doc.Comments.Add(range, commentText);
        if (author) comment.Author = author;
        return { success: true, documentName: doc.Name, commentId: doc.Comments.Count, message: `已成功添加批注: "${commentText}"` };
      }
      case "list_comments": {
        const list = [];
        const count = doc.Comments ? doc.Comments.Count : 0;
        for (let i = 1; i <= count; i++) {
          const c = doc.Comments.Item(i);
          list.push({
            index: i,
            author: c.Author,
            text: (c.Range.Text || "").trim(),
            scopeText: (c.Scope.Text || "").trim()
          });
        }
        return { success: true, documentName: doc.Name, total: count, comments: list };
      }
      default:
        throw new Error(`未知的 Word 审阅操作: ${action}`);
    }
  }

  function wordPageLayoutAndWatermark(app, params) {
    const { documentName, headerText, footerText, differentFirstPage, differentOddEvenPages, watermarkText } = params || {};
    const doc = getWordDocument(app, documentName);
    const section = doc.Sections.Item(1);

    if (differentFirstPage !== undefined) doc.PageSetup.DifferentFirstPageHeaderFooter = differentFirstPage;
    if (differentOddEvenPages !== undefined) doc.PageSetup.OddAndEvenPagesHeaderFooter = differentOddEvenPages;

    if (headerText !== undefined) {
      try { section.Headers.Item(1).Range.Text = headerText; } catch (e) {}
    }

    if (footerText !== undefined) {
      try { section.Footers.Item(1).Range.Text = footerText; } catch (e) {}
    }

    if (watermarkText) {
      try {
        const newShape = doc.Shapes.AddTextEffect(0, watermarkText, "Microsoft YaHei", 54, false, false, 100, 200);
        newShape.Rotation = -315;
        newShape.Fill.Transparency = 0.85;
        newShape.Line.Visible = false;
        newShape.WrapFormat.Type = 3;
      } catch (e) {}
    }

    return {
      success: true,
      documentName: doc.Name,
      header: headerText,
      footer: footerText,
      watermark: watermarkText || undefined,
      message: `已成功更新 [${doc.Name}] 页面版式与水印`
    };
  }

  function wordFindAndReplace(app, params) {
    const { documentName, searchQuery, replaceText, matchCase, matchWholeWord, useWildcards, scope, replaceFormatting } = params || {};
    if (!searchQuery) throw new Error("缺少必要参数: searchQuery");
    const doc = getWordDocument(app, documentName);
    const wordApp = getWordApp() || app;

    let targetRanges = [];
    if (scope === "selection" && wordApp && wordApp.Selection && wordApp.Selection.Range) {
      targetRanges = [wordApp.Selection.Range];
    } else {
      targetRanges = [doc.Content];
      if (doc.Tables) {
        for (let t = 1; t <= doc.Tables.Count; t++) {
          try { targetRanges.push(doc.Tables.Item(t).Range); } catch (te) {}
        }
      }
    }

    let matchCount = 0;
    let replacedAny = false;
    const errors = [];
    const hasFmt = Boolean(replaceFormatting && typeof replaceFormatting === "object");

    for (const rng of targetRanges) {
      try {
        const findObj = rng.Find;
        findObj.ClearFormatting();
        findObj.Text = searchQuery;
        findObj.MatchCase = Boolean(matchCase);
        findObj.MatchWholeWord = Boolean(matchWholeWord);
        findObj.MatchWildcards = Boolean(useWildcards);
        findObj.Forward = true;
        findObj.Wrap = 0; // wdFindStop

        if (replaceText !== undefined) {
          findObj.Replacement.ClearFormatting();
          findObj.Replacement.Text = replaceText;

          if (hasFmt) {
            if (replaceFormatting.bold !== undefined) findObj.Replacement.Font.Bold = Boolean(replaceFormatting.bold);
            if (replaceFormatting.italic !== undefined) findObj.Replacement.Font.Italic = Boolean(replaceFormatting.italic);
            if (replaceFormatting.fontSizePt !== undefined) findObj.Replacement.Font.Size = Number(replaceFormatting.fontSizePt);
            if (replaceFormatting.fontName) {
              findObj.Replacement.Font.NameFarEast = replaceFormatting.fontName;
              findObj.Replacement.Font.NameAscii = replaceFormatting.fontName;
            }
          }

          // 第 9 个参数 Format 传入 hasFmt，使 Replacement.Font 格式化必定生效。
          // 返回值必须看：原实现忽略了它并无条件 matchCount++，导致"没命中"也报成功（问题台账 ISS-68）。
          const did = findObj.Execute(
            searchQuery,
            Boolean(matchCase),
            Boolean(matchWholeWord),
            Boolean(useWildcards),
            false,
            false,
            true,
            1,
            hasFmt,
            replaceText,
            2 // wdReplaceAll
          );
          if (did) {
            matchCount++;
            replacedAny = true;
          }
        } else {
          // 纯查找 / 查找并格式化：**原实现缺这一支**，所以只查找时 matchCount 恒为 0、
          // 文案却写"已找到并应用格式化"。这里真正遍历计数（带上限防死循环）。
          let guard = 0;
          while (findObj.Execute() && guard++ < 5000) {
            matchCount++;
            if (hasFmt) {
              if (replaceFormatting.bold !== undefined) findObj.Parent.Font.Bold = Boolean(replaceFormatting.bold);
              if (replaceFormatting.italic !== undefined) findObj.Parent.Font.Italic = Boolean(replaceFormatting.italic);
              if (replaceFormatting.fontSizePt !== undefined) findObj.Parent.Font.Size = Number(replaceFormatting.fontSizePt);
              if (replaceFormatting.fontName) {
                findObj.Parent.Font.NameFarEast = replaceFormatting.fontName;
                findObj.Parent.Font.NameAscii = replaceFormatting.fontName;
              }
            }
          }
        }
      } catch (e) {
        // 不再静默吞异常：把原始错误带回给调用方
        errors.push(e.message);
      }
    }

    if (replaceText !== undefined) {
      return {
        success: true,
        documentName: doc.Name,
        searchQuery,
        replaceText,
        action: replacedAny ? "replaced_all" : "no_match",
        matchCount,
        errors: errors.length ? errors : undefined,
        message: replacedAny
          ? `已将 [${doc.Name}] 中的 "${searchQuery}" 全文穿透替换为 "${replaceText}"（命中 ${matchCount} 段）`
          : `未在 [${doc.Name}] 中找到 "${searchQuery}"，未做任何替换`
      };
    }

    return {
      success: true,
      documentName: doc.Name,
      searchQuery,
      matchCount,
      errors: errors.length ? errors : undefined,
      message: hasFmt
        ? `已为 "${searchQuery}" 找到并应用格式化（${matchCount} 处）`
        : `在 [${doc.Name}] 中找到 "${searchQuery}" ${matchCount} 处`
    };
  }

  function wordCapturePreview(app, params) {
    const { documentName } = params || {};
    const doc = getWordDocument(app, documentName);
    const tempPdfPath = params.outputPath;
    if (!tempPdfPath) throw new Error("缺少 Bridge 指定的预览输出路径");
    try {
      // wdExportFormatPDF = 17
      doc.ExportAsFixedFormat(tempPdfPath, 17);
      return {
        success: true,
        documentName: doc.Name,
        pdfPath: tempPdfPath,
        hasPdf: true,
        message: `已成功导出 Word 文档页面快照 PDF: ${tempPdfPath}`
      };
    } catch (e) {
      log(`Word 页面导出异常: ${e.message}`);
      return {
        success: false,
        documentName: doc.Name,
        error: e.message
      };
    }
  }

  // ---------------------------------------------------------------------------
  // ppt-layout.js — PPT 布局几何/字号/容量计算（纯函数，不调用任何宿主 API）
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进。
  // 例外：文件末尾的 @build-strip 导出块只为 Node 单元测试直接 import 使用，
  //      构建拼接时由 scripts/build-wps-addon.mjs 整段移除，不会进入 addon-core.js。
  // ---------------------------------------------------------------------------

  // Office geometry and font sizes are points (72 pt = 1 inch).
  // 只读 pres.PageSetup，可传入普通对象（{ PageSetup: { SlideWidth, SlideHeight } }），
  // 因此本函数可在无宿主环境下直接单元测试。
  function pptPageSize(pres) {
    const width = Number(pres.PageSetup.SlideWidth);
    const height = Number(pres.PageSetup.SlideHeight);
    if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
      throw new Error("无法读取有效的 PPT 页面尺寸，停止布局；请检查 PageSetup");
    }
    return { pageWidth: width, pageHeight: height, unit: "pt" };
  }

  // 设计基准 720x405 (pt) 到真实页面尺寸的缩放系数；算法与拆分前 fitGeneratedPptShapes 内联写法一致。
  function pptScaleForPage(page) {
    const sx = page.pageWidth / 720, sy = page.pageHeight / 405;
    return { sx: sx, sy: sy, scale: Math.min(sx, sy) };
  }

  // 文本视觉宽度：全角 1、半角 0.6；与拆分前 linesAt 内的换算一致。
  function pptCountTextUnits(text) {
    return Array.from(text).reduce((n, ch) => n + (ch.charCodeAt(0) > 255 ? 1 : 0.6), 0);
  }

  // 指定字号下的预估行数：显式换行逐段累加，每段按内宽向上取整（最少 1 行）。
  function pptEstimateLines(text, size, innerWidth) {
    return text.split(/\r\n|\r|\n/).reduce((sum, line) => {
      const units = pptCountTextUnits(line);
      return sum + Math.max(1, Math.ceil(units * size / innerWidth));
    }, 0);
  }

  // 在 [minimum, preferred] 内按 scale 步长下调字号，直到估算高度不再溢出（或触底）。
  function pptFitFontSize(text, preferred, minimum, scale, innerWidth, innerHeight) {
    let size = preferred;
    while (size > minimum && pptEstimateLines(text, size, innerWidth) * size * 1.25 > innerHeight) {
      size = Math.max(minimum, size - scale);
    }
    return size;
  }

  // ---------------------------------------------------------------------------
  // ppt.js — PPT 宿主写入与 RPC 实现
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
  // ==========================================
  // PowerPoint (演示) 模块 7 大核心操作实现
  // ==========================================

  function hexToPptColor(hex) {
    if (!hex) return 0;
    const clean = hex.replace("#", "");
    if (clean.length !== 6) return 0;
    const r = parseInt(clean.substring(0, 2), 16);
    const g = parseInt(clean.substring(2, 4), 16);
    const b = parseInt(clean.substring(4, 6), 16);
    return (b << 16) | (g << 8) | r;
  }

  // Adapt only newly generated shapes; existing slide content stays untouched.
  // 几何/字号/容量计算全部委托给 ppt-layout.js 的纯函数，本函数只读写宿主对象。
  function fitGeneratedPptShapes(slide, firstIndex, page, warnings) {
    const ratio = pptScaleForPage(page);
    const sx = ratio.sx, sy = ratio.sy, scale = ratio.scale;
    for (let i = firstIndex; i <= slide.Shapes.Count; i++) {
      const shape = slide.Shapes.Item(i);
      shape.Left *= sx; shape.Top *= sy;
      shape.Width *= sx; shape.Height *= sy;
      if (!shape.HasTextFrame || !shape.TextFrame.HasText) continue;
      const frame = shape.TextFrame, range = frame.TextRange;
      frame.AutoSize = 0;
      frame.WordWrap = true;
      const preferred = Number(range.Font.Size) * scale;
      const minimum = Math.min(preferred, 12 * scale);
      const innerW = Math.max(1, shape.Width - Number(frame.MarginLeft || 0) - Number(frame.MarginRight || 0));
      const innerH = Math.max(1, shape.Height - Number(frame.MarginTop || 0) - Number(frame.MarginBottom || 0));
      const text = String(range.Text || "");
      const size = pptFitFontSize(text, preferred, minimum, scale, innerW, innerH);
      range.Font.Size = size;
      if (pptEstimateLines(text, size, innerW) * size * 1.25 > innerH) warnings.push({ shapeId: shape.Id, reason: "文字可能溢出，请缩短内容或扩大文本框并检查预览" });
    }
  }

  function addSlideHeader(slide, title, themeColor) {
    const titleBox = slide.Shapes.AddTextbox(1, 50, 30, 620, 50);
    titleBox.TextFrame.TextRange.Text = title || "主题要点";
    titleBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
    titleBox.TextFrame.TextRange.Font.Size = 24;
    titleBox.TextFrame.TextRange.Font.Bold = true;
    titleBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor(themeColor || "#0F4C81");

    try {
      const line = slide.Shapes.AddLine(50, 75, 120, 75);
      line.Line.ForeColor.RGB = hexToPptColor(themeColor || "#0F4C81");
      line.Line.Weight = 2.5;
    } catch (e) {}
  }

  function renderPptCards(slide, cards, colCount, themeColor, customTop, customH) {
    const cols = Math.max(1, Math.min(Number(colCount) || 3, 4));
    if (cards.length > cols) throw new Error("卡片数量超过分栏数，请拆分为多页，避免内容被截断");
    const count = cards.length;
    const totalW = 620;
    const startX = 50;
    const startY = customTop === undefined ? 100 : Number(customTop);
    const cardH = customH === undefined ? 260 : Number(customH);
    if (!Number.isFinite(startY) || !Number.isFinite(cardH) || startY < 0 || cardH <= 0 || startY + cardH > 405) throw new Error("卡片位置或高度超出页面");
    const gap = 16;
    const cardW = (totalW - (cols - 1) * gap) / cols;

    for (let i = 0; i < count; i++) {
      const item = cards[i];
      const x = startX + i * (cardW + gap);
      const accent = item.accentColor || themeColor || "#0F4C81";

      // 1. 底卡 (圆角矩形 5)
      try {
        const bgShape = slide.Shapes.AddShape(5, x, startY, cardW, cardH);
        bgShape.Fill.Solid();
        bgShape.Fill.ForeColor.RGB = hexToPptColor("#F8FAFC");
        bgShape.Line.ForeColor.RGB = hexToPptColor("#E2E8F0");
        bgShape.Line.Weight = 1;
      } catch (e) {}

      // 2. 顶条强调线
      try {
        const topBar = slide.Shapes.AddShape(1, x + 10, startY + 12, 36, 4);
        topBar.Fill.Solid();
        topBar.Fill.ForeColor.RGB = hexToPptColor(accent);
        topBar.Line.Visible = false;
      } catch (e) {}

      // 3. 卡片标签 (Tag)
      if (item.tag) {
        const tagBox = slide.Shapes.AddTextbox(1, x + 10, startY + 20, cardW - 20, 22);
        tagBox.TextFrame.TextRange.Text = item.tag.toUpperCase();
        tagBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
        tagBox.TextFrame.TextRange.Font.Size = 10;
        tagBox.TextFrame.TextRange.Font.Bold = true;
        tagBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor(accent);
      }

      // 4. 卡片标题
      const titleY = item.tag ? startY + 42 : startY + 22;
      const cardTitleBox = slide.Shapes.AddTextbox(1, x + 10, titleY, cardW - 20, 36);
      cardTitleBox.TextFrame.TextRange.Text = item.title || `要点 ${i + 1}`;
      cardTitleBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
      cardTitleBox.TextFrame.TextRange.Font.Size = cols >= 4 ? 14 : 16;
      cardTitleBox.TextFrame.TextRange.Font.Bold = true;
      cardTitleBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#0F172A");

      // 5. 卡片正文描述
      const descY = titleY + 36;
      const descBox = slide.Shapes.AddTextbox(1, x + 10, descY, cardW - 20, cardH - (descY - startY) - 10);
      descBox.TextFrame.TextRange.Text = item.description || "";
      descBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
      descBox.TextFrame.TextRange.Font.Size = cols >= 4 ? 11 : 12;
      descBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#475569");
      descBox.TextFrame.WordWrap = true;
    }
  }

  function renderPptChart(slide, chartSpec, left, top, width, height) {
    const { chartType, categories, series, title, hasLegend, showDataLabels } = chartSpec || {};
    let typeCode = 51; // xlColumnClustered 柱状图
    if (chartType === "line") typeCode = 4; // xlLine 折线图
    else if (chartType === "pie") typeCode = 5; // xlPie 饼图
    else if (chartType === "bar") typeCode = 57; // xlBarClustered 条形图
    else if (chartType === "column_stacked" || chartType === "stacked_column") typeCode = 52;
    else if (chartType === "bar_stacked" || chartType === "stacked_bar") typeCode = 58;
    else if (chartType === "bar_of_pie" || chartType === "pie_bar") typeCode = 68;
    else if (typeof chartType === "number") typeCode = chartType;

    let chartShape = null;
    try {
      chartShape = slide.Shapes.AddChart(typeCode, left, top, width, height);
    } catch (e) {
      try {
        chartShape = slide.Shapes.AddChart2(-1, typeCode, left, top, width, height);
      } catch (err) {
        throw new Error("当前宿主无法创建原生图表：" + err.message);
      }
    }

    try {
      const chart = chartShape.Chart;
      if (title) {
        try {
          chart.HasTitle = true;
          chart.ChartTitle.Text = title;
        } catch (e) { throw e; }
      }
      if (hasLegend !== undefined) {
        try { chart.HasLegend = Boolean(hasLegend); } catch (e) { throw e; }
      }

      if (Array.isArray(series) && series.length > 0) {
        const sc = chart.SeriesCollection();
        while (sc.Count > series.length) {
          try { sc.Item(sc.Count).Delete(); } catch (e) { throw e; }
        }
        for (let i = 0; i < series.length; i++) {
          const sData = series[i];
          let sObj = null;
          if (i + 1 <= sc.Count) {
            sObj = sc.Item(i + 1);
          } else {
            try { sObj = sc.NewSeries(); } catch (e) { throw e; }
          }
          if (!sObj) throw new Error("无法创建目标数据系列");

          if (sData.name) {
            try { sObj.Name = sData.name; } catch (e) { throw e; }
          }
          if (categories && categories.length > 0 && i === 0) {
            try { sObj.XValues = categories; } catch (e) { throw e; }
          }
          if (Array.isArray(sData.values)) {
            try { sObj.Values = sData.values; } catch (e) { throw e; }
          }
          if (sData.chartType) {
            const sType = sData.chartType;
            try {
              if (sType === "line" || sType === 4) sObj.ChartType = 4;
              else if (sType === "line_markers" || sType === 65) sObj.ChartType = 65;
              else if (typeof sType === "number") sObj.ChartType = sType;
            } catch (e) { throw e; }
          }
          if (sData.axisGroup === 2 || sData.secondaryAxis) {
            try { sObj.AxisGroup = 2; } catch (e) { throw e; }
          }
          if (sData.color) {
            try {
              sObj.Format.Fill.Solid();
              sObj.Format.Fill.ForeColor.RGB = hexToPptColor(sData.color);
            } catch (e) { throw e; }
          }
          if (sData.hasDataLabels || showDataLabels) {
            try { sObj.HasDataLabels = true; } catch (e) { throw e; }
          }
        }
      } else if (categories && categories.length > 0) {
        try {
          const sc = chart.SeriesCollection();
          if (sc.Count > 0) {
            sc.Item(1).XValues = categories;
          }
        } catch (e) { throw e; }
      }
    } catch (e) {
      throw new Error("图表已创建，但数据配置未完成：" + e.message);
    }

    return chartShape;
  }

  function pptReadPresentation(app, params) {
    const { presentationName, includeNotes, maxSlides } = params || {};
    const pres = getPptPresentation(app, presentationName);
    const pptApp = getPptApp() || app;
    const slides = [];
    const count = pres.Slides ? pres.Slides.Count : 0;
    const maxS = Number(maxSlides) || 50;

    for (let i = 1; i <= Math.min(count, maxS); i++) {
      try {
        const slide = pres.Slides.Item(i);
        let title = `Slide ${i}`;
        const snippets = [];
        let notes = "";

        const shapeCount = slide.Shapes ? slide.Shapes.Count : 0;
        for (let s = 1; s <= shapeCount; s++) {
          const shp = slide.Shapes.Item(s);
          if (shp.HasTextFrame && shp.TextFrame.HasText) {
            const text = (shp.TextFrame.TextRange.Text || "").trim().replace(/[\r\n\x07]/g, " ");
            if (shp.Type === 14 || s === 1) {
              if (text && title === `Slide ${i}`) title = text;
            }
            if (text) snippets.push(text);
          }
        }

        if (includeNotes !== false) {
          try {
            if (slide.NotesPage && slide.NotesPage.Shapes) {
              const notesShape = slide.NotesPage.Shapes.Placeholders.Item(2);
              if (notesShape && notesShape.HasTextFrame && notesShape.TextFrame.HasText) {
                notes = (notesShape.TextFrame.TextRange.Text || "").trim();
              }
            }
          } catch (e) {}
        }

        slides.push({
          index: i,
          title,
          shapeCount,
          notes: notes || undefined,
          textSnippets: snippets.slice(0, 5)
        });
      } catch (e) {}
    }

    let activeIdx = 1;
    try {
      if (pptApp.ActiveWindow && pptApp.ActiveWindow.Selection && pptApp.ActiveWindow.Selection.SlideRange) {
        activeIdx = pptApp.ActiveWindow.Selection.SlideRange.SlideIndex;
      }
    } catch (e) {}

    return {
      presentationName: pres.Name,
      fullName: pres.FullName || pres.Name,
      ...pptPageSize(pres),
      slideCount: count,
      activeSlideIndex: activeIdx,
      slides
    };
  }

  function pptGenerateDeck(app, params) {
    const { presentationName, themeColor, themePreset, slides } = params || {};
    if (!Array.isArray(slides) || slides.length === 0) {
      throw new Error("缺少必要参数: slides 数组");
    }
    const pres = getPptPresentation(app, presentationName);
    const baseColor = themeColor || (themePreset === "tech_purple" ? "#4B38B3" : "#0F4C81");
    const createdIndices = [];
    const page = pptPageSize(pres);
    const warnings = [];
    // Design coordinates are mapped to the actual page before returning.
    const pageWidth = 720;
    const pageHeight = 405;

    for (let i = 0; i < slides.length; i++) {
      const spec = slides[i];
      const slide = pres.Slides.Add(pres.Slides.Count + 1, 12);
      const slideIdx = slide.SlideIndex;
      createdIndices.push(slideIdx);

      if (spec.layout === "title") {
        try {
          const bg = slide.Shapes.AddShape(1, 0, 0, pageWidth, pageHeight);
          bg.Fill.Solid();
          bg.Fill.ForeColor.RGB = hexToPptColor(baseColor);
          bg.Line.Visible = false;
        } catch (e) {}

        const titleBox = slide.Shapes.AddTextbox(1, 60, 140, pageWidth - 120, 80);
        titleBox.TextFrame.TextRange.Text = spec.title || "演示文稿";
        titleBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
        titleBox.TextFrame.TextRange.Font.Size = 36;
        titleBox.TextFrame.TextRange.Font.Bold = true;
        titleBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#FFFFFF");

        if (spec.subtitle) {
          const subBox = slide.Shapes.AddTextbox(1, 60, 230, pageWidth - 120, 40);
          subBox.TextFrame.TextRange.Text = spec.subtitle;
          subBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
          subBox.TextFrame.TextRange.Font.Size = 18;
          subBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#E2E8F0");
        }
      } else if (spec.layout === "cards_2" || spec.layout === "cards_3" || spec.layout === "cards_4" || spec.cards) {
        addSlideHeader(slide, spec.title, baseColor);
        const colCount = spec.layout === "cards_2" ? 2 : (spec.layout === "cards_4" ? 4 : (spec.cards ? Math.min(spec.cards.length, 4) : 3));
        renderPptCards(slide, spec.cards || [], colCount, baseColor);
      } else if (spec.layout === "chart" && spec.chart) {
        addSlideHeader(slide, spec.title, baseColor);
        renderPptChart(slide, spec.chart, 60, 90, pageWidth - 120, pageHeight - 120);
      } else if (spec.layout === "end") {
        try {
          const bg = slide.Shapes.AddShape(1, 0, 0, pageWidth, pageHeight);
          bg.Fill.Solid();
          bg.Fill.ForeColor.RGB = hexToPptColor(baseColor);
          bg.Line.Visible = false;
        } catch (e) {}

        const titleBox = slide.Shapes.AddTextbox(1, 60, 160, pageWidth - 120, 80);
        titleBox.TextFrame.TextRange.Text = spec.title || "THANK YOU";
        titleBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
        titleBox.TextFrame.TextRange.Font.Size = 40;
        titleBox.TextFrame.TextRange.Font.Bold = true;
        titleBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#FFFFFF");
        titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2;
      } else {
        addSlideHeader(slide, spec.title, baseColor);
        if (Array.isArray(spec.bulletPoints) && spec.bulletPoints.length > 0) {
          const contentBox = slide.Shapes.AddTextbox(1, 60, 100, pageWidth - 120, pageHeight - 140);
          contentBox.TextFrame.TextRange.Text = spec.bulletPoints.join("\n");
          contentBox.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
          contentBox.TextFrame.TextRange.Font.Size = 16;
          contentBox.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#334155");
        } else {
          // content 布局缺正文时，原实现静默只出标题、不报任何问题（问题台账 ISS-54）。
          // 这里显式告警，避免调用方以为这一页已经做完了。
          warnings.push({
            slideIndex: slideIdx,
            reason: "content 布局未提供 bulletPoints：本页只有标题、没有正文。请补 bulletPoints，或改用 cards/chart 布局"
          });
        }
      }

      fitGeneratedPptShapes(slide, 1, page, warnings);

      if (spec.notes) {
        try {
          if (slide.NotesPage && slide.NotesPage.Shapes) {
            const notesShape = slide.NotesPage.Shapes.Placeholders.Item(2);
            if (notesShape && notesShape.HasTextFrame) {
              notesShape.TextFrame.TextRange.Text = spec.notes;
            }
          }
        } catch (e) {}
      }
    }

    return {
      success: true,
      ...page,
      layoutWarnings: warnings,
      visualVerificationRequired: true,
      presentationName: pres.Name,
      createdSlidesCount: slides.length,
      createdSlideIndices: createdIndices,
      themeColor: baseColor,
      message: `已成功基于大纲批量生成 ${slides.length} 页专业商业演示胶片！`
    };
  }

  function pptManageSlides(app, params) {
    const { presentationName, action, slideIndex, targetIndex, layoutIndex, backgroundColor } = params || {};
    const pres = getPptPresentation(app, presentationName);

    switch (action) {
      case "add": {
        const idx = slideIndex ? Number(slideIndex) : (pres.Slides.Count + 1);
        const lIndex = Number(layoutIndex) || 12;
        const newSlide = pres.Slides.Add(idx, lIndex);
        return { success: true, presentationName: pres.Name, slideIndex: newSlide.SlideIndex, message: `已在位置 ${newSlide.SlideIndex} 新增幻灯片` };
      }
      case "delete": {
        if (!slideIndex) throw new Error("delete 操作必须提供 slideIndex");
        const idx = Number(slideIndex);
        const slide = pres.Slides.Item(idx);
        slide.Delete();
        return { success: true, presentationName: pres.Name, deletedIndex: idx, message: `已成功删除第 ${idx} 页幻灯片` };
      }
      case "move": {
        if (!slideIndex || !targetIndex) throw new Error("move 操作必须提供 slideIndex 与 targetIndex");
        const slide = pres.Slides.Item(Number(slideIndex));
        slide.MoveTo(Number(targetIndex));
        return { success: true, presentationName: pres.Name, from: slideIndex, to: targetIndex, message: `幻灯片已移动至第 ${targetIndex} 页` };
      }
      case "duplicate": {
        if (!slideIndex) throw new Error("duplicate 操作必须提供 slideIndex");
        const slide = pres.Slides.Item(Number(slideIndex));
        slide.Duplicate();
        return { success: true, presentationName: pres.Name, originalIndex: slideIndex, message: `已成功克隆第 ${slideIndex} 页幻灯片` };
      }
      case "set_background": {
        if (!slideIndex || !backgroundColor) throw new Error("set_background 操作必须提供 slideIndex 与 backgroundColor");
        const slide = pres.Slides.Item(Number(slideIndex));
        const total = pres.Slides.Count;

        // 读回背景色的辅助：不同宿主返回的 RGB 可能是 number 也可能是其它形态，统一成 "R,G,B"
        const readRgb = (target) => {
          try {
            const rgb = Number(target.Background.Fill.ForeColor.RGB);
            if (!Number.isFinite(rgb)) return null;
            const b = rgb & 0xff, g = (rgb >> 8) & 0xff, r = (rgb >> 16) & 0xff;
            return r + "," + g + "," + b;
          } catch (e) { return null; }
        };
        const neighbourIndex = Number(slideIndex) === 1 ? Math.min(2, total) : Number(slideIndex) - 1;
        const neighbourBefore = neighbourIndex !== Number(slideIndex) ? readRgb(pres.Slides.Item(neighbourIndex)) : null;

        // 关键：该页若仍"跟随母版背景"，`slide.Background` 可能指向母版对象，写入会**串改全部页**
        // （问题台账 ISS-87，受控复现：设第 1 页后第 2 页也变红）。先显式断开与母版的关联再写。
        try { slide.FollowMasterBackground = false; } catch (e) {}
        slide.Background.Fill.Solid();
        slide.Background.Fill.ForeColor.RGB = hexToPptColor(backgroundColor);

        // 写后校验：目标页确实变了，且相邻页**没有被串改**
        const applied = readRgb(slide);
        const neighbourAfter = neighbourIndex !== Number(slideIndex) ? readRgb(pres.Slides.Item(neighbourIndex)) : null;
        if (neighbourBefore !== null && neighbourAfter !== null && neighbourBefore !== neighbourAfter) {
          throw new Error(
            `设置背景时串改了相邻页：第 ${neighbourIndex} 页背景由 ${neighbourBefore} 变成了 ${neighbourAfter}。` +
            `已尝试先断开 FollowMasterBackground；请检查该稿的母版/版式是否被直接修改。`
          );
        }
        return {
          success: true,
          presentationName: pres.Name,
          slideIndex,
          backgroundColor,
          appliedRgb: applied,
          neighbourChecked: neighbourIndex !== Number(slideIndex) ? { slideIndex: neighbourIndex, before: neighbourBefore, after: neighbourAfter } : null,
          message: `已将第 ${slideIndex} 页背景设为 ${backgroundColor}`
        };
      }
      default:
        throw new Error(`未知的 PPT 页面操作: ${action}`);
    }
  }

  function pptAddBusinessCards(app, params) {
    const { presentationName, slideIndex, columnCount, cards, topY, cardHeight } = params || {};
    if (!Array.isArray(cards) || cards.length === 0) throw new Error("缺少 cards 数组");
    const pres = getPptPresentation(app, presentationName);
    const idx = Number(slideIndex) || (pres.Slides.Count > 0 ? 1 : 1);
    const slide = pres.Slides.Item(idx);
    const cols = columnCount || (cards.length === 2 ? 2 : (cards.length === 4 ? 4 : 3));

    const page = pptPageSize(pres);
    const warnings = [];
    const firstIndex = slide.Shapes.Count + 1;
    renderPptCards(slide, cards, cols, "#0F4C81",
      topY === undefined ? undefined : Number(topY) * 405 / page.pageHeight,
      cardHeight === undefined ? undefined : Number(cardHeight) * 405 / page.pageHeight);
    fitGeneratedPptShapes(slide, firstIndex, page, warnings);
    return {
      success: true,
      presentationName: pres.Name,
      slideIndex: idx,
      ...page,
      layoutWarnings: warnings,
      visualVerificationRequired: true,
      columns: cols,
      cardsCount: cards.length,
      message: `已在第 ${idx} 页成功排版 ${cards.length} 张现代化商业信息卡片`
    };
  }

  function pptInsertNativeChart(app, params) {
    const { presentationName, slideIndex, chartType, title, categories, series, left, top, width, height } = params || {};
    const pres = getPptPresentation(app, presentationName);
    const idx = Number(slideIndex) || (pres.Slides.Count > 0 ? 1 : 1);
    const slide = pres.Slides.Item(idx);

    const page = pptPageSize(pres);
    const l = left !== undefined ? Number(left) : page.pageWidth / 12;
    const t = top !== undefined ? Number(top) : page.pageHeight * 90 / 405;
    const w = width !== undefined ? Number(width) : page.pageWidth * 600 / 720;
    const h = height !== undefined ? Number(height) : page.pageHeight * 280 / 405;

    renderPptChart(slide, { chartType: chartType || "column", title, categories, series }, l, t, w, h);
    return {
      success: true,
      presentationName: pres.Name,
      slideIndex: idx,
      chartType: chartType || "column",
      message: `已在第 ${idx} 页成功插入原生矢量图表`
    };
  }

  function getShapeTypeName(typeCode) {
    switch (typeCode) {
      case 1: return "shape"; // msoAutoShape
      case 3: return "chart"; // msoChart
      case 6: return "group"; // msoGroup
      case 13: return "picture"; // msoPicture
      case 14: return "placeholder"; // msoPlaceholder
      case 17: return "textbox"; // msoTextBox
      case 19: return "table"; // msoTable
      case 24: return "smartArt"; // msoSmartArt
      default: return `type_${typeCode}`;
    }
  }

  function pptGetSlideShapes(app, params) {
    const { presentationName, slideIndex } = params || {};
    const pres = getPptPresentation(app, presentationName);
    const pptApp = getPptApp() || app;
    let idx = Number(slideIndex);
    if (!idx || isNaN(idx)) {
      try {
        if (pptApp.ActiveWindow && pptApp.ActiveWindow.Selection && pptApp.ActiveWindow.Selection.SlideRange) {
          idx = pptApp.ActiveWindow.Selection.SlideRange.SlideIndex;
        }
      } catch (e) {}
      if (!idx) idx = 1;
    }

    const slide = pres.Slides.Item(idx);
    const count = slide.Shapes ? slide.Shapes.Count : 0;
    const shapes = [];

    for (let s = 1; s <= count; s++) {
      try {
        const shp = slide.Shapes.Item(s);
        const hasText = Boolean(shp.HasTextFrame && shp.TextFrame.HasText);
        let textContent = "";
        if (hasText) {
          textContent = (shp.TextFrame.TextRange.Text || "").trim();
        }

        const hasTable = Boolean(shp.HasTable);
        let tableMeta = null;
        if (hasTable && shp.Table) {
          tableMeta = {
            rows: shp.Table.Rows ? shp.Table.Rows.Count : 0,
            columns: shp.Table.Columns ? shp.Table.Columns.Count : 0
          };
        }

        const hasChart = Boolean(shp.HasChart);

        shapes.push({
          shapeIndex: s,
          shapeId: shp.Id,
          name: shp.Name || `Shape_${s}`,
          typeCode: shp.Type,
          typeName: getShapeTypeName(shp.Type),
          left: Math.round(Number(shp.Left) * 10) / 10,
          top: Math.round(Number(shp.Top) * 10) / 10,
          width: Math.round(Number(shp.Width) * 10) / 10,
          height: Math.round(Number(shp.Height) * 10) / 10,
          rotation: shp.Rotation || 0,
          zOrderPosition: shp.ZOrderPosition || s,
          hasText,
          fontSize: hasText ? Number(shp.TextFrame.TextRange.Font.Size) : undefined,
          text: textContent ? (textContent.length > 200 ? textContent.slice(0, 200) + "..." : textContent) : undefined,
          hasTable,
          table: tableMeta || undefined,
          hasChart
        });
      } catch (e) {
        log(`读取第 ${idx} 页第 ${s} 个形状元数据失败: ${e.message}`);
      }
    }

    return {
      success: true,
      presentationName: pres.Name,
      slideIndex: idx,
      ...pptPageSize(pres),
      shapeCount: count,
      shapes,
      message: `已成功获取第 ${idx} 页幻灯片中全部 ${shapes.length} 个形状的几何与属性信息`
    };
  }

  function findShapeOnSlide(slide, shapeIdOrIndex) {
    if (shapeIdOrIndex === undefined || shapeIdOrIndex === null) return null;
    const count = slide.Shapes.Count;
    // 1. 如果传入数字且在范围内，先尝试直接按索引或 ID 获取
    if (typeof shapeIdOrIndex === "number" || /^\d+$/.test(String(shapeIdOrIndex))) {
      const num = Number(shapeIdOrIndex);
      // 先遍历按 Id 匹配
      for (let s = 1; s <= count; s++) {
        const item = slide.Shapes.Item(s);
        if (item.Id === num) return item;
      }
      // 再按索引匹配
      if (num >= 1 && num <= count) {
        return slide.Shapes.Item(num);
      }
    }
    // 2. 按名称查找
    for (let s = 1; s <= count; s++) {
      const item = slide.Shapes.Item(s);
      if (item.Name === String(shapeIdOrIndex)) return item;
    }
    // 3. 兜底直接 Item
    try {
      return slide.Shapes.Item(shapeIdOrIndex);
    } catch (e) {
      return null;
    }
  }

  function pptManageShapesAndMedia(app, params) {
    const {
      presentationName,
      slideIndex,
      action,
      shapeId,
      shapeId1,
      shapeId2,
      shapeType,
      text,
      left,
      top,
      width,
      height,
      fontSize,
      fontColor,
      fontBold,
      alignment,
      fillColor,
      lineColor,
      rotation,
      zOrderAction,
      alignType
    } = params || {};

    const pres = getPptPresentation(app, presentationName);
    const pptApp = getPptApp() || app;
    let idx = Number(slideIndex);
    if (!idx || isNaN(idx)) {
      try {
        if (pptApp.ActiveWindow && pptApp.ActiveWindow.Selection && pptApp.ActiveWindow.Selection.SlideRange) {
          idx = pptApp.ActiveWindow.Selection.SlideRange.SlideIndex;
        }
      } catch (e) {}
      if (!idx) idx = 1;
    }
    const slide = pres.Slides.Item(idx);

    const l = left !== undefined ? Number(left) : 100;
    const t = top !== undefined ? Number(top) : 100;
    const w = width !== undefined ? Number(width) : 200;
    const h = height !== undefined ? Number(height) : 100;

    switch (action) {
      case "add_textbox": {
        const tb = slide.Shapes.AddTextbox(1, l, t, w, h);
        const tr = tb.TextFrame.TextRange;
        tr.Text = text || "新文本框";
        tr.Font.Name = "Microsoft YaHei";
        if (fontSize) tr.Font.Size = Number(fontSize);
        if (fontColor) tr.Font.Color.RGB = hexToPptColor(fontColor);
        if (fontBold !== undefined) tr.Font.Bold = Boolean(fontBold);
        if (alignment !== undefined) {
          const alignMap = { left: 1, center: 2, right: 3, justify: 4 };
          tr.ParagraphFormat.Alignment = alignMap[alignment] || 1;
        }
        return { success: true, shapeId: tb.Id, left: tb.Left, top: tb.Top, width: tb.Width, height: tb.Height, fontSize: Number(tr.Font.Size), unit: "pt", message: "已成功添加文本框" };
      }

      case "add_shape": {
        let typeCode = 1;
        if (shapeType === "rounded_rectangle") typeCode = 5;
        else if (shapeType === "oval") typeCode = 9;
        else if (shapeType === "arrow") typeCode = 13;

        const shp = slide.Shapes.AddShape(typeCode, l, t, w, h);
        if (text) {
          shp.TextFrame.TextRange.Text = text;
          shp.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
          if (fontSize) shp.TextFrame.TextRange.Font.Size = Number(fontSize);
        }
        if (fillColor) {
          shp.Fill.Solid();
          shp.Fill.ForeColor.RGB = hexToPptColor(fillColor);
        }
        if (lineColor) {
          shp.Line.ForeColor.RGB = hexToPptColor(lineColor);
        }
        return { success: true, shapeId: shp.Id, left: shp.Left, top: shp.Top, width: shp.Width, height: shp.Height, unit: "pt", message: "已成功添加形状" };
      }

      case "update_shape": {
        const targetId = shapeId;
        if (targetId === undefined || targetId === null) throw new Error("update_shape 操作必须提供 shapeId");
        const shp = findShapeOnSlide(slide, targetId);
        if (!shp) throw new Error(`未在第 ${idx} 页找到形状: ${targetId}`);

        if (left !== undefined) shp.Left = Number(left);
        if (top !== undefined) shp.Top = Number(top);
        if (width !== undefined) shp.Width = Number(width);
        if (height !== undefined) shp.Height = Number(height);
        if (rotation !== undefined) shp.Rotation = Number(rotation);

        if (text !== undefined && shp.HasTextFrame) {
          shp.TextFrame.TextRange.Text = text;
        }
        if (shp.HasTextFrame && shp.TextFrame.HasText) {
          const tr = shp.TextFrame.TextRange;
          if (fontSize !== undefined) tr.Font.Size = Number(fontSize);
          if (fontColor !== undefined) tr.Font.Color.RGB = hexToPptColor(fontColor);
          if (fontBold !== undefined) tr.Font.Bold = Boolean(fontBold);
          if (alignment !== undefined) {
            const alignMap = { left: 1, center: 2, right: 3, justify: 4 };
            tr.ParagraphFormat.Alignment = alignMap[alignment] || 1;
          }
        }
        if (fillColor !== undefined) {
          shp.Fill.Solid();
          shp.Fill.ForeColor.RGB = hexToPptColor(fillColor);
        }
        if (lineColor !== undefined) {
          shp.Line.ForeColor.RGB = hexToPptColor(lineColor);
        }

        return {
          success: true,
          shapeId: shp.Id,
          left: shp.Left,
          top: shp.Top,
          width: shp.Width,
          height: shp.Height,
          message: `已成功更新形状 [${shp.Id}] 的属性`
        };
      }

      case "swap_shapes": {
        const id1 = shapeId1 !== undefined ? shapeId1 : shapeId;
        const id2 = shapeId2;
        if (!id1 || !id2) throw new Error("swap_shapes 必须提供 shapeId1 和 shapeId2 两个目标形状标识");

        const shp1 = findShapeOnSlide(slide, id1);
        const shp2 = findShapeOnSlide(slide, id2);
        if (!shp1) throw new Error(`未找到第一个形状: ${id1}`);
        if (!shp2) throw new Error(`未找到第二个形状: ${id2}`);

        const top1 = Number(shp1.Top);
        const height1 = Number(shp1.Height);
        const left1 = Number(shp1.Left);

        const top2 = Number(shp2.Top);
        const height2 = Number(shp2.Height);
        const left2 = Number(shp2.Left);

        // 智能垂直互换（保持上下文视觉流）
        if (top1 < top2) {
          // shp1 在上方，shp2 在下方
          const gap = top2 - (top1 + height1);
          const effectiveGap = gap > 0 ? gap : 15;
          // 将 shp2 移到上方原 shp1 的 Top
          shp2.Top = top1;
          // 将 shp1 移到 shp2 下方
          shp1.Top = top1 + height2 + effectiveGap;
        } else {
          // shp2 在上方，shp1 在下方
          const gap = top1 - (top2 + height2);
          const effectiveGap = gap > 0 ? gap : 15;
          shp1.Top = top2;
          shp2.Top = top2 + height1 + effectiveGap;
        }

        return {
          success: true,
          slideIndex: idx,
          shape1: { id: shp1.Id, oldTop: top1, newTop: shp1.Top, height: height1 },
          shape2: { id: shp2.Id, oldTop: top2, newTop: shp2.Top, height: height2 },
          message: `已成功将形状 [${shp1.Id}] 与 [${shp2.Id}] 在第 ${idx} 页进行精准垂直互换！`
        };
      }

      case "set_z_order": {
        if (!shapeId) throw new Error("set_z_order 必须提供 shapeId");
        const shp = findShapeOnSlide(slide, shapeId);
        if (!shp) throw new Error(`未找到形状: ${shapeId}`);

        // 0: msoBringToFront, 1: msoSendToBack, 2: msoBringForward, 3: msoSendBackward
        const zMap = {
          bring_to_front: 0,
          send_to_back: 1,
          bring_forward: 2,
          send_backward: 3
        };
        const cmd = zMap[zOrderAction || "bring_to_front"];
        if (cmd !== undefined) {
          shp.ZOrder(cmd);
        }
        return { success: true, shapeId: shp.Id, zOrderAction, message: `已成功调整形状 [${shp.Id}] 的图层层级` };
      }

      case "align_shapes": {
        const ids = params.shapeIds || (shapeId1 && shapeId2 ? [shapeId1, shapeId2] : []);
        if (!Array.isArray(ids) || ids.length === 0) throw new Error("align_shapes 必须提供 shapeIds 数组");
        const alignMode = alignType || "center"; // left | center | right | top | middle | bottom

        let refVal = null;
        for (let i = 0; i < ids.length; i++) {
          const shp = findShapeOnSlide(slide, ids[i]);
          if (!shp) continue;
          if (i === 0) {
            if (alignMode === "left") refVal = shp.Left;
            else if (alignMode === "center") refVal = shp.Left + shp.Width / 2;
            else if (alignMode === "right") refVal = shp.Left + shp.Width;
            else if (alignMode === "top") refVal = shp.Top;
            else if (alignMode === "middle") refVal = shp.Top + shp.Height / 2;
            else if (alignMode === "bottom") refVal = shp.Top + shp.Height;
          } else {
            if (alignMode === "left") shp.Left = refVal;
            else if (alignMode === "center") shp.Left = refVal - shp.Width / 2;
            else if (alignMode === "right") shp.Left = refVal - shp.Width;
            else if (alignMode === "top") shp.Top = refVal;
            else if (alignMode === "middle") shp.Top = refVal - shp.Height / 2;
            else if (alignMode === "bottom") shp.Top = refVal - shp.Height;
          }
        }
        return { success: true, alignMode, alignedCount: ids.length, message: `已完成 ${ids.length} 个形状的 [${alignMode}] 对齐` };
      }

      case "delete_shape": {
        if (!shapeId) throw new Error("delete_shape 操作必须提供 shapeId");
        const shp = findShapeOnSlide(slide, shapeId);
        if (!shp) throw new Error(`未找到形状: ${shapeId}`);
        const delId = shp.Id;
        shp.Delete();
        return { success: true, message: `已成功删除形状: ${delId}` };
      }

      default:
        throw new Error(`未知的形状/多媒体操作: ${action}`);
    }
  }

  function pptManageTable(app, params) {
    const {
      presentationName,
      slideIndex,
      action,
      shapeId,
      tableIndex,
      rows,
      columns,
      columnWidths,
      rowHeights,
      left,
      top,
      width,
      height,
      data,
      row,
      column,
      text,
      fontSize,
      fontColor,
      fontBold,
      fillColor,
      headerFillColor,
      headerFontSize,
      bodyFontSize,
      borderColor,
      zebra
    } = params || {};

    const pres = getPptPresentation(app, presentationName);
    const pptApp = getPptApp() || app;
    let idx = Number(slideIndex);
    if (!idx || isNaN(idx)) {
      try {
        if (pptApp.ActiveWindow && pptApp.ActiveWindow.Selection && pptApp.ActiveWindow.Selection.SlideRange) {
          idx = pptApp.ActiveWindow.Selection.SlideRange.SlideIndex;
        }
      } catch (e) {}
      if (!idx) idx = 1;
    }
    const slide = pres.Slides.Item(idx);

    function findTableShape() {
      if (shapeId) {
        const shp = findShapeOnSlide(slide, shapeId);
        if (shp && shp.HasTable && shp.Table) return shp;
      }
      let tCount = 0;
      const targetTableIdx = Number(tableIndex) || 1;
      for (let s = 1; s <= slide.Shapes.Count; s++) {
        const shp = slide.Shapes.Item(s);
        if (shp.HasTable && shp.Table) {
          tCount++;
          if (tCount === targetTableIdx) return shp;
        }
      }
      return null;
    }

    function applyCellBorders(tbl, bColorHex) {
      const bColor = hexToPptColor(bColorHex || "#333333");
      const whiteColor = hexToPptColor("#FFFFFF");
      const rowCount = tbl.Rows.Count;
      const colCount = tbl.Columns.Count;

      for (let r = 1; r <= rowCount; r++) {
        for (let c = 1; c <= colCount; c++) {
          try {
            const cell = tbl.Cell(r, c);
            for (let b = 1; b <= 4; b++) {
              const border = cell.Borders.Item(b);
              border.Visible = true;
              border.Weight = 1;

              if (r === 1) {
                // 表头行：上下外框与最左/最右外侧使用黑色闭合线，内部纵向使用高保真白色分割线
                if (b === 1 || b === 3) {
                  border.ForeColor.RGB = bColor;
                  border.Weight = 1;
                } else if (b === 2) {
                  if (c === 1) {
                    border.ForeColor.RGB = bColor;
                    border.Weight = 1;
                  } else {
                    border.ForeColor.RGB = whiteColor;
                    border.Weight = 1.5;
                  }
                } else if (b === 4) {
                  if (c === colCount) {
                    border.ForeColor.RGB = bColor;
                    border.Weight = 1;
                  } else {
                    border.ForeColor.RGB = whiteColor;
                    border.Weight = 1.5;
                  }
                }
              } else {
                border.ForeColor.RGB = bColor;
                border.Weight = 1;
              }
            }
          } catch (e) {}
        }
      }
    }

    switch (action) {
      case "create_table": {
        const rCount = Number(rows) || 3;
        const cCount = Number(columns) || 3;
        const l = left !== undefined ? Number(left) : 32.6;
        const t = top !== undefined ? Number(top) : 58;
        const w = width !== undefined ? Number(width) : 895;
        const h = height !== undefined ? Number(height) : 435;

        const tableShape = slide.Shapes.AddTable(rCount, cCount, l, t, w, h);
        const tbl = tableShape.Table;

        try {
          tbl.ApplyStyle("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", false);
        } catch (e) {}
        try {
          tbl.FirstRow = false;
          tbl.BandedRows = false;
          tbl.BandedColumns = false;
        } catch (e) {}

        // 设置列宽
        if (Array.isArray(columnWidths) && columnWidths.length > 0) {
          for (let c = 0; c < Math.min(columnWidths.length, tbl.Columns.Count); c++) {
            try {
              tbl.Columns.Item(c + 1).Width = Number(columnWidths[c]);
            } catch (e) {}
          }
        }

        // 填充数据与富文本分段加粗
        if (Array.isArray(data) && data.length > 0) {
          for (let r = 0; r < Math.min(data.length, tbl.Rows.Count); r++) {
            const rowData = data[r];
            if (Array.isArray(rowData)) {
              for (let c = 0; c < Math.min(rowData.length, tbl.Columns.Count); c++) {
                const cell = tbl.Cell(r + 1, c + 1);
                const cellText = String(rowData[c] || "");
                const tr = cell.Shape.TextFrame.TextRange;
                tr.Text = cellText;
                tr.Font.Name = "Microsoft YaHei";

                const hFontSize = Number(headerFontSize) || 16;
                const bFontSize = Number(bodyFontSize) || 14;

                // 垂直居中与内边距规范
                try {
                  cell.Shape.TextFrame.VerticalAnchor = 3; // msoAnchorMiddle 垂直居中
                  cell.Shape.TextFrame.MarginLeft = 5;
                  cell.Shape.TextFrame.MarginRight = 5;
                  cell.Shape.TextFrame.MarginTop = 3;
                  cell.Shape.TextFrame.MarginBottom = 3;
                  cell.Shape.TextFrame.WordWrap = true;
                } catch (e) {}

                if (r === 0) {
                  // 表头标题：16pt 白色加粗居中
                  cell.Shape.Fill.Solid();
                  cell.Shape.Fill.ForeColor.RGB = hexToPptColor(headerFillColor || "#0072C6");
                  tr.Font.Size = hFontSize;
                  tr.Font.Bold = true;
                  tr.Font.Color.RGB = hexToPptColor("#FFFFFF");
                  tr.ParagraphFormat.Alignment = 2; // 居中
                } else {
                  // 数据行底色纯白
                  cell.Shape.Fill.Solid();
                  cell.Shape.Fill.ForeColor.RGB = 16777215; // 0xFFFFFF 纯白
                  tr.Font.Color.RGB = hexToPptColor("#000000");

                  if (c === 0) {
                    // 第一列：指标名称 14pt 居中加粗
                    tr.Font.Size = bFontSize;
                    tr.Font.Bold = true;
                    tr.ParagraphFormat.Alignment = 2;
                  } else if (c === 3) {
                    // 第四列：判定状态列 15.5pt 加粗居中
                    tr.Font.Size = bFontSize + 1.5;
                    tr.Font.Bold = true;
                    tr.ParagraphFormat.Alignment = 2;
                    if (cellText === "达成") {
                      tr.Font.Color.RGB = hexToPptColor("#009132");
                    } else if (cellText === "进行中") {
                      tr.Font.Color.RGB = hexToPptColor("#E36C09");
                    } else if (cellText === "未开展") {
                      tr.Font.Color.RGB = hexToPptColor("#FF202E");
                    }
                  } else {
                    // 内容列：统一 14pt，首句分段精准加粗，破折号居中
                    tr.Font.Size = bFontSize;
                    if (cellText === "——") {
                      tr.ParagraphFormat.Alignment = 2;
                    } else {
                      tr.ParagraphFormat.Alignment = 1;
                    }

                    const sepIdx = cellText.search(/[；;:：]/);
                    if (sepIdx > 0 && sepIdx < 35) {
                      try {
                        const part1 = tr.Characters(1, sepIdx + 1);
                        part1.Font.Bold = true;
                        part1.Font.Size = bFontSize;
                        const part2 = tr.Characters(sepIdx + 2);
                        part2.Font.Bold = false;
                        part2.Font.Size = bFontSize;
                      } catch (e) {}
                    } else if (cellText.startsWith("最高认证效率")) {
                      try {
                        tr.Font.Bold = true;
                        tr.Font.Size = bFontSize;
                      } catch (e) {}
                    } else {
                      tr.Font.Bold = false;
                      tr.Font.Size = bFontSize;
                    }
                  }
                }
              }
            }
          }
        }

        // 动态阶梯行高矩阵优化
        if (Array.isArray(rowHeights) && rowHeights.length > 0) {
          for (let r = 0; r < Math.min(rowHeights.length, tbl.Rows.Count); r++) {
            try {
              tbl.Rows.Item(r + 1).Height = Number(rowHeights[r]);
            } catch (e) {}
          }
        } else {
          try {
            tbl.Rows.Item(1).Height = 30;
            for (let r = 2; r <= tbl.Rows.Count; r++) {
              tbl.Rows.Item(r).Height = 46;
            }
          } catch (e) {}
        }

        // 应用网格实线边框
        applyCellBorders(tbl, borderColor || "#333333");

        return {
          success: true,
          shapeId: tableShape.Id,
          slideIndex: idx,
          rows: rCount,
          columns: cCount,
          message: `已在第 ${idx} 页成功创建 8 行 5 列表格（含细实线网格、精准列宽与分段加粗）`
        };
      }

      case "read_table": {
        const shp = findTableShape();
        if (!shp) throw new Error(`未在第 ${idx} 页找到表格形状`);
        const tbl = shp.Table;
        const tableData = [];
        const rCount = tbl.Rows.Count;
        const cCount = tbl.Columns.Count;

        for (let r = 1; r <= rCount; r++) {
          const rowArr = [];
          for (let c = 1; c <= cCount; c++) {
            try {
              const cell = tbl.Cell(r, c);
              const txt = cell.Shape && cell.Shape.HasTextFrame && cell.Shape.TextFrame.HasText
                ? cell.Shape.TextFrame.TextRange.Text.trim()
                : "";
              rowArr.push(txt);
            } catch (e) {
              rowArr.push("");
            }
          }
          tableData.push(rowArr);
        }

        return {
          success: true,
          shapeId: shp.Id,
          slideIndex: idx,
          rows: rCount,
          columns: cCount,
          data: tableData,
          message: `已成功读取第 ${idx} 页表格数据 (${rCount} 行 ${cCount} 列)`
        };
      }

      case "set_cell_text": {
        const shp = findTableShape();
        if (!shp) throw new Error(`未在第 ${idx} 页找到表格形状`);
        const tbl = shp.Table;
        const rIdx = Number(row) || 1;
        const cIdx = Number(column) || 1;
        const cell = tbl.Cell(rIdx, cIdx);

        if (text !== undefined) {
          cell.Shape.TextFrame.TextRange.Text = String(text);
          cell.Shape.TextFrame.TextRange.Font.Name = "Microsoft YaHei";
        }
        if (fontSize) cell.Shape.TextFrame.TextRange.Font.Size = Number(fontSize);
        if (fontColor) cell.Shape.TextFrame.TextRange.Font.Color.RGB = hexToPptColor(fontColor);
        if (fontBold !== undefined) cell.Shape.TextFrame.TextRange.Font.Bold = Boolean(fontBold);
        if (fillColor) {
          cell.Shape.Fill.Solid();
          cell.Shape.Fill.ForeColor.RGB = hexToPptColor(fillColor);
        }

        return {
          success: true,
          shapeId: shp.Id,
          row: rIdx,
          column: cIdx,
          message: `已成功更新单元格 [${rIdx}, ${cIdx}] 的内容与格式`
        };
      }

      case "style_table": {
        const shp = findTableShape();
        if (!shp) throw new Error(`未在第 ${idx} 页找到表格形状`);
        const tbl = shp.Table;
        const hColor = headerFillColor || "#0072C6";

        for (let r = 1; r <= tbl.Rows.Count; r++) {
          for (let c = 1; c <= tbl.Columns.Count; c++) {
            const cell = tbl.Cell(r, c);
            if (r === 1) {
              cell.Shape.Fill.Solid();
              cell.Shape.Fill.ForeColor.RGB = hexToPptColor(hColor);
              if (cell.Shape.HasTextFrame && cell.Shape.TextFrame.HasText) {
                cell.Shape.TextFrame.TextRange.Font.Bold = true;
                cell.Shape.TextFrame.TextRange.Font.Color.RGB = hexToPptColor("#FFFFFF");
              }
            } else {
              cell.Shape.Fill.Solid();
              cell.Shape.Fill.ForeColor.RGB = hexToPptColor("#FFFFFF");
            }
          }
        }

        applyCellBorders(tbl, borderColor || "#333333");

        return {
          success: true,
          shapeId: shp.Id,
          headerColor: hColor,
          message: `已成功应用专业商业表格配色主题与边框网格`
        };
      }

      default:
        throw new Error(`未知的表格操作: ${action}`);
    }
  }

  function pptCaptureSlidePreview(app, params) {
    const { presentationName, slideIndex } = params || {};
    const pres = getPptPresentation(app, presentationName);
    const pptApp = getPptApp() || app;
    let idx = Number(slideIndex);
    if (!idx || isNaN(idx)) {
      try {
        if (pptApp.ActiveWindow && pptApp.ActiveWindow.Selection && pptApp.ActiveWindow.Selection.SlideRange) {
          idx = pptApp.ActiveWindow.Selection.SlideRange.SlideIndex;
        }
      } catch (e) {}
      if (!idx) idx = 1;
    }
    const slide = pres.Slides.Item(idx);

    const tempPngPath = params.outputPath;
    if (!tempPngPath) throw new Error("缺少 Bridge 指定的预览输出路径");
    try {
      slide.Export(tempPngPath, "PNG", 1280, 720);
      return {
        success: true,
        presentationName: pres.Name,
        slideIndex: idx,
        imagePath: tempPngPath,
        hasImage: true,
        message: `已成功导出第 ${idx} 页幻灯片高保真预览图至: ${tempPngPath}`
      };
    } catch (e) {
      log(`幻灯片导出异常: ${e.message}`);
      return {
        success: false,
        presentationName: pres.Name,
        slideIndex: idx,
        hasImage: false,
        error: e.message
      };
    }
  }

  // ---------------------------------------------------------------------------
  // bootstrap.js — 加载项启动引导：初次连接、5 秒注册心跳、DOM 就绪后重连
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
  try {
    initWebSocket();
  } catch (e) {}

  setInterval(function () {
    if (!isConnected || !ws || ws.readyState !== WebSocket.OPEN) return;
    try {
      const host = detectHostComponent();
      sendPacket({ type: "register", client: host === "word" ? "wps-word-addon" : host === "ppt" ? "wps-ppt-addon" : "wps-et-addon", version: "2.1.0", summary: getWorkspaceSummary(getApp()) });
    } catch (error) { log("工作区状态更新失败: " + error.message); }
  }, 5000);

  if (typeof document !== "undefined") {
    if (document.readyState === "loading") {
      window.addEventListener("DOMContentLoaded", () => {
        initWebSocket();
      });
    }
    window.addEventListener("load", () => {
      initWebSocket();
    });
  }
})();
