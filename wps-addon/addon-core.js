// 本文件由 scripts/build-wps-addon.mjs 生成，请勿手改；改动请改 wps-addon/src/**
// ADDON_BUILD_FINGERPRINT: 503976199c3e9e818e7ce17f070c96011eee6deb18de78f83a501419f32c668b
(function () {
  var ADDON_BUILD_FINGERPRINT = "503976199c3e9e818e7ce17f070c96011eee6deb18de78f83a501419f32c668b";
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
        case "configure_print_layout":
          result = configurePrintLayout(app, params);
          break;
        case "export_sheet_pdf":
          result = exportSheetPdf(app, params);
          break;
        case "set_sheet_view":
          result = setSheetView(app, params);
          break;
        case "add_shape":
          result = addShape(app, params);
          break;
        case "list_shapes":
          result = listShapes(app, params);
          break;
        case "update_shape":
          result = updateShape(app, params);
          break;
        case "group_shapes":
          result = groupShapes(app, params);
          break;
        case "ungroup_shapes":
          result = ungroupShapes(app, params);
          break;
        case "set_shape_zorder":
          result = setShapeZorder(app, params);
          break;
        case "format_text_segment":
          result = formatTextSegment(app, params);
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

        case "word_update_fields":
          result = wordUpdateFields(app, params);
          break;
        case "word_manage_content_controls":
          result = wordManageContentControls(app, params);
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

          // 安全序列化返回值（防止 Office 原生对象循环引用）。
          // 原实现 `depth > 2` 时直接 `String(val)` —— 三层以上的对象**属性被静默丢弃**，
          // 调用方拿到 undefined 或 "[object Object]" 却没有任何提示（问题台账 ISS-56）。
          // 现在：放宽到 6 层，超限节点写成**显式占位**并记录路径，随结果返回 truncatedPaths。
          const MAX_SERIALIZE_DEPTH = 6;
          const truncatedPaths = [];
          function safeSerialize(val, depth = 0, path = "$") {
            if (val === null || val === undefined) return val;
            if (typeof val !== "object") return val;
            if (depth > MAX_SERIALIZE_DEPTH) {
              truncatedPaths.push(path);
              return `[第 ${depth} 层超出上限 ${MAX_SERIALIZE_DEPTH}，属性已省略；需要完整数据请自行 JSON.stringify 后返回字符串]`;
            }
            if (Array.isArray(val)) {
              return val.map((item, i) => safeSerialize(item, depth + 1, `${path}[${i}]`));
            }
            const out = {};
            for (const k in val) {
              try {
                const v = val[k];
                if (typeof v === "function") continue;
                out[k] = safeSerialize(v, depth + 1, `${path}.${k}`);
              } catch (e) {
                out[k] = `<读取属性失败: ${e.message}>`;
              }
            }
            return Object.keys(out).length > 0 ? out : String(val);
          }

          const serialized = safeSerialize(evalResult);
          result = {
            success: true,
            executionTimeMs: Date.now() - startTime,
            returnValue: serialized,
            // 被截断就明确说出来，不再静默丢数据
            truncated: truncatedPaths.length > 0,
            truncatedPaths: truncatedPaths.length ? truncatedPaths.slice(0, 10) : undefined,
            logs,
            message: `WPS 原生图灵脚本执行完毕（耗时 ${Date.now() - startTime}ms）` +
              (truncatedPaths.length ? `；返回值有 ${truncatedPaths.length} 处超出深度上限被省略，见 truncatedPaths` : "")
          };
          break;
        }

        case "inspect_api": {
          const targetExpr = params?.expression || params?.path || "app";
          // ISS-89 护栏：反射**默认只列成员名，不对成员求值**。
          // 真机取证：逐成员 `obj[name]` 求值会撞进宿主原生层，实测令 WPS 进程崩溃
          //（3 份崩溃报告，2 份调用栈逐帧一致：kso → etcore → etapi → jsetapi → ksojscore；
          // 崩溃点定位在整表 `Worksheet.Cells` 那一类表达式）。列名字用
          // Object.getOwnPropertyNames 是安全的，求值才是危险动作。
          const evaluateMembers = params?.evaluate === true;
          const maxMembers = Number.isFinite(Number(params?.maxMembers)) ? Number(params.maxMembers) : 150;

          /**
           * 已证实/高度可疑的危险成员：即使用户显式要求求值也跳过。
           * - 一组：**已有崩溃取证**——整表/整列/整行范围与整表集合，
           *   求值会构造覆盖整表的原生对象，实测崩溃。
           * - 另一组：**未经证实但保守跳过**——会触达宿主内部集合的原生 getter，
           *   历来是最容易出问题的一类；如需探测请单独用 wps_execute_script 并自行承担风险。
           */
          const DANGEROUS_CONFIRMED = new Set(['Cells', 'Rows', 'Columns', 'UsedRange', 'EntireRow', 'EntireColumn']);
          const DANGEROUS_SUSPECT = new Set([
            'CurrentRegion', 'Precedents', 'Dependents', 'SpecialCells',
            'Comment', 'CommentThreaded', 'Comments', 'CommentsThreaded',
            'Sort', 'SortFields', 'AutoFilter', 'Filters',
            'FormatConditions', 'Validation', 'Names', 'QueryTables', 'Connections',
            'ChartObjects', 'PivotCaches', 'PivotTables', 'ListObjects', 'Styles', 'CommandBars'
          ]);

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
          const skipped = [];
          const memberMap = new Set();
          let evaluatedCount = 0;

          // 遍历对象属性与原型链
          let curr = obj;
          let depth = 0;
          while (curr && depth < 3) {
            try {
              const names = Object.getOwnPropertyNames(curr);
              for (const name of names) {
                if (memberMap.has(name) || name.startsWith("__")) continue;
                memberMap.add(name);

                // 不求值：连 typeof 都不取（取 typeof 同样会触发一次属性访问）
                if (!evaluateMembers) {
                  methodList.push(name);
                  continue;
                }

                if (DANGEROUS_CONFIRMED.has(name)) {
                  skipped.push({ name, reason: "已证实：求值会构造整表范围对象并令 WPS 进程崩溃（ISS-89）" });
                  continue;
                }
                if (DANGEROUS_SUSPECT.has(name)) {
                  skipped.push({ name, reason: "保守跳过：会触达宿主内部集合的原生 getter；如确需探测请用 wps_execute_script 自行承担风险" });
                  continue;
                }
                if (evaluatedCount >= maxMembers) {
                  skipped.push({ name, reason: `已达 maxMembers(${maxMembers}) 上限，未求值` });
                  continue;
                }
                evaluatedCount++;

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
            evaluatedMembers: evaluateMembers,
            propertyCount: propList.length,
            methodCount: methodList.length,
            properties: propList.slice(0, 100),
            methods: methodList.sort(),
            skippedCount: skipped.length,
            skipped: skipped.slice(0, 40),
            message: evaluateMembers
              ? `成功完成对 [${targetExpr}] 的运行时反射（已求值 ${evaluatedCount} 个成员，跳过 ${skipped.length} 个危险/超限成员）`
              : `已列出 [${targetExpr}] 的成员名（**未求值**）。默认不求值是为了避免触及宿主原生 getter 导致 WPS 崩溃（ISS-89）；需要类型/取值时传 evaluate:true，并接受危险成员会被跳过。`
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
      headerPreview: headerPreview,
      // ISS-94 / ISS-64：把"能写但读不回来"的工作表级状态一并带回，
      // 让 AI 写完保护/标签色/冻结/筛选/条件格式后能自检，而不是只能相信 success。
      sheetState: readSheetState(app, sheet)
    };
  }

  /**
   * 读回条件格式的"是否启用"。
   *
   * 真机实测：本机 WPS 的 `FormatCondition` **没有 `Enabled` 属性**（读到 `undefined`）。
   * 直接 `Boolean(rule.Enabled)` 会把它写成 `false`，让读回看起来像"规则被禁用了"——
   * **读回说谎比没有读回更糟**。属性不存在时如实返回 null（未知），不猜测。
   */
  function readConditionEnabled(rule) {
    try {
      const raw = rule.Enabled;
      if (raw === undefined || raw === null) return null;
      return Boolean(raw);
    } catch (e) {
      return null;
    }
  }

  /**
   * 工作表级读回（只读，逐项 try/catch）。
   *
   * 这些状态原本都只有"写"没有"读"（问题台账 ISS-94）：调用方 set 完拿不到任何证据。
   * 冻结窗格依赖 `app.ActiveWindow`，**只在目标表处于活动状态时才有意义**，
   * 因此非活动表返回 null 并说明原因，避免拿别人的窗口状态冒充本表状态。
   */
  function readSheetState(app, sheet) {
    const state = {};
    const safe = (fn, fallback) => { try { const v = fn(); return v === undefined ? fallback : v; } catch (e) { return fallback; } };

    state.protection = safe(() => ({
      protectContents: Boolean(sheet.ProtectContents),
      protectDrawingObjects: Boolean(sheet.ProtectDrawingObjects),
      protectionMode: Boolean(sheet.ProtectionMode)
    }), null);

    state.tabColor = safe(() => {
      const color = sheet.Tab && sheet.Tab.Color !== undefined ? Number(sheet.Tab.Color) : null;
      if (color === null || !Number.isFinite(color)) return null;
      return excelColorToHex(color);
    }, null);

    state.autoFilter = safe(() => ({
      filterMode: Boolean(sheet.AutoFilterMode),
      range: sheet.AutoFilter && sheet.AutoFilter.Range ? sheet.AutoFilter.Range.Address() : null
    }), null);

    state.freezePanes = safe(() => {
      const active = app.ActiveWindow;
      if (!active) return { available: false, reason: "宿主没有活动窗口" };
      let isTarget = false;
      try { isTarget = String(active.ActiveSheet && active.ActiveSheet.Name) === String(sheet.Name); } catch (e) { isTarget = false; }
      if (!isTarget) {
        return { available: false, reason: "冻结窗格属窗口状态，只有目标表处于活动状态时才能读；请先激活该表" };
      }
      return {
        available: true,
        freezePanes: Boolean(active.FreezePanes),
        splitRow: safe(() => Number(active.SplitRow), null),
        splitColumn: safe(() => Number(active.SplitColumn), null)
      };
    }, { available: false, reason: "读取活动窗口状态失败" });

    // 条件格式：按已用区域扫描，返回每个区域的规则摘要（类型/优先级/是否启用）
    state.conditionalFormats = safe(() => {
      const rules = [];
      const used = sheet.UsedRange;
      if (!used) return rules;
      const fc = used.FormatConditions;
      const count = fc && typeof fc.Count === "number" ? fc.Count : 0;
      for (let i = 1; i <= Math.min(count, 50); i++) {
        try {
          const rule = fc.Item(i);
          rules.push({
            index: i,
            type: safe(() => Number(rule.Type), null),
            enabled: readConditionEnabled(rule),
            priority: safe(() => Number(rule.Priority), null),
            formula1: safe(() => (rule.Formula1 === undefined ? null : String(rule.Formula1)), null),
            interiorColor: safe(() => {
              const c = rule.Interior && rule.Interior.Color !== undefined ? Number(rule.Interior.Color) : null;
              return c === null || !Number.isFinite(c) ? null : excelColorToHex(c);
            }, null)
          });
        } catch (e) { /* 单条规则读失败不影响整体 */ }
      }
      return { count, scanned: Math.min(count, 50), rules };
    }, null);

    return state;
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
    if (wants("validation")) {
      // ISS-94：数据有效性原来只写不读，AI 设完下拉/范围校验后无法自检。
      result.validation = safeRead(() => {
        const v = range.Validation;
        const type = Number(v.Type);
        return {
          type,
          typeName: ({ 1: "xlValidateWholeNumber", 2: "xlValidateDecimal", 3: "xlValidateList", 4: "xlValidateDate", 5: "xlValidateTime", 6: "xlValidateTextLength", 7: "xlValidateCustom" })[type] || null,
          operator: Number(v.Operator),
          formula1: v.Formula1 === undefined ? null : String(v.Formula1),
          formula2: v.Formula2 === undefined ? null : String(v.Formula2),
          ignoreBlank: Boolean(v.IgnoreBlank),
          inCellDropdown: Boolean(v.InCellDropdown),
          prompt: v.InputMessage === undefined ? null : String(v.InputMessage),
          errorMessage: v.ErrorMessage === undefined ? null : String(v.ErrorMessage)
        };
      }, null);
    }

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
    const allowed = ["fontName", "fontSize", "bold", "fontColor", "backgroundColor", "numberFormat", "horizontalAlignment", "verticalAlignment", "wrapText", "rowHeight", "columnWidth", "merged", "mergeArea", "borders", "validation"];
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
      // 本机 WPS 上给跨多行的 `range.RowHeight` 赋值会被静默忽略（问题台账 ISS-22）。
      // 改为逐行写 Rows.Item(n).RowHeight，并读回校验。
      // 注意：本函数的区域变量名是 `range`（不是 targetRange）——写成 targetRange 会抛
      // ReferenceError，而 npm test 未覆盖该路径，只会在真机调用时暴露。
      const firstRow = range.Row;
      const rowTotal = range.Rows.Count;
      for (let i = 0; i < rowTotal; i++) {
        sheet.Rows.Item(firstRow + i).RowHeight = Number(rowHeight);
      }
      let readBackHeight = null;
      try { readBackHeight = Number(sheet.Rows.Item(firstRow).RowHeight); } catch (e) {}
      if (readBackHeight !== null && Number.isFinite(readBackHeight) && Math.abs(readBackHeight - Number(rowHeight)) > 0.6) {
        throw new Error(
          `行高设置未生效：请求 ${rowHeight}，第 ${firstRow} 行读回 ${readBackHeight}。` +
          `请确认该行未被合并单元格或工作表保护锁定。`
        );
      }
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

  // 10.05 单元格内**局部**格式（富文本）—— CAP-03
  /**
   * 只给单元格里的一段文字加格式（如"一句话里只把某段加粗/变红/改字号"）。
   *
   * 依据：WPS 宿主有 `Range.Characters(Start, Length)`（真机探测确认：返回对象且 `Count` 可用）。
   * 在此之前 AI 只能整格统一格式，做不到局部强调——这是"精细操控文档"最典型的诉求之一。
   *
   * 定位方式二选一：`find`（按文本找第 occurrence 处）或 `start`（1 基起点）+ `length`。
   * 写完逐项读回核对，任何一项没落上都报错，不做静默降级。
   */
  function formatTextSegment(app, params) {
    const {
      sheetName, workbookName, address,
      find, occurrence = 1, start, length,
      bold, italic, underline, fontColor, fontSize, fontName
    } = params || {};
    if (!address) throw new Error("缺少必要参数: address (如 'A1')");

    const sheet = getWorksheet(app, sheetName, workbookName);
    const range = sheet.Range(address);

    let text = "";
    try { text = range.Value2 === null || range.Value2 === undefined ? "" : String(range.Value2); } catch (e) {}
    if (!text) throw new Error(`单元格 ${range.Address ? range.Address() : address} 没有文本内容，无法做局部格式`);

    let begin = null;
    let segLen;
    if (find !== undefined && find !== null && String(find) !== "") {
      const query = String(find);
      const want = Math.max(1, Number(occurrence) || 1);
      let from = 0, hit = 0, idx = -1;
      for (;;) {
        idx = text.indexOf(query, from);
        if (idx < 0) break;
        hit++;
        if (hit === want) break;
        from = idx + 1;
      }
      if (idx < 0) {
        throw new Error(
          `在 ${address} 中找不到第 ${want} 处 "${query}"（共找到 ${hit} 处，单元格文本长度 ${text.length}）。` +
          `当前文本前 60 字：${text.slice(0, 60)}`
        );
      }
      begin = idx + 1; // Characters 是 1 基
      segLen = query.length;
    } else {
      if (!Number.isFinite(Number(start))) {
        throw new Error("需要提供 find（按文本定位）或 start（1 基起始位置，配合 length 使用）");
      }
      begin = Number(start);
      if (Number.isFinite(Number(length))) segLen = Number(length);
    }

    const chars = segLen === undefined ? range.Characters(begin) : range.Characters(begin, segLen);
    if (!chars) throw new Error(`无法定位字符片段（start=${begin}, length=${segLen}）：宿主未返回对象`);

    const requested = {};
    // VBA/WPS 的布尔格式用 -1/0（msoTrue/msoFalse）；传 JS 布尔在部分宿主上会被忽略
    if (bold !== undefined) { chars.Font.Bold = bold ? -1 : 0; requested.bold = Boolean(bold); }
    if (italic !== undefined) { chars.Font.Italic = italic ? -1 : 0; requested.italic = Boolean(italic); }
    if (underline !== undefined) { chars.Font.Underline = underline ? 2 : -4142; requested.underline = Boolean(underline); }
    if (fontSize !== undefined) { chars.Font.Size = Number(fontSize); requested.fontSize = Number(fontSize); }
    if (fontName) { chars.Font.Name = String(fontName); requested.fontName = String(fontName); }
    if (fontColor) {
      const bgr = hexToExcelColor(fontColor);
      if (bgr === null) throw new Error(`fontColor 无法解析: ${fontColor}（应为 #RRGGBB）`);
      chars.Font.Color = bgr;
      requested.fontColor = String(fontColor).toUpperCase();
    }

    // 读回核对：局部格式最怕"看着像生效了"
    const readBack = {};
    const mismatches = [];
    const num = (v) => (typeof v === "boolean" ? (v ? -1 : 0) : Number(v));
    try { readBack.segmentText = chars.Text === undefined ? null : String(chars.Text); } catch (e) { readBack.segmentText = null; }
    if (requested.bold !== undefined) {
      readBack.bold = num(safeRead(() => chars.Font.Bold, null)) !== 0;
      if (readBack.bold !== requested.bold) mismatches.push(`bold 请求 ${requested.bold} 读回 ${readBack.bold}`);
    }
    if (requested.italic !== undefined) {
      readBack.italic = num(safeRead(() => chars.Font.Italic, null)) !== 0;
      if (readBack.italic !== requested.italic) mismatches.push(`italic 请求 ${requested.italic} 读回 ${readBack.italic}`);
    }
    if (requested.fontSize !== undefined) {
      readBack.fontSize = Number(safeRead(() => chars.Font.Size, null));
      if (Math.abs(readBack.fontSize - requested.fontSize) > 0.26) mismatches.push(`fontSize 请求 ${requested.fontSize} 读回 ${readBack.fontSize}`);
    }
    if (requested.fontName !== undefined) {
      readBack.fontName = safeRead(() => chars.Font.Name, null);
      if (String(readBack.fontName) !== requested.fontName) mismatches.push(`fontName 请求 ${requested.fontName} 读回 ${readBack.fontName}`);
    }
    if (requested.fontColor !== undefined) {
      readBack.fontColor = excelColorToHex(safeRead(() => chars.Font.Color, null));
      if (String(readBack.fontColor).toUpperCase() !== requested.fontColor) mismatches.push(`fontColor 请求 ${requested.fontColor} 读回 ${readBack.fontColor}`);
    }
    if (requested.underline !== undefined) {
      readBack.underline = num(safeRead(() => chars.Font.Underline, null)) !== -4142;
      if (readBack.underline !== requested.underline) mismatches.push(`underline 请求 ${requested.underline} 读回 ${readBack.underline}`);
    }

    if (mismatches.length > 0) {
      throw new Error(
        `局部格式未完全生效：${mismatches.join("；")}。` +
        `该单元格文本为 "${text.slice(0, 60)}"，片段起点 ${begin}、长度 ${segLen === undefined ? "至末尾" : segLen}。`
      );
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address ? range.Address() : address,
      cellText: text.slice(0, 120),
      segmentStart: begin,
      segmentLength: segLen === undefined ? text.length - begin + 1 : segLen,
      segmentText: readBack.segmentText,
      requested,
      readBack,
      message: `已将 [${sheet.Name}] ${address} 的第 ${begin} 个字符起 ${segLen === undefined ? "到末尾" : segLen + " 个字符"}「${readBack.segmentText ?? ""}」设置为指定格式，并读回核对通过`
    };
  }

  // 10.1 条件格式与数据条/色阶
  function addConditionalFormatting(app, params) {
    // CAP-30 条件格式读回：此前只有"加"没有"读"，AI 无法回答
    // "这块区域现在挂了哪些规则"，改完也只能靠人看。
    // 覆盖单元格值 / 公式 / 色阶 / 数据条 / Top10 / 图标集 / 重复值 / 唯一值 / 文本 / 空值。
    if ((params || {}).action === "read") {
      const { sheetName, address, workbookName } = params || {};
      const sheet = getWorksheet(app, sheetName, workbookName);
      const TYPE = {
        1: "cell_value", 2: "formula", 3: "color_scale", 4: "data_bar", 5: "top10",
        6: "icon_set", 8: "unique_values", 9: "text_contains", 10: "blanks", 11: "time_period",
        12: "above_average", 13: "no_blanks", 16: "duplicate_values"
      };
      const OPER = { 1: "between", 2: "not_between", 3: "equal", 4: "not_equal", 5: "greater_than", 6: "less_than", 7: "greater_equal", 8: "less_equal" };
      const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
      const toHex = (v) => (v === null || v === undefined || v < 0 || v === 16777215) ? null
        : "#" + [v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff].map(n => n.toString(16).padStart(2, "0")).join("").toUpperCase();

      const ranges = address ? [sheet.Range(address)] : (() => {
        // 整表：宿主没有"枚举所有条件格式区域"的 API，按行扫描已用区域
        const ur = sheet.UsedRange, out = [];
        const colName = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
        const rows = Math.min(ur.Rows.Count, 2000), cols = Math.min(ur.Columns.Count, 100);
        for (let r = ur.Row; r < ur.Row + rows; r++) {
          for (let c = ur.Column; c < ur.Column + cols; c++) out.push(sheet.Range(`${colName(c)}${r}`));
        }
        return out;
      })();

      const items = [];
      for (const rng of ranges) {
        let fcs = null;
        try { fcs = rng.FormatConditions; } catch (e) { continue; }
        const n = g(() => Number(fcs.Count), 0);
        if (!n) continue;
        const entry = { address: g(() => rng.Address(), null), rules: [] };
        for (let i = 1; i <= n; i++) {
          const fc = (() => { try { return fcs.Item(i); } catch (e) { return null; } })();
          if (!fc) continue;
          const tc = g(() => Number(fc.Type), null);
          const rule = {
            index: i,
            type: TYPE[tc] || `unknown(${tc})`,
            typeCode: tc,
            operator: OPER[g(() => Number(fc.Operator), null)] || null,
            formula1: g(() => fc.Formula1, null),
            formula2: g(() => fc.Formula2, null),
            priority: g(() => Number(fc.Priority), null),
            stopIfTrue: g(() => Boolean(fc.StopIfTrue), null),
            fillColor: toHex(g(() => Number(fc.Interior.Color), null)),
            fontColor: toHex(g(() => Number(fc.Font.Color), null)),
            fontBold: g(() => Boolean(fc.Font.Bold), null),
            // 图标集 / 数据条 / 色阶的特有属性（宿主不支持的会落到 null）
            iconSet: g(() => String(fc.IconSet), null),
            showIconOnly: g(() => Boolean(fc.ShowIconOnly), null),
            barColor: toHex(g(() => Number(fc.BarColor), null)),
            text: g(() => fc.Text, null)
          };
          entry.rules.push(rule);
        }
        items.push(entry);
      }
      const total = items.reduce((s, x) => s + x.rules.length, 0);
      return {
        success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name,
        areaCount: items.length, ruleCount: total, areas: items, warnings: [],
        message: `工作表 [${sheet.Name}] 上 ${items.length} 个区域共 ${total} 条条件格式规则`
      };
    }
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
      clearExisting = false,
      // CAP-06 新增：图标集 / Top-N / 重复值 / 公式 / 文字包含
      iconSet,
      iconThresholds,
      topBottom,
      topRank,
      topPercent,
      containsText
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
    } else if (ruleType === "icon_set") {
      // CAP-06：红黄绿灯这类"业务信号"是最常见的报表诉求。
      // 真机探测：宿主有 AddIconSetCondition（AddTextString 不存在，文字规则走公式规则）。
      const ICON_SET_CODES = {
        "3_arrows": 1, "3_arrows_gray": 2, "3_flags": 3,
        "3_traffic_lights": 4, "3_traffic_lights_rimmed": 5,
        "3_signs": 6, "3_symbols": 7, "3_symbols_circled": 8,
        "4_arrows": 9, "4_arrows_gray": 10, "4_red_to_black": 11,
        "4_ratings": 12, "4_traffic_lights": 13,
        "5_arrows": 14, "5_arrows_gray": 15, "5_quarters": 16,
        "5_ratings": 17, "5_boxes": 18
      };
      const iconSetName = iconSet || "3_traffic_lights";
      const code = ICON_SET_CODES[iconSetName];
      if (code === undefined) {
        throw new Error(`不支持的 iconSet: ${iconSetName}（可用: ${Object.keys(ICON_SET_CODES).join(" / ")}）`);
      }
      const icons = range.FormatConditions.AddIconSetCondition();
      icons.IconSet = app.IconSets ? app.IconSets.Item(code) : code;
      // 阈值：不给就用宿主默认；给了就逐档设置（最多 5 档）
      if (Array.isArray(iconThresholds) && iconThresholds.length > 0) {
        for (let i = 0; i < Math.min(iconThresholds.length, 5); i++) {
          const spec = iconThresholds[i];
          const crit = icons.IconCriteria.Item(i + 2); // 第 1 档是"最低值"，从第 2 档开始可设阈值
          if (spec && typeof spec === "object") {
            if (spec.type !== undefined) crit.Type = Number(spec.type);
            if (spec.operator !== undefined) crit.Operator = Number(spec.operator);
            if (spec.value !== undefined) crit.Value = spec.value;
          } else if (spec !== undefined && spec !== null) {
            crit.Value = spec;
          }
        }
      }
      if (backgroundColor) {
        const bgc = hexToExcelColor(backgroundColor);
        if (bgc !== null) icons.Interior.Color = bgc;
      }
    } else if (ruleType === "top10") {
      // CAP-06：Top/Bottom N 或百分比
      const top = range.FormatConditions.AddTop10();
      if (topBottom !== undefined) top.TopBottom = Number(topBottom);   // 1=xlTop 2=xlBottom
      if (topRank !== undefined) top.Rank = Number(topRank);
      if (topPercent !== undefined) top.Percent = topPercent ? -1 : 0;
      if (backgroundColor) {
        const bgc = hexToExcelColor(backgroundColor);
        if (bgc !== null) top.Interior.Color = bgc;
      }
      if (fontColor) {
        const fgc = hexToExcelColor(fontColor);
        if (fgc !== null) top.Font.Color = fgc;
      }
    } else if (ruleType === "duplicate_values" || ruleType === "unique_values") {
      // CAP-06：重复值/唯一值高亮（对账、查重最常用）
      const dv = range.FormatConditions.AddUniqueValues();
      dv.DupeUnique = ruleType === "duplicate_values" ? 1 : 2; // 1=xlDuplicate 2=xlUnique
      if (backgroundColor) {
        const bgc = hexToExcelColor(backgroundColor);
        if (bgc !== null) dv.Interior.Color = bgc;
      }
      if (fontColor) {
        const fgc = hexToExcelColor(fontColor);
        if (fgc !== null) dv.Font.Color = fgc;
      }
    } else if (ruleType === "formula" || ruleType === "text_contains") {
      // CAP-06：公式规则（xlExpression=2）。文字包含没有独立宿主方法（真机确认 AddTextString 不存在），
      // 用标准的 SEARCH 公式实现——这也是业务上更通用的做法。
      let expr = formula1;
      if (ruleType === "text_contains") {
        if (!containsText) throw new Error("text_contains 规则必须提供 containsText");
        const anchor = String(address).split(":")[0];
        expr = `=ISNUMBER(SEARCH("${String(containsText).replace(/"/g, '""')}",${anchor}))`;
      }
      if (!expr) throw new Error("formula 规则必须提供 formula1（如 '=A1>100'）");
      const fx = range.FormatConditions.Add(2, 0, String(expr)); // 2=xlExpression, 0=xlNone
      if (backgroundColor) {
        const bgc = hexToExcelColor(backgroundColor);
        if (bgc !== null) fx.Interior.Color = bgc;
      }
      if (fontColor) {
        const fgc = hexToExcelColor(fontColor);
        if (fgc !== null) fx.Font.Color = fgc;
      }
    } else if (ruleType === "clear") {
      range.FormatConditions.Delete();
    } else {
      throw new Error(
        `未知的条件格式类型: ${ruleType}（支持 cell_value, data_bar, color_scale, icon_set, ` +
        `top10, duplicate_values, unique_values, formula, text_contains, clear）`
      );
    }

    // 读回核对：条件格式是"只写不读"的重灾区，写完必须能确认到底加了几条
    let appliedCount = null;
    const conditions = [];
    try {
      const fc = range.FormatConditions;
      appliedCount = Number(fc.Count);
      for (let i = 1; i <= Math.min(appliedCount, 20); i++) {
        const rule = fc.Item(i);
        conditions.push({
          index: i,
          type: Number(safeRead(() => rule.Type, null)),
          enabled: readConditionEnabled(rule),
          priority: safeRead(() => Number(rule.Priority), null),
          formula1: safeRead(() => (rule.Formula1 === undefined ? null : String(rule.Formula1)), null)
        });
      }
    } catch (e) {}

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      address: range.Address(),
      ruleType: ruleType,
      appliedConditionCount: appliedCount,
      conditions,
      message: `已成功在 ${range.Address()} 应用 ${ruleType} 条件格式${appliedCount === null ? "" : `（该区域现有 ${appliedCount} 条规则）`}`
    };
  }

  // 10.2 冻结窗格吸顶
  function freezePanes(app, params) {
    const { sheetName, workbookName, freezeRowIndex, freezeColumnIndex, unfreeze, action = "apply" } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}

    const win = app.ActiveWindow;
    if (!win) throw new Error("无法获取 WPS 活动窗口句柄");

    // CAP-31 冻结窗格读回：只写不读的话，AI 无法确认"到底冻了几行"，
    // 也就没法在多次操作后知道自己当前处于什么状态。
    if (action === "read") {
      const frozen = (() => { try { return Boolean(win.FreezePanes); } catch (e) { return null; } })();
      const splitRow = (() => { try { return Number(win.SplitRow) || 0; } catch (e) { return null; } })();
      const splitColumn = (() => { try { return Number(win.SplitColumn) || 0; } catch (e) { return null; } })();
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        frozen,
        // 对外给"冻结到第几行/列"（1 基，更接近用户说法），同时保留原始 SplitRow/SplitColumn
        freezeRowIndex: splitRow === null ? null : splitRow + 1,
        freezeColumnIndex: splitColumn === null ? null : splitColumn + 1,
        splitRow, splitColumn,
        warnings: [],
        message: frozen
          ? `工作表 [${sheet.Name}] 已冻结：第 ${splitRow} 行之上的 ${splitRow} 行、第 ${splitColumn} 列之左的 ${splitColumn} 列保持不动`
          : `工作表 [${sheet.Name}] 当前未冻结窗格`
      };
    }

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

  // 10.25 打印与分页设置（CAP-01）
  /**
   * 读写工作表的打印设置：打印区域、打印标题行列、纸张、方向、页边距、页眉页脚、
   * 网格线/居中/缩放，以及手动分页符。
   *
   * 依据：真机探测确认 `PageSetup` 上这些属性都存在（PrintArea / PrintTitleRows / Orientation /
   * PaperSize / Zoom / CenterHorizontally / PrintGridlines / 各边距 / CenterHeader / LeftFooter /
   * FitToPagesWide），`HPageBreaks` / `VPageBreaks` 也可用。
   * 此前 AI 做不出"可直接打印装订"的报表。
   *
   * `action: "read"` 只读回当前设置；默认 `"apply"` 先写后**逐项读回核对**。
   */
  function configurePrintLayout(app, params) {
    const {
      sheetName, workbookName, action = "apply",
      printArea, printTitleRows, printTitleColumns,
      orientation, paperSize, zoom, fitToPagesWide, fitToPagesTall,
      centerHorizontally, centerVertically, printGridlines,
      leftMargin, rightMargin, topMargin, bottomMargin, headerMargin, footerMargin,
      centerHeader, leftHeader, rightHeader, centerFooter, leftFooter, rightFooter,
      clearPrintArea, addHorizontalPageBreak, addVerticalPageBreak, clearPageBreaks
    } = params || {};

    const sheet = getWorksheet(app, sheetName, workbookName);
    const ps = sheet.PageSetup;

    // VBA 常量：方向 xlPortrait=1 / xlLandscape=2；纸张 xlPaperLetter=1 / xlPaperA4=9 / xlPaperA3=8 …
    const ORIENT = { portrait: 1, landscape: 2 };
    const PAPER = { letter: 1, a4: 9, a3: 8, legal: 5, b5: 13 };
    const applied = {};
    const warnings = [];

    if (action !== "read") {
      if (clearPrintArea) {
        ps.PrintArea = "";
        ps.PrintTitleRows = "";
        ps.PrintTitleColumns = "";
        applied.clearedPrintArea = true;
      }
      if (printArea !== undefined) { ps.PrintArea = printArea === null ? "" : String(printArea); applied.printArea = String(printArea); }
      if (printTitleRows !== undefined) { ps.PrintTitleRows = printTitleRows === null ? "" : String(printTitleRows); applied.printTitleRows = String(printTitleRows); }
      if (printTitleColumns !== undefined) { ps.PrintTitleColumns = printTitleColumns === null ? "" : String(printTitleColumns); applied.printTitleColumns = String(printTitleColumns); }
      if (orientation !== undefined) {
        const code = typeof orientation === "number" ? orientation : ORIENT[String(orientation).toLowerCase()];
        if (code === undefined) throw new Error(`不支持的 orientation: ${orientation}（可用 portrait / landscape，或 1 / 2）`);
        ps.Orientation = code;
        applied.orientation = code;
      }
      if (paperSize !== undefined) {
        const code = typeof paperSize === "number" ? paperSize : PAPER[String(paperSize).toLowerCase()];
        if (code === undefined) throw new Error(`不支持的 paperSize: ${paperSize}（可用 a4 / a3 / letter / legal / b5，或 VBA 常量）`);
        ps.PaperSize = code;
        applied.paperSize = code;
      }
      if (zoom !== undefined) { ps.Zoom = Number(zoom); applied.zoom = Number(zoom); }
      if (fitToPagesWide !== undefined) { ps.FitToPagesWide = Number(fitToPagesWide); applied.fitToPagesWide = Number(fitToPagesWide); }
      if (fitToPagesTall !== undefined) { ps.FitToPagesTall = Number(fitToPagesTall); applied.fitToPagesTall = Number(fitToPagesTall); }
      if (centerHorizontally !== undefined) { ps.CenterHorizontally = Boolean(centerHorizontally); applied.centerHorizontally = Boolean(centerHorizontally); }
      if (centerVertically !== undefined) { ps.CenterVertically = Boolean(centerVertically); applied.centerVertically = Boolean(centerVertically); }
      if (printGridlines !== undefined) { ps.PrintGridlines = Boolean(printGridlines); applied.printGridlines = Boolean(printGridlines); }
      for (const [key, value] of Object.entries({ leftMargin, rightMargin, topMargin, bottomMargin, headerMargin, footerMargin })) {
        if (value === undefined) continue;
        ps[key.charAt(0).toUpperCase() + key.slice(1)] = Number(value);
        applied[key] = Number(value);
      }
      for (const [key, value] of Object.entries({ centerHeader, leftHeader, rightHeader, centerFooter, leftFooter, rightFooter })) {
        if (value === undefined) continue;
        ps[key.charAt(0).toUpperCase() + key.slice(1)] = String(value);
        applied[key] = String(value);
      }

      if (clearPageBreaks) {
        try { sheet.ResetAllPageBreaks(); applied.clearedPageBreaks = true; } catch (e) { warnings.push(`清除分页符失败: ${e.message}`); }
      }
      if (Number.isFinite(Number(addHorizontalPageBreak))) {
        try { sheet.HPageBreaks.Add(sheet.Rows.Item(Number(addHorizontalPageBreak))); } catch (e) { warnings.push(`添加水平分页符失败: ${e.message}`); }
      }
      if (addVerticalPageBreak !== undefined) {
        if (!Number.isFinite(Number(addVerticalPageBreak))) {
          throw new Error(`addVerticalPageBreak 需传列号（数字），收到 ${JSON.stringify(addVerticalPageBreak)}；列字母请先换算成序号`);
        }
        try { sheet.VPageBreaks.Add(sheet.Columns.Item(Number(addVerticalPageBreak))); } catch (e) { warnings.push(`添加垂直分页符失败: ${e.message}`); }
      }
    }

    // 读回：打印设置最怕"设了但没生效"
    const safe = (fn) => { try { const v = fn(); return v === undefined ? null : v; } catch (e) { return null; } };
    const settings = {
      printArea: safe(() => String(ps.PrintArea || "")),
      printTitleRows: safe(() => String(ps.PrintTitleRows || "")),
      printTitleColumns: safe(() => String(ps.PrintTitleColumns || "")),
      orientation: safe(() => Number(ps.Orientation)),
      paperSize: safe(() => Number(ps.PaperSize)),
      zoom: safe(() => Number(ps.Zoom)),
      fitToPagesWide: safe(() => Number(ps.FitToPagesWide)),
      fitToPagesTall: safe(() => Number(ps.FitToPagesTall)),
      centerHorizontally: safe(() => Boolean(ps.CenterHorizontally)),
      centerVertically: safe(() => Boolean(ps.CenterVertically)),
      printGridlines: safe(() => Boolean(ps.PrintGridlines)),
      margins: {
        left: safe(() => Number(ps.LeftMargin)), right: safe(() => Number(ps.RightMargin)),
        top: safe(() => Number(ps.TopMargin)), bottom: safe(() => Number(ps.BottomMargin)),
        header: safe(() => Number(ps.HeaderMargin)), footer: safe(() => Number(ps.FooterMargin))
      },
      headers: {
        center: safe(() => String(ps.CenterHeader || "")), left: safe(() => String(ps.LeftHeader || "")),
        right: safe(() => String(ps.RightHeader || ""))
      },
      footers: {
        center: safe(() => String(ps.CenterFooter || "")), left: safe(() => String(ps.LeftFooter || "")),
        right: safe(() => String(ps.RightFooter || ""))
      },
      pageBreaks: {
        horizontal: safe(() => Number(sheet.HPageBreaks.Count)),
        vertical: safe(() => Number(sheet.VPageBreaks.Count))
      }
    };
    settings.orientationName = settings.orientation === 1 ? "portrait" : settings.orientation === 2 ? "landscape" : null;

    // 打印区域写后核对（最容易"设了却没生效"的一项）
    if (action !== "read" && printArea !== undefined && printArea !== null && String(printArea) !== "") {
      const norm = (s) => String(s || "").toLowerCase().replace(/\$/g, "");
      if (norm(printArea) !== norm(settings.printArea)) {
        throw new Error(`打印区域设置未生效：请求 "${printArea}"，读回 "${settings.printArea}"`);
      }
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      action,
      applied,
      settings,
      warnings,
      message: action === "read"
        ? `已读回 [${sheet.Name}] 的打印设置（打印区域 ${settings.printArea || "未设置"}，${settings.orientationName || "方向未知"}）`
        : `已应用并读回核对 [${sheet.Name}] 的打印设置`
    };
  }

  // 10.26 导出工作表/工作簿为 PDF（CAP-02）
  /**
   * 用宿主的 `ExportAsFixedFormat` 导出 PDF（xlTypePDF = 0）。
   *
   * **加载项没有文件系统访问，无法确认落盘**——本函数只负责发起导出并回传路径与耗时，
   * **落盘校验由桥接侧完成**（与 Word 预览同一套机制）。输出目录必须是宿主可写位置，
   * 见 `src/bridge/runtime.ts` 的 `previewDir()`：WPS 是沙箱应用，`~/.wps-bridge` 之类写不进去。
   */
  function exportSheetPdf(app, params) {
    const { sheetName, workbookName, outputPath, scope = "workbook", quality = "standard" } = params || {};
    if (!outputPath) throw new Error("缺少必要参数: outputPath（由桥接指定的输出路径）");
    // Quality: xlQualityStandard = 0 / xlQualityMinimum = 1
    const qualityCode = String(quality).toLowerCase() === "minimum" ? 1 : 0;
    const target = scope === "sheet" ? getWorksheet(app, sheetName, workbookName) : getWorkbook(app, workbookName);

    const startedAt = Date.now();
    try {
      // ExportAsFixedFormat(Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish)
      target.ExportAsFixedFormat(0, String(outputPath), qualityCode, true, false, undefined, undefined, false);
    } catch (e) {
      return { success: false, outputPath: String(outputPath), scope, hostError: e.message, message: `导出 PDF 失败：${e.message}` };
    }
    return {
      success: true,
      outputPath: String(outputPath),
      scope,
      elapsedMs: Date.now() - startedAt,
      hostCannotVerify: true,
      message: `宿主已接受导出请求（${scope === "sheet" ? "工作表 " + (sheetName || "活动表") : "整个工作簿"} → ${outputPath}）；文件是否真的写出由桥接侧校验`
    };
  }

  // 10.27 表格矢量绘图（CAP-07）
  //
  // 依据真机探测（WPS 12.1.28496）：宿主 `Sheet.Shapes` 全套可用——
  //   AddShape / AddLine / AddTextbox / **AddTextEffect（艺术字）** / AddPicture / BuildFreeform，
  //   分组走 `Shapes.Range([名...]).Group()`（**WPS 表格能分组，Office.js 侧这条未打通**）。
  // 这是 WPS 相对 Microsoft Excel 的**能力优势**：艺术字与自由曲线在 Office.js 侧没有对应 API。
  // 此前我们一个表格绘图工具都没有，AI 只能用 wps_execute_script 手写。

  /** 几何形状枚举（msoAutoShapeType 常用值）。 */
  const AUTO_SHAPE_TYPES = {
    rectangle: 1, rounded_rectangle: 5, oval: 9, ellipse: 9, diamond: 4, triangle: 7,
    right_triangle: 8, pentagon: 51, hexagon: 10, arrow_right: 33, arrow_left: 34,
    arrow_up: 35, arrow_down: 36, chevron: 52, cross: 11, star_5: 12, star_6: 13,
    callout_rounded: 106, cloud: 179, heart: 21, lightning: 22, sun: 24, moon: 23,
    can: 13, cube: 17, donut: 18, flow_chart_process: 61, flow_chart_decision: 63,
    flow_chart_terminator: 68, flow_chart_data: 62, line_horizontal: 130, text_box: 17
  };

  /** 形状文字的默认字体。中文报表里微软雅黑比宿主默认的宋体更清晰、字重更统一。 */
  const DEFAULT_SHAPE_FONT = "微软雅黑";

  function resolveShapeType(raw) {
    if (raw === undefined || raw === null || raw === "") return AUTO_SHAPE_TYPES.rectangle;
    if (typeof raw === "number") return raw;
    const key = String(raw).toLowerCase().trim();
    if (AUTO_SHAPE_TYPES[key] !== undefined) return AUTO_SHAPE_TYPES[key];
    if (/^\d+$/.test(key)) return parseInt(key, 10);
    throw new Error(`不认识的形状类型 "${raw}"。可用：${Object.keys(AUTO_SHAPE_TYPES).join(" / ")}`);
  }

  /**
   * 判断某个十六进制底色偏亮还是偏暗，用于给形状文字选黑字还是白字。
   *
   * 为什么需要：WPS 给**自选图形默认白字**。如果调用方填了白底又不指定字色，
   * 就会得到"白字白底"——文字完全看不见，而且没有任何报错（真机踩到过：
   * 流程框全部隐形）。这里按底色亮度自动选，调用方显式给 fontColor 时以调用方为准。
   */
  function isLightFill(hex) {
    const rgb = hexToExcelColor(hex);
    if (rgb === null || rgb === undefined) return null;   // 无底色 → 交给宿主默认
    const r = rgb & 0xff, g = (rgb >> 8) & 0xff, b = (rgb >> 16) & 0xff;
    return (0.299 * r + 0.587 * g + 0.114 * b) > 150;
  }

  /** 单元格地址取值：宿主方法是原生实现，**必须保留接收者**直接调用（摘出来会丢 this）。 */
  function shapeAddressOf(cell) {
    try { return String(cell.Address()); } catch (e) {
      try { return String(cell.Address); } catch (e2) { return ""; }
    }
  }

  /** 读回单个形状的完整状态。 */
  function readShape(shape, sheet) {
    const safe = (fn, fallback) => { try { const v = fn(); return v === undefined ? fallback : v; } catch (e) { return fallback; } };
    let text = null;
    try { text = String(shape.TextFrame.Characters().Text); } catch (e) { text = null; }
    return {
      name: safe(() => String(shape.Name), null),
      type: safe(() => Number(shape.Type), null),
      autoShapeType: safe(() => Number(shape.AutoShapeType), null),
      left: safe(() => Math.round(Number(shape.Left) * 100) / 100, null),
      top: safe(() => Math.round(Number(shape.Top) * 100) / 100, null),
      width: safe(() => Math.round(Number(shape.Width) * 100) / 100, null),
      height: safe(() => Math.round(Number(shape.Height) * 100) / 100, null),
      rotation: safe(() => Number(shape.Rotation), null),
      visible: safe(() => Number(shape.Visible) !== 0, null),
      fillColor: safe(() => { const v = Number(shape.Fill.ForeColor.RGB); return Number.isFinite(v) && v >= 0 ? excelColorToHex(v) : null; }, null),
      lineColor: safe(() => { const v = Number(shape.Line.ForeColor.RGB); return Number.isFinite(v) && v >= 0 ? excelColorToHex(v) : null; }, null),
      lineWeight: safe(() => Number(shape.Line.Weight), null),
      text,
      // 真实对齐（Excel 对象模型：HorizontalAlignment/VerticalAlignment）。
      // 放进 readShape 而不是只在 addShape 里，这样 list_shapes 也能核对——
      // 「框居中了但字没居中」这类静默失败必须能被读回发现。
      textAlign: safe(() => { const h = Number(shape.TextFrame.HorizontalAlignment); return h === -4108 ? "center" : h === -4152 ? "right" : h === -4131 ? "left" : null; }, null),
      textVAlign: safe(() => { const v = Number(shape.TextFrame.VerticalAlignment); return v === -4108 ? "middle" : v === -4107 ? "bottom" : v === -4160 ? "top" : null; }, null),
      topLeftCell: safe(() => shapeAddressOf(shape.TopLeftCell), null),
      isGroup: safe(() => Number(shape.Type) === 6, false),
      groupItemCount: safe(() => (Number(shape.Type) === 6 && shape.GroupItems ? Number(shape.GroupItems.Count) : null), null)
    };
  }

  /** 找出形状：优先按名字，其次按序号（1 基）。 */
  function findShape(sheet, params) {
    const shapes = sheet.Shapes;
    if (params.name || params.shapeName) {
      const name = String(params.name || params.shapeName);
      for (let i = 1; i <= shapes.Count; i++) {
        const sh = shapes.Item(i);
        try { if (String(sh.Name) === name) return sh; } catch (e) {}
      }
      const available = [];
      for (let i = 1; i <= Math.min(shapes.Count, 30); i++) {
        try { available.push(String(shapes.Item(i).Name)); } catch (e) {}
      }
      throw new Error(`找不到名为 "${name}" 的形状。现有形状（最多 30 个）：${available.join(" / ") || "无"}`);
    }
    if (Number.isFinite(Number(params.shapeIndex))) {
      const idx = Number(params.shapeIndex);
      if (idx < 1 || idx > shapes.Count) throw new Error(`形状序号 ${idx} 越界（当前共 ${shapes.Count} 个）`);
      return shapes.Item(idx);
    }
    throw new Error("需要提供 name（形状名）或 shapeIndex（序号，从 1 开始）");
  }

  /**
   * 画矢量形状：几何形状 / 直线 / 文本框 / **艺术字**。
   * 写完**读回真实几何与样式**，宿主没接受某个属性时写进 warnings，不做假成功。
   */
  function addShape(app, params) {
    const { sheetName, workbookName, kind = "geometric", shapeType, text, name,
            left = 0, top = 0, width = 160, height = 80, rotation,
            x1, y1, x2, y2, fillColor, lineColor, lineWeight,
            wordArtPreset, fontName, fontSize, bold, italic, fontColor,
            textAlign, textVAlign, marginLeft, marginRight, marginTop, marginBottom } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shapes = sheet.Shapes;
    const warnings = [];

    let shape = null;
    const k = String(kind).toLowerCase();
    if (k === "geometric" || k === "autoshape") {
      shape = shapes.AddShape(resolveShapeType(shapeType), Number(left), Number(top), Number(width), Number(height));
    } else if (k === "line" || k === "connector") {
      shape = shapes.AddLine(
        Number(x1 !== undefined ? x1 : left), Number(y1 !== undefined ? y1 : top),
        Number(x2 !== undefined ? x2 : Number(left) + Number(width)),
        Number(y2 !== undefined ? y2 : Number(top) + Number(height))
      );
    } else if (k === "textbox" || k === "text") {
      shape = shapes.AddTextbox(1, Number(left), Number(top), Number(width), Number(height));
      // 文本框默认**不画边框、不填底色**——否则页面上每段文字都套一个框。
      // 调用方显式传 fillColor/lineColor 时才打开对应可见性（见下方）。
      try { shape.Line.Visible = 0; } catch (e) {}
      try { shape.Fill.Visible = 0; } catch (e) {}
      // 关掉自动缩放：宿主默认会把文本框收缩到文字宽度，导致"按给定宽度算的居中"全部偏掉
      // （真机踩到：柱顶数据标签偏离柱心 1~2pt）。宽度按调用方给的固定，版面才可预测。
      try { shape.TextFrame.AutoSize = 0; } catch (e) {}
      try { shape.TextFrame.WordWrap = 0; } catch (e) {}
    } else if (k === "wordart" || k === "texteffect") {
      // 艺术字：WPS 相对 Office.js 的独有能力
      const preset = Number.isFinite(Number(wordArtPreset)) ? Number(wordArtPreset) : 0;
      shape = shapes.AddTextEffect(
        preset, String(text === undefined ? "" : text), String(fontName || "宋体"),
        Number(fontSize || 36), bold === false ? 0 : -1, italic ? -1 : 0, Number(left), Number(top)
      );
    } else {
      throw new Error(`不支持的形状种类: ${kind}（可用 geometric / line / textbox / wordart）`);
    }
    if (!shape) throw new Error("宿主没有返回形状对象（创建失败）");

    if (name) { try { shape.Name = String(name); } catch (e) { warnings.push(`设置名字失败: ${e.message}`); } }
    if (k !== "line" && Number.isFinite(Number(rotation))) {
      try { shape.Rotation = Number(rotation); } catch (e) { warnings.push(`设置旋转失败: ${e.message}`); }
    }
    if (fillColor) {
      const rgb = hexToExcelColor(fillColor);
      if (rgb === null) warnings.push(`fillColor 无法解析: ${fillColor}`);
      else { try { shape.Fill.Visible = -1; shape.Fill.ForeColor.RGB = rgb; } catch (e) { warnings.push(`设置填充失败: ${e.message}`); } }
    }
    if (lineColor) {
      const rgb = hexToExcelColor(lineColor);
      if (rgb === null) warnings.push(`lineColor 无法解析: ${lineColor}`);
      else { try { shape.Line.Visible = -1; shape.Line.ForeColor.RGB = rgb; } catch (e) { warnings.push(`设置线条色失败: ${e.message}`); } }
    }
    if (Number.isFinite(Number(lineWeight))) {
      try { shape.Line.Weight = Number(lineWeight); } catch (e) { warnings.push(`设置线宽失败: ${e.message}`); }
    }
    // 文字：几何形状/文本框都能写字（艺术字的文字在创建时给）。
    // 字号/加粗/斜体/字体/字色对**所有形状类型**生效——原来只对艺术字生效，
    // 于是"画个文本框并给它 22pt 加粗"这类最常见需求被静默忽略。
    if (k !== "wordart" && k !== "texteffect" && text !== undefined && text !== null && String(text) !== "") {
      try {
        const chars = shape.TextFrame.Characters();
        chars.Text = String(text);
        if (Number.isFinite(Number(fontSize))) chars.Font.Size = Number(fontSize);
        if (bold !== undefined) chars.Font.Bold = bold ? -1 : 0;
        if (italic !== undefined) chars.Font.Italic = italic ? -1 : 0;
        chars.Font.Name = String(fontName || DEFAULT_SHAPE_FONT);
        if (fontColor) {
          const fc = hexToExcelColor(fontColor);
          if (fc !== null) chars.Font.Color = fc;
        } else {
          // 未指定字色：按填充色亮度选黑/白，避免"白字白底"隐形
          const light = isLightFill(fillColor);
          if (light === true) chars.Font.Color = hexToExcelColor("#333333");
          else if (light === false) chars.Font.Color = hexToExcelColor("#FFFFFF");
        }
        // 对齐：文本框默认左对齐/顶对齐，自选图形默认居中/垂直居中。
        //
        // ⚠️ 这里是 **Excel 的对象模型**，不是 PowerPoint 的：
        //   `TextFrame.TextRange.ParagraphFormat.Alignment` 在 Excel 里**不存在**
        //   （TextRange=undefined、Characters().ParagraphFormat=undefined），
        //   正确属性是 `TextFrame.HorizontalAlignment` / `VerticalAlignment`。
        // 之前用错的那条路径被 try/catch 吞掉，于是"框居中了、框里的字还是居左"，
        // 既不报错也看不出来（使用者一眼就看出来了）——所以这里改成**失败即告警**，不再静默。
        const tf = shape.TextFrame;
        const H = { left: -4131, center: -4108, right: -4152 };   // xlHAlign*
        const V = { top: -4160, middle: -4108, bottom: -4107 };   // xlVAlign*
        const wantH = textAlign || ((k !== "textbox" && k !== "text") ? "center" : null);
        const wantV = textVAlign || ((k !== "textbox" && k !== "text") ? "middle" : null);
        if (wantH && H[wantH] !== undefined) {
          try { tf.HorizontalAlignment = H[wantH]; } catch (e) { warnings.push(`设置水平对齐失败: ${e.message}`); }
        }
        if (wantV && V[wantV] !== undefined) {
          try { tf.VerticalAlignment = V[wantV]; } catch (e) { warnings.push(`设置垂直对齐失败: ${e.message}`); }
        }
        for (const [key, val] of Object.entries({ MarginLeft: marginLeft, MarginRight: marginRight, MarginTop: marginTop, MarginBottom: marginBottom })) {
          if (Number.isFinite(Number(val))) { try { tf[key] = Number(val); } catch (e) { warnings.push(`设置 ${key} 失败: ${e.message}`); } }
        }
      } catch (e) { warnings.push(`写入文字失败: ${e.message}`); }
    }

    const actual = readShape(shape, sheet);

    // 核对请求与读回：不一致就如实告警（不抛错，因为部分属性宿主可能合法地做了归一）
    const requested = { kind: k, left: Number(left), top: Number(top), width: Number(width), height: Number(height) };
    if (k === "geometric" || k === "autoshape") {
      const tol = 1.5;
      for (const key of ["left", "top", "width", "height"]) {
        if (actual[key] !== null && Math.abs(Number(actual[key]) - requested[key]) > tol) {
          warnings.push(`${key} 请求 ${requested[key]} 读回 ${actual[key]}`);
        }
      }
    }
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      kind: k,
      shape: actual,
      requested,
      shapeCount: shapes.Count,
      warnings,
      message: `已在 [${sheet.Name}] 创建 ${k} 形状「${actual.name || name || ""}」（当前共 ${shapes.Count} 个形状）`
    };
  }

  /** 读回工作表上的全部形状（绘图能力的验收入口）。 */
  function listShapes(app, params) {
    const { sheetName, workbookName, detail = true, filterName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shapes = sheet.Shapes;
    const items = [];
    for (let i = 1; i <= shapes.Count; i++) {
      const sh = shapes.Item(i);
      if (filterName) {
        let nm = null;
        try { nm = String(sh.Name); } catch (e) {}
        if (nm !== String(filterName)) continue;
      }
      items.push(detail ? readShape(sh, sheet) : { name: (() => { try { return String(sh.Name); } catch (e) { return null; } })() });
    }
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      count: items.length,
      shapes: items,
      message: `工作表 [${sheet.Name}] 上共有 ${items.length} 个形状`
    };
  }

  /** 修改形状：位置/尺寸/旋转/填充/线条/文字/可见性，或删除。 */
  function updateShape(app, params) {
    const { sheetName, workbookName, action, left, top, width, height, rotation,
            fillColor, lineColor, lineWeight, text, visible, newName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shape = findShape(sheet, params);
    const warnings = [];

    const act = String(action || "update").toLowerCase();
    if (act === "delete") {
      const nm = (() => { try { return String(shape.Name); } catch (e) { return null; } })();
      shape.Delete();
      return {
        success: true, workbookName: sheet.Parent.Name, sheetName: sheet.Name,
        action: "delete", deletedName: nm, shapeCountAfter: sheet.Shapes.Count,
        message: `已删除形状「${nm}」（剩余 ${sheet.Shapes.Count} 个）`
      };
    }

    if (Number.isFinite(Number(left))) shape.Left = Number(left);
    if (Number.isFinite(Number(top))) shape.Top = Number(top);
    if (Number.isFinite(Number(width))) shape.Width = Number(width);
    if (Number.isFinite(Number(height))) shape.Height = Number(height);
    if (Number.isFinite(Number(rotation))) shape.Rotation = Number(rotation);
    if (newName) { try { shape.Name = String(newName); } catch (e) { warnings.push(`改名失败: ${e.message}`); } }
    if (visible !== undefined) { try { shape.Visible = visible ? -1 : 0; } catch (e) { warnings.push(`设置可见性失败: ${e.message}`); } }
    if (fillColor) {
      const rgb = hexToExcelColor(fillColor);
      if (rgb === null) warnings.push(`fillColor 无法解析: ${fillColor}`);
      else { try { shape.Fill.Visible = -1; shape.Fill.ForeColor.RGB = rgb; } catch (e) { warnings.push(`设置填充失败: ${e.message}`); } }
    }
    if (lineColor) {
      const rgb = hexToExcelColor(lineColor);
      if (rgb === null) warnings.push(`lineColor 无法解析: ${lineColor}`);
      else { try { shape.Line.Visible = -1; shape.Line.ForeColor.RGB = rgb; } catch (e) { warnings.push(`设置线条色失败: ${e.message}`); } }
    }
    if (Number.isFinite(Number(lineWeight))) {
      try { shape.Line.Weight = Number(lineWeight); } catch (e) { warnings.push(`设置线宽失败: ${e.message}`); }
    }
    if (text !== undefined && text !== null) {
      try { shape.TextFrame.Characters().Text = String(text); } catch (e) { warnings.push(`写入文字失败: ${e.message}`); }
    }

    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      action: "update",
      shape: readShape(shape, sheet),
      warnings,
      message: `已更新形状「${(() => { try { return String(shape.Name); } catch (e) { return ""; } })()}」`
    };
  }

  /** 分组：`Shapes.Range([名...]).Group()`（WPS 表格可用；Office.js 侧未打通）。 */
  function groupShapes(app, params) {
    const { sheetName, workbookName, names, groupName } = params || {};
    const list = Array.isArray(names) ? names.filter(Boolean).map(String) : [];
    if (list.length < 2) throw new Error("group_shapes 需要至少 2 个形状（names 传形状名数组）");
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shapes = sheet.Shapes;
    // 先确认都存在，错误信息才有用
    const existing = [];
    for (let i = 1; i <= shapes.Count; i++) { try { existing.push(String(shapes.Item(i).Name)); } catch (e) {} }
    const missing = list.filter(n => existing.indexOf(n) < 0);
    if (missing.length) throw new Error(`这些形状不存在: ${missing.join(", ")}。现有形状：${existing.join(" / ") || "无"}`);

    const group = shapes.Range(list).Group();
    if (groupName) { try { group.Name = String(groupName); } catch (e) {} }
    const members = [];
    try {
      const gi = group.GroupItems;
      for (let i = 1; i <= gi.Count; i++) { try { members.push(String(gi.Item(i).Name)); } catch (e) {} }
    } catch (e) {}
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      group: readShape(group, sheet),
      requestedMembers: list,
      memberCount: members.length,
      members,
      verified: members.length === list.length,
      warnings: members.length === list.length ? [] : [`请求组合 ${list.length} 个，读回组内 ${members.length} 个`],
      message: `已把 ${list.length} 个形状组合为「${(() => { try { return String(group.Name); } catch (e) { return groupName || ""; } })()}」`
    };
  }

  /** 解散分组，返回释放后的形状名单。 */
  function ungroupShapes(app, params) {
    const { sheetName, workbookName } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shape = findShape(sheet, params);
    const type = (() => { try { return Number(shape.Type); } catch (e) { return null; } })();
    if (type !== 6) {
      throw new Error(`形状「${(() => { try { return String(shape.Name); } catch (e) { return ""; } })()}」不是组合（Type=${type}，组合应为 6）`);
    }
    const before = [];
    try { const gi = shape.GroupItems; for (let i = 1; i <= gi.Count; i++) { try { before.push(String(gi.Item(i).Name)); } catch (e) {} } } catch (e) {}
    shape.Ungroup();
    const after = [];
    const shapes = sheet.Shapes;
    for (let i = 1; i <= shapes.Count; i++) { try { after.push(String(shapes.Item(i).Name)); } catch (e) {} }
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      released: before,
      releasedStillPresent: before.filter(n => after.indexOf(n) >= 0),
      shapeCountAfter: shapes.Count,
      message: `已解散组合，释放 ${before.length} 个形状`
    };
  }

  /** 调整层级。VBA ZOrder 常量：0=置顶 1=置底 2=上移一层 3=下移一层。 */
  function setShapeZorder(app, params) {
    const { sheetName, workbookName, zOrder = "bringToFront" } = params || {};
    const Z = { bringtofront: 0, sendtoback: 1, bringforward: 2, sendbackward: 3 };
    const code = Z[String(zOrder).toLowerCase()];
    if (code === undefined) throw new Error(`不支持的 zOrder: ${zOrder}（可用 bringToFront / sendToBack / bringForward / sendBackward）`);
    const sheet = getWorksheet(app, sheetName, workbookName);
    const shape = findShape(sheet, params);
    shape.ZOrder(code);
    // 读回全部形状以核对层级（WPS 上 ZOrderPosition 不可读，改用"绘制顺序"近似：
    // 按 ZOrder(0) 逐个置顶来还原顺序代价太高，这里只如实返回当前形状与形状总数）
    const names = [];
    const shapes = sheet.Shapes;
    for (let i = 1; i <= shapes.Count; i++) { try { names.push(String(shapes.Item(i).Name)); } catch (e) {} }
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      applied: zOrder,
      shape: readShape(shape, sheet),
      shapeNames: names,
      warnings: ["WPS 表格读不到 ZOrderPosition，无法直接回读层级；shapeNames 的顺序是 Shapes 集合顺序，可作为近似参考"],
      message: `已把形状「${(() => { try { return String(shape.Name); } catch (e) { return ""; } })()}」调整为 ${zOrder}`
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

    // 单元格地址取值。**必须保留接收者直接调用**：宿主方法是原生实现，
    // 摘出来再调（`const f = cell.Address; f()`）会丢 this，抛
    // "Address called on null or undefined"，被 catch 吞掉后表现为地址为空
    // （真机实测：消息变成"已在单元格  添加批注"，ISS-66 的根因就在这里）。
    const addressOf = (cell) => {
      try {
        return String(cell.Address());
      } catch (e) {
        try {
          // 少数宿主把 Address 暴露成属性而非方法
          return String(cell.Address);
        } catch (e2) {
          return "";
        }
      }
    };

    if (action === "add") {
      if (!text) throw new Error("添加批注时必须提供 text 参数");
      const cell = targetRange.Cells.Item(1, 1);
      // 地址必须在 AddComment **之前**取：真机实测 AddComment 之后该单元格的 Address 取不到，
      // 会让成功消息变成"已在单元格  添加批注"（地址为空）。
      const cellAddressText = addressOf(cell); // 先取地址，AddComment 后取不到
      try {
        if (cell.Comment) cell.Comment.Delete();
      } catch (e) {}
      // 不再把作者拼进正文（ISS-43）：先尝试设置真实 Author，设不上再退回"正文前缀"，
      // 并在返回体里**如实区分**"请求的作者"与"读回的作者"（原来直接回显入参，冒充读回，ISS-61）。
      const comment = cell.AddComment(text);
      comment.Visible = false;
      let authorApplied = false;
      if (author) {
        try {
          comment.Author = author;
          authorApplied = true;
        } catch (e) {
          authorApplied = false;
        }
      }
      let authorOnHost = null;
      try { authorOnHost = comment.Author || null; } catch (e) {}
      if (author && !authorApplied && authorOnHost !== author) {
        try {
          comment.Text(text ? `${author}:\n${text}` : text);
        } catch (e) {}
      }
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        address: cellAddressText,
        action: "add",
        commentText: text,
        authorRequested: author || null,
        authorOnHost,
        authorApplied: authorOnHost === author,
        message: `已在单元格 ${cellAddressText} 添加批注${author && authorOnHost !== author ? `（宿主未接受自定义作者，已把作者写进正文；当前宿主作者为 ${authorOnHost ?? "未知"}）` : ""}`
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
                address: addressOf(cell),
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

  // 13.9 工作表视图：网格线 / 行列标题 / 缩放
  //
  // 为什么必须是个独立能力：用矢量形状搭画布页时**必须先隐藏网格线**，
  // 否则形状浮在格线上、观感很乱。此前没有任何工具入口，
  // 只能让 AI 退回 wps_execute_script 手写（真机踩到过）。
  function setSheetView(app, params) {
    const { sheetName, workbookName, showGridlines, showHeadings, zoom } = params || {};
    const sheet = getWorksheet(app, sheetName, workbookName);
    try { sheet.Activate(); } catch (e) {}
    const win = app.ActiveWindow;
    const before = {
      showGridlines: (() => { try { return Boolean(win.DisplayGridlines); } catch (e) { return null; } })(),
      showHeadings: (() => { try { return Boolean(win.DisplayHeadings); } catch (e) { return null; } })(),
      zoom: (() => { try { return Number(win.Zoom); } catch (e) { return null; } })()
    };
    const warnings = [];
    if (showGridlines !== undefined) {
      try { win.DisplayGridlines = showGridlines ? true : false; } catch (e) { warnings.push(`设置网格线失败: ${e.message}`); }
    }
    if (showHeadings !== undefined) {
      try { win.DisplayHeadings = showHeadings ? true : false; } catch (e) { warnings.push(`设置行列标题失败: ${e.message}`); }
    }
    if (Number.isFinite(Number(zoom))) {
      try { win.Zoom = Number(zoom); } catch (e) { warnings.push(`设置缩放失败: ${e.message}`); }
    }
    // 读回核对：写没写进去必须能验证
    const after = {
      showGridlines: (() => { try { return Boolean(win.DisplayGridlines); } catch (e) { return null; } })(),
      showHeadings: (() => { try { return Boolean(win.DisplayHeadings); } catch (e) { return null; } })(),
      zoom: (() => { try { return Number(win.Zoom); } catch (e) { return null; } })()
    };
    return {
      success: true,
      workbookName: sheet.Parent.Name,
      sheetName: sheet.Name,
      before, after, warnings,
      message: `工作表 [${sheet.Name}] 视图：网格线 ${after.showGridlines ? "显示" : "隐藏"}、行列标题 ${after.showHeadings ? "显示" : "隐藏"}、缩放 ${after.zoom}%`
    };
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
      // 默认范围必须是"**整张画布**"，不能只用 UsedRange——
      // UsedRange 只统计**有数据的单元格，不含形状**。当一页全是矢量图形、
      // 单元格没有任何值时，UsedRange 会退化成 $A$1，截出来只有 146x50 的废图
      //（真机踩到过）。所以这里把"所有形状的外框"并入 UsedRange。
      const ur = sheet.UsedRange;
      let r1 = ur.Row, c1 = ur.Column;
      let r2 = r1 + ur.Rows.Count - 1, c2 = c1 + ur.Columns.Count - 1;
      try {
        const shapes = sheet.Shapes;
        for (let i = 1; i <= shapes.Count; i++) {
          const sh = shapes.Item(i);
          let tl = null, br = null;
          try { tl = sh.TopLeftCell; br = sh.BottomRightCell; } catch (e) {}
          if (!tl || !br) continue;
          r1 = Math.min(r1, tl.Row); c1 = Math.min(c1, tl.Column);
          r2 = Math.max(r2, br.Row); c2 = Math.max(c2, br.Column);
        }
      } catch (e) {}
      // 注意：WPS JSA 的 Worksheet **没有 `Cells()` 方法**（Excel VBA 有），
      // 调它会报 "sheet.Cells is not a function"。这里自己拼 A1 地址。
      const colName = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
      targetRange = sheet.Range(`${colName(c1)}${r1}:${colName(c2)}${r2}`);
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
      replaceExisting = true,
      left: explicitLeft,
      top: explicitTop,
      width: explicitWidth,
      height: explicitHeight,
      startCell,
      endCell,
      cellRange
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

    /** 本次建图的告警集合（类型降级、单系列多色等），随返回体带出。 */
    const warnings = [];

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

    // 原实现只认 `position.leftCell/width/height`，而 startCell / cellRange / endCell 与
    // 顶层 left/top/width/height **全被静默忽略**，多图会叠在默认的 360/40（问题台账 ISS-18）。
    // 这里把它们真正接上，并让"锚点"统一走一套解析顺序：position.leftCell > startCell/cellRange > 顶层像素。
    const anchorRef = (position && position.leftCell)
      ? String(position.leftCell)
      : (startCell ? String(startCell) : (cellRange ? String(cellRange).split(":")[0] : null));
    if (anchorRef && !(position && position.leftCell)) {
      try {
        const anchorRange = sheet.Range(anchorRef.includes(":") ? anchorRef.split(":")[0] : anchorRef);
        left = anchorRange.Left;
        top = anchorRange.Top;
      } catch (e) {
        log("图表定位锚点警告: " + e.message);
      }
    }
    // 顶层像素参数直接生效（显式传入优先于上面的锚点推算）
    if (Number.isFinite(Number(explicitLeft))) left = Number(explicitLeft);
    if (Number.isFinite(Number(explicitTop))) top = Number(explicitTop);
    if (Number.isFinite(Number(explicitWidth))) width = Number(explicitWidth);
    if (Number.isFinite(Number(explicitHeight))) height = Number(explicitHeight);
    // endCell：用"锚点单元格 → endCell"的矩形尺寸作为图表宽高
    if (endCell) {
      try {
        const from = sheet.Range(anchorRef ? anchorRef.split(":")[0] : "A1");
        const to = sheet.Range(String(endCell));
        const w = to.Left + to.Width - from.Left;
        const h = to.Top + to.Height - from.Top;
        if (w > 0) width = w;
        if (h > 0) height = h;
      } catch (e) {
        log("图表 endCell 尺寸推算警告: " + e.message);
      }
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
          // 单系列图表上传多个颜色，宿主会按"逐点染色"处理 → 应单色的柱图/条形图变成彩虹柱
          // （问题台账 ISS-24）。这里只应用第一个颜色，并把"其余被忽略"如实告知。
          if (seriesCount <= 1 && seriesColors.length > 1) {
            warnings.push(
              `seriesColors 传了 ${seriesColors.length} 个颜色，但该图只有 ${seriesCount || 1} 个数据系列；` +
              `已只应用第 1 个（宿主对单系列多色的处理是逐点染色，会出现彩虹柱）。`
            );
          }
          const colorLimit = seriesCount <= 1 ? Math.min(1, seriesColors.length) : seriesColors.length;
          seriesColors.slice(0, colorLimit).forEach((colorHex, idx) => {
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
      if (actualChartType !== null && Number.isFinite(actualChartType) && actualChartType !== xlChartType) {
        warnings.push(
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
        warnings,
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
    // ChartType 数字 → 可读枚举名（问题台账 ISS-23：detail 原来只给数字，调用方没法判类型）
    const CHART_TYPE_NAMES = {
      1: "xlArea", 4: "xlLine", 5: "xlPie", 51: "xlColumnClustered", 52: "xlColumnStacked",
      57: "xlBarClustered", 58: "xlBarStacked", 65: "xlLineMarkers", 68: "xlBarOfPie",
      122: "xlPareto", "-4120": "xlDoughnut", "-4169": "xlXYScatter"
    };

    for (let i = 1; i <= sheet.Shapes.Count; i++) {
      const shape = sheet.Shapes.Item(i);
      if (!shape.HasChart) continue;
      ordinal++;
      const chart = shape.Chart;
      const title = safeRead(() => chart.HasTitle && chart.ChartTitle ? chart.ChartTitle.Text : "", "");
      if (shapeName && shape.Name !== shapeName) continue;
      if (chartIndex && ordinal !== Number(chartIndex)) continue;
      if (chartTitle && String(title).indexOf(chartTitle) < 0) continue;

      const position = { left: safeRead(() => shape.Left, null), top: safeRead(() => shape.Top, null) };
      const item = {
        chartIndex: ordinal,
        shapeName: shape.Name,
        title,
        chartType: safeRead(() => chart.ChartType, null),
        chartTypeName: CHART_TYPE_NAMES[safeRead(() => chart.ChartType, null)] || null,
        hasLegend: !!safeRead(() => chart.HasLegend, false),
        left: position.left,
        top: position.top,
        width: safeRead(() => shape.Width, null),
        height: safeRead(() => shape.Height, null),
        // leftCell 是"形状左上角所在单元格"，**不是唯一锚点**：多张图叠在同一像素位时会报同一个值
        // （问题台账 ISS-23）。额外给出 isDefaultPosition，避免被当成唯一定位依据。
        leftCell: safeRead(() => shape.TopLeftCell.Address(), null),
        isDefaultPosition: Math.abs(Number(position.left) - 360) < 2 && Math.abs(Number(position.top) - 40) < 2,
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
    // CAP-36 透视表读回：此前只能建、不能读，AI 无法回答
    // "这张表上现有哪些透视表、数据源是哪、字段怎么摆、刷新过没有"。
    if ((params || {}).action === "read") {
      const { workbookName, sheetName, pivotTableName } = params || {};
      const wb = getWorkbook(app, workbookName);
      const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
      const readOne = (pt) => {
        const fields = [];
        try {
          const raw = pt.PivotFields();
          const n = Number(raw.Count);
          for (let i = 1; i <= n; i++) {
            const f = raw.Item(i);
            fields.push({
              name: g(() => String(f.Name), null),
              orientation: g(() => Number(f.Orientation), null),   // 1=行 2=列 4=页 0=隐藏
              position: g(() => Number(f.Position), null),
              function: g(() => Number(f.Function), null),
              subtotals: g(() => Boolean(f.Subtotals), null)
            });
          }
        } catch (e) {}
        return {
          name: g(() => String(pt.Name), null),
          sourceData: g(() => String(pt.SourceData), null),
          rowRange: g(() => String(pt.RowRange.Address()), null),
          tableRange1: g(() => String(pt.TableRange1.Address()), null),
          refreshDate: g(() => (pt.RefreshDate ? new Date(pt.RefreshDate).toISOString() : null), null),
          version: g(() => String(pt.Version), null),
          fieldCount: fields.length,
          fields
        };
      };
      const out = [];
      const sheets = sheetName ? [sheetName] : (() => {
        const names = [];
        for (let i = 1; i <= wb.Worksheets.Count; i++) { try { names.push(String(wb.Worksheets.Item(i).Name)); } catch (e) {} }
        return names;
      })();
      for (const sn of sheets) {
        let pts = null;
        try { pts = wb.Worksheets.Item(sn).PivotTables(); } catch (e) { continue; }
        const n = g(() => Number(pts.Count), 0);
        for (let i = 1; i <= n; i++) {
          const pt = (() => { try { return pts.Item(i); } catch (e) { return null; } })();
          if (!pt) continue;
          const one = readOne(pt);
          if (pivotTableName && one.name !== String(pivotTableName)) continue;
          out.push({ sheetName: sn, ...one });
        }
      }
      return {
        success: true, workbookName: wb.Name,
        count: out.length, pivotTables: out, warnings: [],
        message: `工作簿 [${wb.Name}] 共 ${out.length} 个数据透视表`
      };
    }
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
    // 防御性激活。真机探到 Create(1, srcRange) 在活动表上返回对象、非活动表上返回 null，
    // 随后 .CreatePivotTable 报错；同时源区域必须**有数据**，空区域同样建不出透视表
    // （ISS-118 复核结论：双因，不是单一原因）。
    try { srcSheet.Activate(); } catch (e) {}
    const srcRange = srcSheet.Range(sourceRange);

    // 目标表不存在时自动新建（原来必须由调用方先建好，问题台账 ISS-63 的第二个坑）
    let destSheet = null;
    try {
      destSheet = getWorksheet(app, destSheetName, workbookName);
    } catch (e) {
      if (!destSheetName) throw e;
      destSheet = wb.Worksheets.Add();
      try { destSheet.Name = String(destSheetName); } catch (nameErr) {
        throw new Error(`目标工作表 "${destSheetName}" 不存在，自动新建后重命名失败（可能重名或含非法字符）：${nameErr.message}`);
      }
    }
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

    // 新建后是"空骨架"，字段配好也必须显式刷新才会真正取数（问题台账 ISS-63：
    // 调用方拿到 success 却看到一张空表，得自己去手动 Refresh）。
    let refreshedBy = null;
    try {
      pivotTable.RefreshTable();
      refreshedBy = "pivotTable.RefreshTable";
    } catch (e) {
      try {
        pivotCache.Refresh();
        refreshedBy = "pivotCache.Refresh";
      } catch (e2) {
        log("透视表刷新失败: " + e.message + " / " + e2.message);
      }
    }

    // 读回实际结果：记录数与表区域是"真的取到数"的唯一证据
    let recordCount = null;
    let tableRange = null;
    try { recordCount = Number(pivotTable.RecordCount); } catch (e) {}
    try { tableRange = pivotTable.TableRange1 ? pivotTable.TableRange1.Address() : null; } catch (e) {}
    const warnings = [];
    if (!refreshedBy) warnings.push("透视表刷新调用失败，可能是空骨架，请在 WPS 里右键手动刷新后再读回确认。");
    if (recordCount === 0) warnings.push("刷新后 RecordCount 仍为 0：源区域可能没有可用数据，或字段未正确落位。");

    return {
      success: true,
      workbookName: wb.Name,
      pivotTableName: ptName,
      destSheetName: destSheet.Name,
      destCell,
      rowCount: rowFields.length,
      colCount: columnFields.length,
      dataCount: dataFields.length,
      refreshedBy,
      recordCount,
      tableRange,
      warnings,
      message: `已成功在 [${destSheet.Name}] ${destCell} 生成数据透视表${refreshedBy ? "并完成刷新" : "（刷新未成功，见 warnings）"}`
    };
  }

  // 17. 自动筛选与数据排序 (第三梯队)
  function setFilterAndSort(app, params) {
    const { sheetName, workbookName, range, enableAutoFilter, sortRules, action = "apply" } = params || {};

    const sheet = getWorksheet(app, sheetName, workbookName);

    // CAP-33 筛选状态读回：不传 range 也能读（读的是"这张表当前有没有筛选、覆盖哪一块"）。
    if (action === "read") {
      const active = (() => { try { return Boolean(sheet.AutoFilterMode); } catch (e) { return null; } })();
      const addr = (() => {
        try { return sheet.AutoFilterMode && sheet.AutoFilter && sheet.AutoFilter.Range ? sheet.AutoFilter.Range.Address() : null; } catch (e) { return null; }
      })();
      // 逐个字段的筛选条件（Criteria1/Criteria2/Operator）
      const filters = [];
      try {
        if (active && addr) {
          const f = sheet.AutoFilter;
          const rng = f.Range;
          const cols = rng.Columns.Count;
          for (let i = 1; i <= cols; i++) {
            const flt = f.Filters.Item(i);
            let on = false;
            try { on = Boolean(flt.On); } catch (e) {}
            if (!on) continue;
            const one = (() => { try { return flt.Criteria1 === undefined ? null : flt.Criteria1; } catch (e) { return null; } })();
            const two = (() => { try { return flt.Criteria2 === undefined ? null : flt.Criteria2; } catch (e) { return null; } })();
            const op = (() => { try { return Number(flt.Operator); } catch (e) { return null; } })();
            const header = (() => { try { return String(rng.Cells(1, i).Value2); } catch (e) { return null; } })();
            filters.push({ columnIndex: i, header, criteria1: one, criteria2: two, operator: op });
          }
        }
      } catch (e) {}
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        autoFilterOn: active,
        filterRange: addr,
        columnCount: addr ? sheet.Range(addr).Columns.Count : 0,
        filters,
        warnings: [],
        message: active
          ? `工作表 [${sheet.Name}] 筛选范围 ${addr}，其中 ${filters.length} 列设有条件`
          : `工作表 [${sheet.Name}] 当前没有启用筛选`
      };
    }

    if (!range) throw new Error("缺少必要参数: range (例如 'A4:E20')");

    // 防御性激活：WPS 上部分 Range 级操作（AutoFilter / Sort）对非活动表可能静默 no-op。
    // 说明：曾把「筛选不生效」记成产品缺陷（ISS-117），复核后确认是**验证脚本没检查写入结果**——
    // 数据根本没写进表，空区域上 AutoFilter() 自然无效。激活保留作防御，但不是该问题的根因。
    try { sheet.Activate(); } catch (e) {}

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

    const sheet = getWorksheet(app, sheetName, workbookName);

    // CAP-32 数据有效性读回：读整张表**所有**带校验的区域，或指定 address 的那一块。
    // 只写不读时 AI 无法回答"这列现在允许哪些值"，复核只能靠人看。
    if ((params || {}).action === "read") {
      const VALTYPE = { 1: "whole_number", 2: "decimal", 3: "list", 4: "date", 5: "time", 6: "text_length", 7: "custom" };
      const OP = { 1: "between", 2: "not_between", 3: "equal", 4: "not_equal", 5: "greater_than", 6: "less_than", 7: "greater_equal", 8: "less_equal" };
      const readOne = (rng) => {
        const v = rng.Validation;
        let type = null;
        try { type = Number(v.Type); } catch (e) { return null; }
        // Type = -4142 (xlValidateInputOnly) 表示没设校验，跳过
        if (type === -4142 || !Number.isFinite(type)) return null;
        const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
        return {
          address: g(() => rng.Address(), null),
          type: VALTYPE[type] || `unknown(${type})`,
          typeCode: type,
          operator: OP[g(() => Number(v.Operator), null)] || null,
          formula1: g(() => v.Formula1, null),
          formula2: g(() => v.Formula2, null),
          inCellDropdown: g(() => Boolean(v.InCellDropdown), null),
          ignoreBlank: g(() => Boolean(v.IgnoreBlank), null),
          showError: g(() => Boolean(v.ShowError), null),
          errorTitle: g(() => v.ErrorTitle, null),
          errorMessage: g(() => v.ErrorMessage, null),
          promptTitle: g(() => v.InputTitle, null),
          promptMessage: g(() => v.InputMessage, null)
        };
      };
      let items = [];
      if (address) {
        const one = readOne(sheet.Range(address));
        items = one ? [one] : [];
      } else {
        // 逐行扫描已用区域（宿主没有"枚举所有校验区域"的 API）
        const ur = sheet.UsedRange;
        const r0 = ur.Row, c0 = ur.Column;
        const rows = Math.min(ur.Rows.Count, 2000), cols = Math.min(ur.Columns.Count, 100);
        const colName = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
        for (let r = r0; r < r0 + rows; r++) {
          for (let c = c0; c < c0 + cols; c++) {
            const one = readOne(sheet.Range(`${colName(c)}${r}`));
            if (one) items.push(one);
          }
        }
      }
      return {
        success: true,
        workbookName: sheet.Parent.Name,
        sheetName: sheet.Name,
        count: items.length,
        validations: items,
        warnings: address ? [] : [],
        message: `工作表 [${sheet.Name}] 上共 ${items.length} 处数据有效性设置`
      };
    }

    if (!address) throw new Error("缺少必要参数: address (例如 'E5:E20')");

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
    if (!action) throw new Error("缺少必要参数: action (rename, move, tab_color, protect, unprotect, read)");

    const wb = getWorkbook(app, workbookName);
    const sheet = getWorksheet(app, sheetName, workbookName);

    // CAP-34 工作表状态读回：保护状态与标签色此前只能写不能读，
    // 于是"这张表是不是被保护了""标签什么颜色"只能靠人看。
    if (action === "read") {
      const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
      const prot = g(() => Boolean(sheet.ProtectContents), null);
      const colorRaw = g(() => Number(sheet.Tab.Color), null);
      // Tab.Color 是 BGR 整数；-4142/16777215 之类表示"无颜色"
      const hasColor = colorRaw !== null && colorRaw >= 0 && colorRaw !== 16777215;
      const toHex = (v) => "#" + [v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff].map(n => n.toString(16).padStart(2, "0")).join("").toUpperCase();
      return {
        success: true,
        workbookName: wb.Name,
        sheetName: sheet.Name,
        index: g(() => Number(sheet.Index), null),
        visible: g(() => Number(sheet.Visible), null),
        protection: {
          protectContents: prot,
          protectDrawingObjects: g(() => Boolean(sheet.ProtectDrawingObjects), null),
          protectionMode: g(() => Boolean(sheet.ProtectionMode), null)
        },
        tabColor: hasColor ? toHex(colorRaw) : null,
        tabColorRaw: colorRaw,
        warnings: [],
        message: `工作表 [${sheet.Name}]：${prot ? "已保护" : "未保护"}，标签色 ${hasColor ? toHex(colorRaw) : "未设置（默认）"}`
      };
    }

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

  // ---------------------------------------------------------------------------
  // Word 只读读回（ISS-44 / ISS-73）：节与页面版式、水印形状、页码域、书签、内容控件、
  // 文档属性、样式清单。全部按"逐项 try/catch + 结果里带 error 字段"实现：
  // 单项读不到不影响整次读回，也不会抛宿主内部错误给调用方。
  // ---------------------------------------------------------------------------
  function wordReadPageSetup(doc) {
    const readSetup = (ps) => ({
      pageWidth: Number(ps.PageWidth),
      pageHeight: Number(ps.PageHeight),
      orientation: Number(ps.Orientation),
      topMargin: Number(ps.TopMargin),
      bottomMargin: Number(ps.BottomMargin),
      leftMargin: Number(ps.LeftMargin),
      rightMargin: Number(ps.RightMargin),
      headerDistance: Number(ps.HeaderDistance),
      footerDistance: Number(ps.FooterDistance),
      differentFirstPage: Number(ps.DifferentFirstPageHeaderFooter) !== 0,
      differentOddEvenPages: Number(ps.OddAndEvenPagesHeaderFooter) !== 0
    });
    const sections = [];
    const count = doc.Sections.Count;
    for (let i = 1; i <= count; i++) {
      const entry = { index: i };
      try {
        const s = doc.Sections.Item(i);
        entry.pageSetup = readSetup(s.PageSetup);
        entry.range = { start: s.Range.Start, end: s.Range.End };
        entry.header = (() => {
          try {
            const h = s.Headers.Item(1);
            return {
              text: (h.Range.Text || "").replace(/[\r\n\x07]/g, ""),
              fieldCount: (() => { try { return h.Range.Fields.Count; } catch (e) { return null; } })(),
              shapeCount: (() => { try { return h.Shapes.Count; } catch (e) { return null; } })(),
              linkToPrevious: (() => { try { return h.LinkToPrevious; } catch (e) { return null; } })()
            };
          } catch (e) { return { error: e.message }; }
        })();
        entry.footer = (() => {
          try {
            const f = s.Footers.Item(1);
            const fields = [];
            try {
              for (let k = 1; k <= f.Range.Fields.Count; k++) {
                const fld = f.Range.Fields.Item(k);
                fields.push({ type: Number(fld.Type), code: (fld.Code ? fld.Code.Text : "").trim() });
              }
            } catch (e) {}
            return {
              text: (f.Range.Text || "").replace(/[\r\n\x07]/g, ""),
              fieldCount: (() => { try { return f.Range.Fields.Count; } catch (e) { return null; } })(),
              fields: fields,
              pageNumberCount: (() => { try { return f.PageNumbers.Count; } catch (e) { return null; } })(),
              linkToPrevious: (() => { try { return f.LinkToPrevious; } catch (e) { return null; } })()
            };
          } catch (e) { return { error: e.message }; }
        })();
        // 该节内（或未报到节归属）的水印形状（正文层 WordArt）；按 Name 以 WordArt 前缀识别。
        // 逐个读 Anchor.Start 在超多形状文档上开销不小，限定最多扫 80 个并如实标注是否截断。
        entry.watermarkShapes = (() => {
          try {
            const list = [];
            const s0 = s.Range.Start, e0 = s.Range.End;
            const totalShapes = doc.Shapes.Count;
            const limit = Math.min(totalShapes, 80);
            for (let k = 1; k <= limit; k++) {
              const shp = doc.Shapes.Item(k);
              let anchorStart = null;
              try { anchorStart = shp.Anchor.Start; } catch (e) { anchorStart = null; }
              if (anchorStart === null || (anchorStart >= s0 && anchorStart <= e0)) {
                list.push({
                  name: shp.Name,
                  type: Number(shp.Type),
                  text: (() => { try { return shp.TextEffect ? shp.TextEffect.Text : undefined; } catch (e) { return undefined; } })(),
                  left: Number(shp.Left),
                  top: Number(shp.Top),
                  anchoredInSection: anchorStart !== null && anchorStart >= s0 && anchorStart <= e0
                });
              }
            }
            return { totalShapes: totalShapes, scannedShapes: limit, truncated: totalShapes > limit, shapes: list };
          } catch (e) { return { error: e.message }; }
        })();
      } catch (e) {
        entry.error = e.message;
      }
      sections.push(entry);
    }
    // 页脚页码格式：从域码推断（PAGE / NUMPAGES 组合）
    for (const entry of sections) {
      if (!entry.footer || entry.footer.error) continue;
      const codes = (entry.footer.fields || []).map(f => f.code);
      const hasPage = codes.some(c => /^PAGE\b/.test(c));
      const hasNumPages = codes.some(c => /^NUMPAGES\b/.test(c));
      entry.pageNumberFormat = hasPage && hasNumPages ? "page_of_pages" : (hasPage ? (codes.some(c => /PAGE/.test(c)) && /-\s*$|^-\s/.test(entry.footer.text) ? "dash" : "simple") : null);
    }
    return { document: readSetup(doc.PageSetup), sections: sections };
  }

  function wordReadBookmarks(doc) {
    const list = [];
    try {
      for (let i = 1; i <= doc.Bookmarks.Count; i++) {
        const b = doc.Bookmarks.Item(i);
        list.push({
          name: b.Name,
          start: b.Range.Start,
          end: b.Range.End,
          empty: (() => { try { return Boolean(b.Empty); } catch (e) { return undefined; } })(),
          text: (b.Range.Text || "").replace(/[\r\n\x07]/g, "")
        });
      }
    } catch (e) {
      return { count: 0, list: list, error: e.message };
    }
    return { count: list.length, list: list };
  }

  function wordReadFields(doc) {
    const list = [];
    try {
      const limit = Math.min(doc.Fields.Count, 100);
      for (let i = 1; i <= limit; i++) {
        const f = doc.Fields.Item(i);
        list.push({
          index: i,
          type: Number(f.Type),
          code: (f.Code ? f.Code.Text : "").trim(),
          result: (f.Result ? f.Result.Text : "").replace(/[\r\n\x07]/g, "")
        });
      }
    } catch (e) {
      return { count: 0, list: list, error: e.message };
    }
    return { count: doc.Fields.Count, list: list };
  }

  function wordReadContentControls(doc) {
    const list = [];
    try {
      const count = doc.ContentControls.Count;
      for (let i = 1; i <= count; i++) {
        const c = doc.ContentControls.Item(i);
        list.push({
          index: i,
          title: (() => { try { return c.Title; } catch (e) { return undefined; } })(),
          tag: (() => { try { return c.Tag; } catch (e) { return undefined; } })(),
          type: (() => { try { return Number(c.Type); } catch (e) { return undefined; } })(),
          text: (() => { try { return (c.Range.Text || "").replace(/[\r\n\x07]/g, ""); } catch (e) { return undefined; } })()
        });
      }
      return { count: count, list: list };
    } catch (e) {
      return { count: 0, list: list, error: e.message };
    }
  }

  function wordReadStyles(doc, nameFilter) {
    try {
      const total = doc.Styles.Count;
      const kw = nameFilter ? String(nameFilter).toLowerCase() : null;
      const inUse = [];
      const matched = [];
      for (let i = 1; i <= total; i++) {
        const st = doc.Styles.Item(i);
        let name = "";
        let iu = false;
        try { name = st.NameLocal || st.Name || ""; } catch (e) {}
        try { iu = Boolean(st.InUse); } catch (e) {}
        if (iu) inUse.push({ name: name, type: (() => { try { return Number(st.Type); } catch (e) { return undefined; } })(), builtIn: (() => { try { return Boolean(st.BuiltIn); } catch (e) { return undefined; } })() });
        if (kw && name.toLowerCase().indexOf(kw) >= 0 && matched.length < 100) {
          matched.push({ name: name, inUse: iu });
        }
      }
      return { count: total, inUseCount: inUse.length, inUse: inUse.slice(0, 100), matched: kw ? matched : undefined };
    } catch (e) {
      return { count: 0, error: e.message };
    }
  }

  function wordReadDocumentProperties(doc) {
    const readBuiltIn = (name) => {
      try {
        return doc.BuiltInDocumentProperties.Item(name).Value;
      } catch (e) {
        return undefined;
      }
    };
    const custom = [];
    try {
      for (let i = 1; i <= doc.CustomDocumentProperties.Count; i++) {
        const p = doc.CustomDocumentProperties.Item(i);
        custom.push({
          name: (() => { try { return p.Name; } catch (e) { return undefined; } })(),
          value: (() => { try { return p.Value; } catch (e) { return undefined; } })()
        });
      }
    } catch (e) {}
    return {
      title: readBuiltIn("Title"),
      subject: readBuiltIn("Subject"),
      author: readBuiltIn("Author"),
      keywords: readBuiltIn("Keywords"),
      category: readBuiltIn("Category"),
      comments: readBuiltIn("Comments"),
      lastAuthor: readBuiltIn("Last Author"),
      revision: readBuiltIn("Revision Number"),
      custom: custom
    };
  }

  function wordReadDocument(app, params) {
    const { documentName, scope, maxParagraphs, includeFormatting, includeTables } = params || {};
    const doc = getWordDocument(app, documentName);
    const outline = [];
    const maxP = Number(maxParagraphs) || 200;

    // 新增的读回 scope（ISS-44 / ISS-73）：与旧 scope 并存，均为**增量**返回，
    // 不改变 full / outline / paragraphs / tables / selection 的既有语义。
    const EXTRA_SCOPES = { layout: 1, styles: 1, properties: 1, bookmarks: 1, content_controls: 1, fields: 1 };
    const extra = Object.prototype.hasOwnProperty.call(EXTRA_SCOPES, scope);

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

    // 读回域（ISS-44 / ISS-73）：只在显式请求时读，避免每次读文档都枚举几百个样式
    const readBack = {};
    if (extra) {
      if (scope === "layout") {
        readBack.layout = wordReadPageSetup(doc);
      } else if (scope === "styles") {
        readBack.styles = wordReadStyles(doc, params.styleNameFilter);
      } else if (scope === "properties") {
        readBack.properties = wordReadDocumentProperties(doc);
      } else if (scope === "bookmarks") {
        readBack.bookmarks = wordReadBookmarks(doc);
      } else if (scope === "content_controls") {
        readBack.contentControls = wordReadContentControls(doc);
      } else if (scope === "fields") {
        readBack.fields = wordReadFields(doc);
      }
    }

    return {
      documentName: doc.Name,
      fullName: doc.FullName || doc.Name,
      paragraphCount: doc.Paragraphs ? doc.Paragraphs.Count : 0,
      tableCount: doc.Tables ? doc.Tables.Count : 0,
      wordCount: doc.Words ? doc.Words.Count : 0,
      outline,
      tables: tablesSummary,
      paragraphDetails: includeFormatting ? paragraphDetails : undefined,
      previewText: scope === "outline" || extra ? undefined : previewText,
      selectionText: selectionText || undefined,
      ...readBack
    };
  }

  // 找到包含指定字符位置的段落 Range（用于 bookmark / after_paragraph 的精确定位）。
  // 说明：宿主 `Paragraphs.Add(targetRange)` 在 WPS for Mac 上**不可靠**——实测传书签范围时
  // 会把内容写到文首、替换掉书签所在文字并销毁该书签（问题台账 ISS-70）。因此这里改成
  // "定位段落 → InsertParagraphAfter 建段 → InsertAfter 写文本"，全部经真实宿主探针验证。
  function wordFindParagraphRangeAt(doc, pos) {
    const count = doc.Paragraphs.Count;
    for (let i = 1; i <= count; i++) {
      const r = doc.Paragraphs.Item(i).Range;
      if (pos >= r.Start && pos <= r.End) return r;
    }
    return null;
  }

  // 在 anchorRange 所指段落之后写入一个新段落，返回新段落 Range 与命中的定位信息。
  function wordInsertParagraphAfterRange(doc, anchorRange, text) {
    const endPos = anchorRange.End - 1;
    const anchor = doc.Range(endPos, endPos);
    anchor.InsertParagraphAfter();
    anchor.InsertAfter(text);
    return { range: anchor, anchorEnd: endPos };
  }

  // 写入单行文本；位置由 location 决定。返回写入后的 Range 供排版与读回使用。
  function wordWriteOneLine(doc, wordApp, location, targetBookmark, paragraphIndex, text, meta) {
    if (location === "bookmark") {
      const bm = doc.Bookmarks.Item(targetBookmark);
      const bmRange = bm.Range;
      meta.bookmarkRange = { start: bmRange.Start, end: bmRange.End };
      const para = wordFindParagraphRangeAt(doc, bmRange.Start);
      if (!para) throw new Error(`书签 [${targetBookmark}] 所在段落无法定位，未写入；请检查文档结构`);
      // 记录书签锚点所属段落，写完后核对书签仍在（宿主可能在插入过程中丢弃书签）
      const anchorParagraphStart = para.Start;
      const res = wordInsertParagraphAfterRange(doc, para, text);
      meta.bookmarkPreserved = doc.Bookmarks.Exists(targetBookmark);
      meta.insertedAtParagraph = (() => {
        try {
          const after = wordFindParagraphRangeAt(doc, res.anchorEnd);
          return after ? { start: after.Start, end: after.End } : null;
        } catch (e) { return null; }
      })();
      meta.anchorParagraphStartBefore = anchorParagraphStart;
      return res.range;
    }
    if (location === "after_paragraph") {
      const idx = Number(paragraphIndex);
      if (!idx || idx < 1 || idx > doc.Paragraphs.Count) {
        throw new Error(`段落索引越界：第 ${paragraphIndex} 段不存在（当前共 ${doc.Paragraphs.Count} 段）`);
      }
      const para = doc.Paragraphs.Item(idx).Range;
      return wordInsertParagraphAfterRange(doc, para, text).range;
    }
    if (location === "selection" && wordApp && wordApp.Selection && wordApp.Selection.Range) {
      const sel = wordApp.Selection.Range;
      const collapsed = sel.Start === sel.End;
      if (collapsed) {
        const r = doc.Range(sel.Start, sel.Start);
        r.InsertAfter(text);
        return r;
      }
      // 非折叠选区：原位替换所选内容（宿主 Range.Text 赋值不可靠，用 InsertAfter 把新文本插到选区之后，
      // 再删掉被推到后面的原选区文本；宿主探针实测：选区 "乙 " → "乙 替换后 abc"）
      const start = sel.Start;
      sel.InsertAfter(text);
      try { doc.Range(start + text.length, sel.End + text.length).Delete(); } catch (e) {}
      return doc.Range(start, start + text.length);
    }
    if (location === "start") {
      const r = doc.Range(0, 0);
      r.InsertAfter(text);
      return doc.Range(0, Math.min(text.length, Math.max(0, doc.Content.End - 1)));
    }
    if (location === "selection") {
      // 宿主无选区时退回文档开头，与旧实现一致
      const r = doc.Range(0, 0);
      r.InsertAfter(text);
      return doc.Range(0, Math.min(text.length, Math.max(0, doc.Content.End - 1)));
    }
    const endPos = doc.Content.End > 1 ? doc.Content.End - 1 : 0;
    const newPara = doc.Paragraphs.Add(doc.Range(endPos, endPos));
    newPara.Range.Text = text;
    return newPara.Range;
  }

  function wordApplyLineFormat(doc, range, type, formatting) {
    if (type === "heading1") {
      try { range.Style = doc.Styles.Item(-2); } catch (e) { range.Font.Bold = true; range.Font.Size = 22; }
    } else if (type === "heading2") {
      try { range.Style = doc.Styles.Item(-3); } catch (e) { range.Font.Bold = true; range.Font.Size = 16; }
    } else if (type === "heading3") {
      try { range.Style = doc.Styles.Item(-4); } catch (e) { range.Font.Bold = true; range.Font.Size = 14; }
    } else if (type === "bullet_list") {
      try { range.ListFormat.ApplyBulletDefault(); } catch (e) {}
    } else if (type === "quote") {
      try {
        range.Font.Italic = true;
        range.ParagraphFormat.LeftIndent = 28;
      } catch (e) {}
    } else if (type === "code_block") {
      try {
        range.Font.NameFarEast = "Consolas";
        range.Font.NameAscii = "Consolas";
        range.Font.Size = 10.5;
      } catch (e) {}
    }

    if (formatting && typeof formatting === "object") {
      if (formatting.bold !== undefined) range.Font.Bold = Boolean(formatting.bold);
      if (formatting.italic !== undefined) range.Font.Italic = Boolean(formatting.italic);
      if (formatting.fontSizePt !== undefined) range.Font.Size = Number(formatting.fontSizePt);
      if (formatting.fontName) {
        range.Font.NameFarEast = formatting.fontName;
        range.Font.NameAscii = formatting.fontName;
      }
      if (formatting.alignment !== undefined) range.ParagraphFormat.Alignment = Number(formatting.alignment);
      if (formatting.firstLineIndentChars !== undefined) range.ParagraphFormat.CharacterUnitFirstLineIndent = Number(formatting.firstLineIndentChars);
      if (formatting.lineSpacingPt !== undefined) {
        range.ParagraphFormat.LineSpacingRule = 4;
        range.ParagraphFormat.LineSpacing = Number(formatting.lineSpacingPt);
      }
      if (formatting.spaceBeforePt !== undefined) range.ParagraphFormat.SpaceBefore = Number(formatting.spaceBeforePt);
      if (formatting.spaceAfterPt !== undefined) range.ParagraphFormat.SpaceAfter = Number(formatting.spaceAfterPt);
    } else if (!type || type === "paragraph") {
      // 普通正文默认强制消除前文加粗继承
      try { range.Font.Bold = false; } catch (e) {}
    }
  }

  function wordWriteContent(app, params) {
    const { documentName, location, targetBookmark, paragraphIndex, type, content, formatting } = params || {};
    const doc = getWordDocument(app, documentName);
    const wordApp = getWordApp() || app;
    const loc = location || "end";

    // 书签不存在时**显式拒绝**，不再静默落到文末（旧实现会给出 success，调用方无从发现）
    if (loc === "bookmark") {
      if (!targetBookmark) throw new Error("location=bookmark 必须提供 targetBookmark（书签名称）");
      if (!doc.Bookmarks.Exists(targetBookmark)) {
        const names = [];
        try {
          for (let i = 1; i <= doc.Bookmarks.Count; i++) names.push(doc.Bookmarks.Item(i).Name);
        } catch (e) {}
        throw new Error(
          `书签 [${targetBookmark}] 在此文档中不存在，未写入任何内容（拒绝静默落到文末）。` +
          (names.length ? `现有书签：${names.join("、")}` : "当前文档没有任何书签，请先用 wps_execute_script 的 doc.Bookmarks.Add 创建。")
        );
      }
    }
    if (loc === "after_paragraph" && !paragraphIndex) {
      throw new Error("location=after_paragraph 必须提供 paragraphIndex（1-based 段落序号）");
    }

    const lines = Array.isArray(content) ? content : [String(content || "")];
    const nonEmpty = lines.filter(t => t !== undefined && t !== null && String(t) !== "");
    if (nonEmpty.length === 0) {
      throw new Error("content 为空，未写入任何内容");
    }

    const meta = { insertedParagraphs: [], bookmarkPreserved: undefined };
    const writtenRanges = [];
    for (const raw of nonEmpty) {
      const text = String(raw);
      // 先查宿主是否吞字符：旧构建曾在 Paragraphs.Add(targetRange) + Range.Text 路径上吞掉小写字母，
      // 所以写入后按长度读回一次，长度不符就明确报错，而不是返回 success 让调用方踩坑（ISS-67）。
      const range = wordWriteOneLine(doc, wordApp, loc, targetBookmark, paragraphIndex, text, meta);
      wordApplyLineFormat(doc, range, type, formatting);
      writtenRanges.push(range);
      try {
        const readBack = (range.Text || "").replace(/[\r\n\x07]/g, "");
        if (readBack.length !== text.length) {
          throw new Error(
            `写入后读回长度不一致：写入 ${text.length} 个字符，读回 ${readBack.length} 个。` +
            `疑似宿主在 Range 文本赋值时吞字符；请勿重试覆盖，先读回核对，或改用 wps_execute_script 的 Range.InsertAfter。`
          );
        }
      } catch (e) {
        if (/疑似宿主在 Range 文本赋值时吞字符/.test(e.message)) throw e;
      }
      meta.insertedParagraphs.push({
        textLength: text.length,
        start: range.Start,
        end: range.End,
        style: (() => { try { return range.Style ? range.Style.NameLocal : undefined; } catch (e) { return undefined; } })()
      });
    }

    return {
      success: true,
      documentName: doc.Name,
      insertedLines: nonEmpty.length,
      location: loc,
      type: type || "paragraph",
      insertedParagraphs: meta.insertedParagraphs,
      bookmarkRange: loc === "bookmark" ? meta.bookmarkRange : undefined,
      bookmarkPreserved: loc === "bookmark" ? meta.bookmarkPreserved : undefined,
      message: loc === "bookmark"
        ? `已在书签 [${targetBookmark}] 所在段落之后写入 ${nonEmpty.length} 行内容` +
          (meta.bookmarkPreserved === false ? `；注意：宿主在插入过程中丢弃了该书签，需重建` : "")
        : `已成功向 [${doc.Name}] 写入 ${nonEmpty.length} 行内容`
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

  // 页码格式 → 页脚页码域。
  //
  // 真实宿主实测（WPS for Mac 12.0 / 12.1.28496）：
  //   - ✅ 唯一可用入口是 `Footers.Item(1).PageNumbers.Add(Alignment, FirstPage)`：
  //     调用后页脚出现真 PAGE 域，可读回 `footer.Range.Fields.Item(1).Code.Text === " PAGE "`；
  //   - ❌ `doc.Fields.Add(range, 33)` 传页脚范围时**静默返回 null、页脚一个域都不加**；
  //   - ❌ `footer.Range.InsertAfter("文字")` 在页脚 story 上被宿主静默丢弃（页脚文本仍为空）；
  //   - ❌ `footer.Range.InsertXML(<w:p>…fldChar/PAGE…)` 同样无效。
  // 因此只有 'simple'（纯页码）能可靠实现；'dash' / 'page_of_pages' 需要额外文字或 NUMPAGES 域，
  // 本机做不到，**显式拒绝**并给出可用替代写法，而不是静默降级成纯页码。
  function wordSetPageNumberFormat(doc, section, format, label) {
    const fmt = String(format || "").toLowerCase();
    if (fmt !== "simple") {
      return {
        ok: false,
        unsupported: true,
        error:
          `pageNumberFormat='${fmt}' 在当前宿主（WPS for Mac）上无法实现。` +
          `宿主只在页脚支持“纯页码”一种写法（可通过 Footers.Item(1).PageNumbers.Add 建真 PAGE 域）；` +
          `页脚的文字拼接与 NUMPAGES 域均被宿主静默丢弃（footer.Range.InsertAfter 与 InsertXML 实测无效）。` +
          `请改用 pageNumberFormat='simple'；需要 "- 1 -" 或 "1 / 5" 这类格式请在 Word 内手动插入页码后自行编辑，` +
          `或由调用方在 Windows/COM 通道处理。`
      };
    }
    let footer;
    try {
      footer = section.Footers.Item(1);
      if (footer.LinkToPrevious) footer.LinkToPrevious = false;
    } catch (e) {
      return { ok: false, error: `无法访问页脚: ${e.message}` };
    }
    try {
      footer.Range.Text = "";
      // wdAlignPageNumberCenter = 1；FirstPage = true
      footer.PageNumbers.Add(1, true);
      const fieldCount = (() => { try { return footer.Range.Fields.Count; } catch (e) { return 0; } })();
      const codes = [];
      try {
        for (let k = 1; k <= fieldCount; k++) codes.push((footer.Range.Fields.Item(k).Code.Text || "").trim());
      } catch (e) {}
      const readText = (footer.Range.Text || "").replace(/[\r\n\x07]/g, "");
      if (!fieldCount || !codes.some(c => /^PAGE\b/.test(c))) {
        return { ok: false, error: `页脚页码域写入后读回为 ${fieldCount} 个域（section ${label}），页码未生效`, footerText: readText };
      }
      return {
        ok: true,
        format: "simple",
        footerText: readText,
        fieldCount: fieldCount,
        fieldCodes: codes,
        pageNumberCount: (() => { try { return footer.PageNumbers.Count; } catch (e) { return null; } })()
      };
    } catch (e) {
      return { ok: false, error: `写入页码域失败: ${e.message}` };
    }
  }

  // 水印：优先尝试"页眉层"（跨页可见），失败则回退正文层并给出告警。
  //
  // 已实测的宿主事实（WPS for Mac 12.0 / 12.1.28496）：
  //   - `section.Headers.Item(1).Shapes.AddTextEffect(...)` **会静默把形状加到正文层**：
  //     调用后 `header.Shapes.Count` 恒为 0、`doc.Shapes.Count` +1（同一形状对象）；
  //   - `header.Range.ShapeRange.AddTextEffect` 是 undefined（"is not a function"）；
  //   - `Range.InsertXML`（VML `<w:pict>`）对页眉 story 无效，页眉 XML 长度不变；
  //   - `doc.Shapes.AddTextEffect` 落正文层：正文层浮动图形**只在第 1 页渲染**，不是"每页可见"。
  // 因此这里如实返回 placement，并把"跨页水印需另想办法"写进 warnings。
  function wordAddWatermarkToSection(doc, section, text, colorHex, pageSetup, warnings, sectionIndex) {
    const fontSize = 54;
    const shapeColor = hexToExcelColor(colorHex || "#C0C0C0") || 0xc0c0c0;
    let shape = null;
    let placement = "body";

    // 路径 1：页眉层
    try {
      const header = section.Headers.Item(1);
      header.Shapes.AddTextEffect(0, text, "Microsoft YaHei", fontSize, false, false, 0, 0);
      let headerCount = 0;
      try { headerCount = header.Shapes.Count; } catch (e) {}
      if (headerCount > 0) {
        placement = "header";
        try { shape = header.Shapes.Item(headerCount); } catch (e) {}
      }
    } catch (e) {
      warnings.push(`第 ${sectionIndex} 节页眉层水印写入失败：${e.message}`);
    }

    // 路径 2：正文层（页眉层不可用时的回退；宿主会把页眉 Shapes 的写入落到这里）
    if (!shape) {
      try {
        shape = doc.Shapes.AddTextEffect(0, text, "Microsoft YaHei", fontSize, false, false, 0, 0);
        placement = "body";
      } catch (e) {
        warnings.push(`第 ${sectionIndex} 节水印创建失败：${e.message}`);
        return { ok: false, error: e.message };
      }
    }

    let geometry = null;
    try {
      shape.Rotation = -315;
      try { shape.Fill.Transparency = 0.85; } catch (e) {}
      try { shape.Fill.ForeColor.RGB = shapeColor; } catch (e) {}
      try { shape.Line.Visible = false; } catch (e) {}
      try { shape.WrapFormat.Type = 3; } catch (e) {}
      // 居中：先把版式设为"相对页面"，再按页面尺寸居中（属性名在 WPS 上为 Range.ParagraphFormat 同族对象）
      const pw = pageSetup.pageWidth, ph = pageSetup.pageHeight;
      const w = Number(shape.Width) || 0, h = Number(shape.Height) || 0;
      shape.Left = Math.round(((pw - w) / 2) * 100) / 100;
      shape.Top = Math.round(((ph - h) / 2) * 100) / 100;
      geometry = { left: shape.Left, top: shape.Top, width: shape.Width, height: shape.Height, rotation: shape.Rotation };
    } catch (e) {
      warnings.push(`第 ${sectionIndex} 节水印属性设置部分失败：${e.message}`);
    }

    return {
      ok: true,
      sectionIndex: sectionIndex,
      placement: placement,
      shapeName: (() => { try { return shape.Name; } catch (e) { return undefined; } })(),
      geometry: geometry
    };
  }

  function wordPageLayoutAndWatermark(app, params) {
    const { documentName, headerText, footerText, pageNumberFormat, differentFirstPage, differentOddEvenPages, watermarkText, watermarkColor } = params || {};
    const doc = getWordDocument(app, documentName);
    const warnings = [];
    const sectionCount = doc.Sections.Count;

    // CAP-35 页眉页脚与水印读回：此前只能写不能读，AI 无法确认
    // "现在页眉里是什么""有没有水印"，也无法在改写前先看现状。
    if ((params || {}).action === "read") {
      const g = (fn, d = null) => { try { const x = fn(); return x === undefined ? d : x; } catch (e) { return d; } };
      const readHF = (hf) => ({
        exists: g(() => Boolean(hf.Exists), null),
        text: g(() => String(hf.Range.Text).replace(/[\r\n\x07]+$/, ""), ""),
        linkToPrevious: g(() => Boolean(hf.LinkToPrevious), null)
      });
      const sections = [];
      for (let i = 1; i <= sectionCount; i++) {
        const sec = doc.Sections.Item(i);
        sections.push({
          index: i,
          header: readHF(sec.Headers.Item(1)),          // wdHeaderFooterPrimary
          footer: readHF(sec.Footers.Item(1)),
          firstPageHeader: g(() => readHF(sec.Headers.Item(2)), null),   // wdHeaderFooterFirstPage
          firstPageFooter: g(() => readHF(sec.Footers.Item(2)), null),
          evenPagesHeader: g(() => readHF(sec.Headers.Item(3)), null),   // wdHeaderFooterEvenPages
          evenPagesFooter: g(() => readHF(sec.Footers.Item(3)), null)
        });
      }
      // 水印：WPS/Word 里是页眉中的 WordArt 形状，从 Header.Shapes 里找
      let watermark = null;
      try {
        const shapes = doc.Sections.Item(1).Headers.Item(1).Shapes;
        const n = Number(shapes.Count);
        for (let i = 1; i <= n; i++) {
          const sh = shapes.Item(i);
          const nm = String(g(() => sh.Name, ""));
          const txt = String(g(() => sh.TextEffect.Text, "") || g(() => sh.TextFrame.TextRange.Text, ""));
          if (nm.indexOf("WordArt") >= 0 || nm.indexOf("水印") >= 0 || txt) {
            watermark = { name: nm, text: txt, type: String(g(() => sh.Type, "")) };
            break;
          }
        }
      } catch (e) {}
      return {
        success: true,
        documentName: doc.Name,
        sectionCount,
        differentFirstPage: g(() => Boolean(doc.PageSetup.DifferentFirstPageHeaderFooter), null),
        differentOddEvenPages: g(() => Boolean(doc.PageSetup.OddAndEvenPagesHeaderFooter), null),
        sections,
        watermark,
        warnings,
        message: `文档 [${doc.Name}] 共 ${sectionCount} 节；第 1 节页眉「${String(sections[0]?.header?.text || "").slice(0, 30)}」页脚「${String(sections[0]?.footer?.text || "").slice(0, 30)}」；水印 ${watermark ? `「${watermark.text}」` : "无"}`
      };
    }

    if (headerText === undefined && footerText === undefined && pageNumberFormat === undefined && watermarkText === undefined &&
        differentFirstPage === undefined && differentOddEvenPages === undefined) {
      throw new Error("headerText / footerText / pageNumberFormat / watermarkText / differentFirstPage / differentOddEvenPages 至少传一项，否则本调用不产生任何变化");
    }

    // 文档级选项
    if (differentFirstPage !== undefined) doc.PageSetup.DifferentFirstPageHeaderFooter = differentFirstPage;
    if (differentOddEvenPages !== undefined) doc.PageSetup.OddAndEvenPagesHeaderFooter = differentOddEvenPages;

    const appliedSections = [];
    for (let i = 1; i <= sectionCount; i++) {
      const section = doc.Sections.Item(i);
      const entry = { sectionIndex: i, header: false, footer: false, pageNumberFormat: null, watermark: null };

      if (headerText !== undefined) {
        try {
          const header = section.Headers.Item(1);
          header.Range.Text = String(headerText);
          entry.header = true;
          entry.headerText = (header.Range.Text || "").replace(/[\r\n\x07]/g, "");
        } catch (e) {
          warnings.push(`第 ${i} 节页眉写入失败：${e.message}`);
        }
      }
      if (footerText !== undefined) {
        try {
          const footer = section.Footers.Item(1);
          footer.Range.Text = String(footerText);
          entry.footer = true;
          entry.footerText = (footer.Range.Text || "").replace(/[\r\n\x07]/g, "");
        } catch (e) {
          warnings.push(`第 ${i} 节页脚写入失败：${e.message}`);
        }
      }
      if (pageNumberFormat !== undefined) {
        const pn = wordSetPageNumberFormat(doc, section, pageNumberFormat, i);
        if (pn.ok) {
          entry.pageNumberFormat = { format: pn.format, footerText: pn.footerText, fieldCount: pn.fieldCount, fieldCodes: pn.fieldCodes, pageNumberCount: pn.pageNumberCount };
        } else {
          entry.pageNumberFormat = { format: String(pageNumberFormat).toLowerCase(), applied: false, error: pn.error };
          warnings.push(`第 ${i} 节页码未写入：${pn.error}`);
        }
      }
      if (watermarkText) {
        const ps = (() => {
          try {
            return { pageWidth: Number(doc.PageSetup.PageWidth) || 0, pageHeight: Number(doc.PageSetup.PageHeight) || 0 };
          } catch (e) { return { pageWidth: 0, pageHeight: 0 }; }
        })();
        const wm = wordAddWatermarkToSection(doc, section, String(watermarkText), watermarkColor, ps, warnings, i);
        if (wm.ok) entry.watermark = wm;
      }

      appliedSections.push(entry);
    }

    if (watermarkText && sectionCount > 1) {
      warnings.push(
        `水印已按 ${sectionCount} 个节分别写入，但宿主 WPS for Mac 无法把形状放进页眉层` +
        `（Headers.Shapes 的写入会静默落到正文层），因此水印仍是**正文层浮动图形**：` +
        `实测只在第 1 页渲染，不是"每页可见"。跨页水印需在 Word 内手动插入（插入 → 水印），` +
        `或由调用方在 Windows/COM 通道用 Section.Headers.Shapes 处理。`
      );
    } else if (watermarkText) {
      warnings.push(
        `水印落在正文层（宿主 WPS for Mac 的 Headers.Shapes 写入会静默落到正文层）：` +
        `正文层浮动图形只在第 1 页渲染，不是"每页可见"。跨页水印需在 Word 内手动插入（插入 → 水印），` +
        `或改用 Windows/COM 通道。`
      );
    }

    return {
      success: true,
      documentName: doc.Name,
      sectionCount: sectionCount,
      appliedSections: appliedSections,
      header: headerText,
      footer: footerText,
      pageNumberFormat: pageNumberFormat || undefined,
      watermark: watermarkText || undefined,
      warnings: warnings,
      message:
        `已更新 [${doc.Name}] 页面版式：${sectionCount} 个节` +
        (headerText !== undefined ? "，页眉" : "") +
        (footerText !== undefined ? "，页脚" : "") +
        (pageNumberFormat !== undefined ? `，页码(${pageNumberFormat})` : "") +
        (watermarkText ? "，水印" : "") +
        (warnings.length ? `；有 ${warnings.length} 条告警，请逐条查看 warnings` : "")
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

  // ===========================================================================
  // CAP-04：Word 域更新与交叉引用（RPC: word_update_fields）
  //
  // 真机实测（WPS for Mac 12.x，2026-09-22；把 `测试文字文稿.docx` 的内容经
  // `Selection.InsertFile` 导入临时文档后测量，原文件未被改动）：
  //   1. 目录页码**不会**随内容自动刷新——在 TOC 之后插入一个整页分页符后，TOC 内 14 个
  //      PAGEREF 结果仍是 ["3"×6,"4"×3,"5"×5]（陈旧）；`doc.Fields.Update()` 100ms 后
  //      变为 ["4"×6,"5"×3,"6"×5]，逐项 +1。这就是"交付物一改目录页码就全错"的根因。
  //   2. `doc.Fields.Update()` **无弹窗、无交互**（实测 100ms / 63 个域）；
  //      但它的**返回值实测恒为 0**，不能当"更新了几个域"的计数——本工具改为更新前后
  //      逐域结果快照对比，自己数 changedFields。
  //   3. `doc.Fields.Update()` 会**连目录条目一起重建**（新增标题会被收录）；
  //      `TablesOfContents.Item(n).Update()` 同样可重建；`UpdatePageNumbers()` 只刷页码、
  //      不收录新标题（实测新增标题后只调 UpdatePageNumbers，目录里仍然没有该标题）。
  //   4. `doc.Fields` **不含页眉/页脚 story 里的域**，页眉页脚域必须逐节
  //      `section.Headers/Footers.Item(k).Range.Fields.Update()` 单独更新。
  //   5. 交叉引用**只能按书签**：`Application.CrossReference` 与 `Document.CrossReference`
  //      实测均为 `undefined`，没有 Word 那套"引用类型 + 引用内容"选择器。
  //      REF(wdFieldRef=3) / PAGEREF(wdFieldPageRef=37) 用 `Fields.Add(range, type, "书名签名 \\h", false)`
  //      插入后**立即**就有结果，无需等待。书签不存在 → 结果恒为 "错误！未定义书签。"；
  //      书签为空（起止位置相同）→ REF 结果为空字符串（两者都实测过）。
  //   6. **域后继续插入内容必须跳过域结束标记**：`Fields.Add` 之后若按
  //      `field.Result.End` 直接插入，插入内容会落在域内部，下一次 `Fields.Update()`
  //      会把这段内容连同后面的域一起清掉（实测 advance=0 时 "（第 1 页）" 与 PAGEREF 域消失）。
  //      正确推进量是 `field.Result.End + 1`（实测 advance=1 时模板完整存活、两个域都在）。
  // ===========================================================================

  const WORD_FIELD_TYPE_NAMES = {
    3: "REF", 13: "TOC", 26: "NUMPAGES", 33: "PAGE", 37: "PAGEREF", 54: "AutoNum", 88: "HYPERLINK"
  };

  /** 域类型名：优先取域码首个单词（更准），否则退回类型号映射。 */
  function wordFieldKindName(type, code) {
    const m = /^\s*([A-Za-z]+)/.exec(String(code || ""));
    if (m) return m[1].toUpperCase();
    return WORD_FIELD_TYPE_NAMES[Number(type)] || String(type);
  }

  function wordNormalizeFieldText(text) {
    return String(text === undefined || text === null ? "" : text)
      .replace(/[\r\n\x07\xa0]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  /** 域快照：更新前后各取一次，用于自己统计"哪些域真的变了"。 */
  function wordFieldSnapshot(doc, limit) {
    const max = Number(limit) || 200;
    let total = 0;
    try { total = doc.Fields.Count; } catch (e) { return { total: 0, list: [], truncated: false, error: e.message }; }
    const list = [];
    for (let i = 1; i <= Math.min(total, max); i++) {
      try {
        const f = doc.Fields.Item(i);
        const code = (f.Code ? f.Code.Text : "").replace(/[\r\n\x07]/g, " ").trim();
        list.push({
          index: i,
          type: Number(f.Type),
          kind: wordFieldKindName(f.Type, code),
          code: code,
          result: wordNormalizeFieldText(f.Result ? f.Result.Text : "")
        });
      } catch (e) {
        list.push({ index: i, kind: "?", code: "", result: "", error: e.message });
      }
    }
    return { total: total, list: list, truncated: total > max };
  }

  /**
   * 两个域快照的差异。
   * 不能按 index 直接比：目录重建后域下标会整体平移（实测更新后域数 65 → 63）。
   * 按 `kind|code` 分桶、同桶内按出现顺序配对比较，才是稳定口径。
   */
  function wordDiffFieldSnapshots(before, after, maxChanges) {
    const limit = Number(maxChanges) || 60;
    const bucket = new Map();
    for (const f of before.list) {
      if (f.error) continue;
      const k = f.kind + "|" + f.code;
      if (!bucket.has(k)) bucket.set(k, []);
      bucket.get(k).push(f.result);
    }
    const cursor = new Map();
    const changes = [];
    let compared = 0;
    let changed = 0;
    for (const f of after.list) {
      if (f.error) continue;
      const k = f.kind + "|" + f.code;
      const arr = bucket.get(k);
      if (!arr) continue;
      const n = cursor.get(k) || 0;
      cursor.set(k, n + 1);
      if (n >= arr.length) continue;
      compared++;
      if (arr[n] !== f.result) {
        changed++;
        if (changes.length < limit) {
          changes.push({ kind: f.kind, code: f.code, resultBefore: arr[n], resultAfter: f.result });
        }
      }
    }
    return { compared: compared, changed: changed, changes: changes, changesTruncated: changed > changes.length };
  }

  /**
   * 目录条目标题：HYPERLINK 域的结果形如 "一、概述\t4"（制表符 + 页码）。
   * **必须在归一化之前切**：wordNormalizeFieldText 会把制表符压成空格，
   * 之后再按 \t 切就切不掉页码，条目比对会全错（本轮实测到过这个假差异）。
   */
  function wordTocEntryTitle(rawResult) {
    const raw = String(rawResult === undefined || rawResult === null ? "" : rawResult).replace(/[\r\n\x07]/g, "");
    const tabIdx = raw.indexOf("\t");
    if (tabIdx >= 0) return raw.slice(0, tabIdx).trim();
    return raw.replace(/[\s\xa0]+\d+\s*$/, "").trim();
  }

  /** 目录快照：条目文本 + 各条目的页码（目录页码本体就是条目里的 PAGEREF 域结果）。 */
  function wordTocSnapshot(doc) {
    let count = 0;
    try { count = doc.TablesOfContents.Count; } catch (e) { return { count: 0, list: [], error: e.message }; }
    const list = [];
    for (let i = 1; i <= count; i++) {
      const entry = { index: i, entryCount: 0, entryTitles: [], pageNumbers: [], range: null };
      try {
        const toc = doc.TablesOfContents.Item(i);
        const fields = toc.Range.Fields;
        for (let k = 1; k <= fields.Count; k++) {
          const f = fields.Item(k);
          const code = (f.Code ? f.Code.Text : "").replace(/[\r\n\x07]/g, " ").trim();
          const rawResult = f.Result ? f.Result.Text : "";
          if (/^PAGEREF\b/i.test(code)) entry.pageNumbers.push(wordNormalizeFieldText(rawResult));
          else if (/^HYPERLINK\b/i.test(code)) {
            entry.entryCount++;
            entry.entryTitles.push(wordTocEntryTitle(rawResult));
          }
        }
        entry.range = { start: toc.Range.Start, end: toc.Range.End };
        entry.upperHeadingLevel = (() => { try { return Number(toc.UpperHeadingLevel); } catch (e) { return undefined; } })();
        entry.lowerHeadingLevel = (() => { try { return Number(toc.LowerHeadingLevel); } catch (e) { return undefined; } })();
      } catch (e) { entry.error = e.message; }
      list.push(entry);
    }
    return { count: count, list: list };
  }

  /** 目录更新前后对照（页码逐条比对 + 条目增删）。 */
  function wordCompareTocs(before, after) {
    const out = [];
    const count = Math.max(before.count, after.count);
    for (let i = 1; i <= count; i++) {
      const b = before.list[i - 1] || {};
      const a = after.list[i - 1] || {};
      const pagesB = b.pageNumbers || [];
      const pagesA = a.pageNumbers || [];
      const titlesB = b.entryTitles || [];
      const titlesA = a.entryTitles || [];
      const changedPages = [];
      const maxLen = Math.max(pagesB.length, pagesA.length);
      for (let k = 0; k < maxLen; k++) {
        if (pagesB[k] !== pagesA[k]) {
          changedPages.push({
            entry: titlesA[k] || titlesB[k] || `#${k + 1}`,
            pageBefore: pagesB[k] === undefined ? null : pagesB[k],
            pageAfter: pagesA[k] === undefined ? null : pagesA[k]
          });
        }
      }
      out.push({
        index: i,
        entryCountBefore: b.entryCount === undefined ? null : b.entryCount,
        entryCountAfter: a.entryCount === undefined ? null : a.entryCount,
        pageNumbersBefore: pagesB.slice(0, 60),
        pageNumbersAfter: pagesA.slice(0, 60),
        pageNumbersChanged: changedPages.length > 0,
        changedPageCount: changedPages.length,
        changedPages: changedPages.slice(0, 30),
        changedPagesTruncated: changedPages.length > 30,
        addedEntries: titlesA.filter(t => titlesB.indexOf(t) < 0).slice(0, 20),
        removedEntries: titlesB.filter(t => titlesA.indexOf(t) < 0).slice(0, 20),
        error: a.error || b.error
      });
    }
    return out;
  }

  /** 节的页眉/页脚 story 槽位：只在文档真的启用"首页不同/奇偶页不同"时才去碰 2/3 号槽位。 */
  function wordSectionStorySlots(doc, sectionIndex) {
    const slots = [1];
    try {
      const ps = doc.Sections.Item(sectionIndex).PageSetup;
      if (Number(ps.DifferentFirstPageHeaderFooter) !== 0) slots.push(2);
      if (Number(ps.OddAndEvenPagesHeaderFooter) !== 0) slots.push(3);
    } catch (e) {}
    return slots;
  }

  function wordStorySnapshotAndUpdate(doc, sectionIndex, kind, slot, doUpdate, warnings) {
    const entry = { story: kind + ".Item(" + slot + ")", ok: false };
    try {
      const section = doc.Sections.Item(sectionIndex);
      const story = kind === "Headers" ? section.Headers.Item(slot) : section.Footers.Item(slot);
      const fields = story.Range.Fields;
      entry.fieldCountBefore = fields.Count;
      entry.fieldCodesBefore = [];
      for (let k = 1; k <= fields.Count; k++) {
        entry.fieldCodesBefore.push((fields.Item(k).Code ? fields.Item(k).Code.Text : "").replace(/[\r\n\x07]/g, " ").trim());
      }
      if (doUpdate) {
        const t0 = Date.now();
        fields.Update();
        entry.elapsedMs = Date.now() - t0;
      }
      entry.fieldCountAfter = story.Range.Fields.Count;
      entry.text = (story.Range.Text || "").replace(/[\r\n\x07]/g, "").slice(0, 120);
      entry.ok = true;
    } catch (e) {
      entry.error = e.message;
      if (doUpdate) warnings.push(`第 ${sectionIndex} 节 ${kind}.Item(${slot}) 域更新失败：${e.message}`);
    }
    return entry;
  }

  /** 按出现次数查一段文本，返回命中的 Range（与生产其它函数一致的 Find 用法）。 */
  function wordFindTextRangeIn(doc, text, occurrence) {
    return wordFindAnchorRange(doc, text, occurrence, null);
  }

  /**
   * 命中范围是否落在**目录或域结果内部**。
   *
   * 真机实测（2026-09-22）：把书签打在目录条目上（目录条目本身是 HYPERLINK/PAGEREF 域的结果），
   * 随后一次 `doc.Fields.Update()` 重建目录就把该书签**销毁**了，
   * 引用它的交叉引用立刻变成 "错误！未定义书签。"（探针 xref_by_anchorText → docTextAfterXref）。
   * 而 `Find` 从文首搜 "第二章 交付内容" 时**先命中目录里的那条**，所以必须显式跳过。
   */
  function wordUnsafeAnchorReason(doc, range) {
    try {
      const tocCount = doc.TablesOfContents.Count;
      for (let i = 1; i <= tocCount; i++) {
        try {
          const r = doc.TablesOfContents.Item(i).Range;
          if (range.Start >= r.Start && range.End <= r.End) return `第 ${i} 个目录内部（目录条目是域结果，域重建会销毁其中的书签）`;
        } catch (e) {}
      }
    } catch (e) {}
    try {
      const total = Math.min(doc.Fields.Count, 200);
      for (let i = 1; i <= total; i++) {
        try {
          const f = doc.Fields.Item(i);
          const res = f.Result;
          if (res && res.End > res.Start && range.Start >= res.Start && range.End <= res.End) {
            const code = (f.Code ? f.Code.Text : "").replace(/[\r\n\x07]/g, " ").trim().slice(0, 40);
            return `域结果内部（域码: ${code}）`;
          }
        } catch (e) {}
      }
    } catch (e) {}
    return null;
  }

  /**
   * 按出现次数查文本，默认**跳过目录/域结果内部**的命中（那里不能放书签与内容控件）。
   * meta.skipped 会记录被跳过的命中与原因，供调用方如实回报。
   */
  function wordFindAnchorRange(doc, text, occurrence, meta) {
    const want = Number(occurrence) || 1;
    const needle = String(text);
    let hitCount = 0;
    const scan = (rng) => {
      let searchFrom = rng.Start;
      const limit = rng.End;
      let guard = 0;
      while (guard++ < 2000 && searchFrom < limit) {
        let hit = null;
        try {
          const sub = doc.Range(searchFrom, limit);
          const f = sub.Find;
          f.ClearFormatting();
          f.Text = needle;
          f.MatchCase = false;
          f.MatchWholeWord = false;
          f.MatchWildcards = false;
          f.Forward = true;
          f.Wrap = 0; // wdFindStop
          if (!f.Execute()) return null;
          hit = f.Parent;
        } catch (e) {
          return null;
        }
        if (!hit) return null;
        const reason = wordUnsafeAnchorReason(doc, hit);
        if (!reason) {
          hitCount++;
          if (hitCount >= want) return hit;
        } else if (meta) {
          meta.skipped = (meta.skipped || []).concat([{
            text: (hit.Text || "").replace(/[\r\n\x07]/g, "").slice(0, 40),
            range: { start: hit.Start, end: hit.End },
            reason: reason
          }]);
        }
        const next = hit.End > searchFrom ? hit.End : searchFrom + 1;
        searchFrom = next;
      }
      return null;
    };
    const hit = scan(doc.Content);
    if (hit) return hit;
    if (doc.Tables) {
      for (let t = 1; t <= doc.Tables.Count; t++) {
        try {
          const h = scan(doc.Tables.Item(t).Range);
          if (h) return h;
        } catch (e) {}
      }
    }
    return null;
  }

  function wordExistingBookmarkNames(doc, limit) {
    const names = [];
    try {
      for (let i = 1; i <= doc.Bookmarks.Count && names.length < (limit || 30); i++) names.push(doc.Bookmarks.Item(i).Name);
    } catch (e) {}
    return names;
  }

  /**
   * 命中范围若落在目录/域结果内部，返回该容器的结束位置（用于把插入点自动挪到容器之外）。
   * 与 wordUnsafeAnchorReason 同一份实测依据（目录条目是域结果，域重建会清掉其中的内容）。
   */
  function wordUnsafeInsertContainer(doc, range) {
    try {
      const tocCount = doc.TablesOfContents.Count;
      for (let i = 1; i <= tocCount; i++) {
        try {
          const r = doc.TablesOfContents.Item(i).Range;
          if (range.Start >= r.Start && range.End <= r.End) {
            return { end: r.End, reason: `插入点落在第 ${i} 个目录内部（目录是域结果，写进去的内容会被下一次域重建清掉）` };
          }
        } catch (e) {}
      }
    } catch (e) {}
    try {
      const total = Math.min(doc.Fields.Count, 200);
      for (let i = 1; i <= total; i++) {
        try {
          const f = doc.Fields.Item(i);
          const res = f.Result;
          if (res && res.End > res.Start && range.Start >= res.Start && range.End <= res.End) {
            const code = (f.Code ? f.Code.Text : "").replace(/[\r\n\x07]/g, " ").trim().slice(0, 40);
            return { end: res.End + 1, reason: `插入点落在域结果内部（域码: ${code}）` };
          }
        } catch (e) {}
      }
    } catch (e) {}
    return null;
  }

  /**
   * 交叉引用插入：REF（引用书签文字）/ PAGEREF（引用书签所在页码）。
   * 支持先用 anchorText 现场建书签，再插域——书签是交叉引用的唯一前提（宿主无 CrossReference）。
   */
  function wordInsertCrossReference(app, doc, params) {
    const {
      targetBookmark, bookmarkName, anchorText, anchorOccurrence,
      crossReferenceType, crossReferenceTemplate, insertLocation, paragraphIndex, prefixText
    } = params || {};
    const warnings = [];
    const name = String(bookmarkName || targetBookmark || "").trim();
    let bookmark = null;

    if (anchorText !== undefined && anchorText !== null && String(anchorText) !== "") {
      if (!name) throw new Error("用 anchorText 现建书签时必须同时给 bookmarkName（书签名，不能含空格）");
      if (/\s/.test(name)) throw new Error(`书签名 [${name}] 含空格，宿主 Word 书签名不允许空格`);
      const meta = { skipped: [] };
      const range = wordFindAnchorRange(doc, String(anchorText), anchorOccurrence, meta);
      if (!range) {
        throw new Error(
          `未在文档中找到可安全锚定的文本 [${anchorText}]（第 ${Number(anchorOccurrence) || 1} 处），未建书签也未插入交叉引用。` +
          (meta.skipped.length
            ? ` 已跳过 ${meta.skipped.length} 处命中：` + meta.skipped.map(s => `[${s.text}] 在${s.reason}`).join("；")
            : "")
        );
      }
      if (meta.skipped.length) {
        warnings.push(
          `有 ${meta.skipped.length} 处命中被跳过（锚点必须落在正文，不能落在目录或域结果里——` +
          `真机实测：目录条目的书签会被下一次域重建销毁，交叉引用随即变成"错误！未定义书签。"）：` +
          meta.skipped.map(s => `[${s.text}] 在${s.reason}`).join("；")
        );
      }
      const rangeInfo = { start: range.Start, end: range.End, text: (range.Text || "").replace(/[\r\n\x07]/g, "") };
      try { doc.Bookmarks.Add(name, range); } catch (e) { throw new Error(`Bookmarks.Add("${name}") 失败：${e.message}`); }
      let exists = false;
      try { exists = doc.Bookmarks.Exists(name); } catch (e) {}
      if (!exists) throw new Error(`Bookmarks.Add("${name}") 执行后书签仍不存在，宿主未创建；已中止，未插入交叉引用`);
      bookmark = { name: name, created: true, range: rangeInfo, text: rangeInfo.text, empty: rangeInfo.start === rangeInfo.end, exists: true };
    } else {
      if (!name) throw new Error("必须提供 bookmarkName 或 targetBookmark —— 本宿主没有 CrossReference 对象，交叉引用**只能**按书签");
      let exists = false;
      try { exists = doc.Bookmarks.Exists(name); } catch (e) {}
      if (!exists) {
        const names = wordExistingBookmarkNames(doc);
        throw new Error(
          `书签 [${name}] 在此文档中不存在，未插入任何交叉引用（REF/PAGEREF 引用不到书签时结果恒为"错误！未定义书签。"）。` +
          (names.length ? `现有书签：${names.join("、")}` : "当前文档没有任何书签；可传 anchorText + bookmarkName 让本工具先建书签。")
        );
      }
      const bm = doc.Bookmarks.Item(name);
      const r = bm.Range;
      bookmark = {
        name: name, created: false, exists: true,
        range: { start: r.Start, end: r.End },
        text: (r.Text || "").replace(/[\r\n\x07]/g, ""),
        empty: r.Start === r.End
      };
    }
    if (bookmark.empty) {
      warnings.push(`书签 [${bookmark.name}] 是空书签（起止位置相同）：真机实测 REF 域结果为空字符串、PAGEREF 仍能给出页码。`);
    }

    // 插入位置
    const loc = insertLocation || "end";
    let pos = null;
    if (loc === "start") pos = 0;
    else if (loc === "selection") {
      const wordApp = getWordApp() || app;
      try {
        const sel = wordApp && wordApp.Selection && wordApp.Selection.Range;
        if (sel) pos = sel.Start === sel.End ? sel.Start : sel.End;
      } catch (e) {}
      if (pos === null) { warnings.push("当前没有可用选区，交叉引用回退插入到文档末尾。"); }
    } else if (loc === "after_paragraph") {
      const idx = Number(paragraphIndex);
      if (!idx || idx < 1 || idx > doc.Paragraphs.Count) {
        throw new Error(`段落索引越界：第 ${paragraphIndex} 段不存在（当前共 ${doc.Paragraphs.Count} 段）`);
      }
      pos = doc.Paragraphs.Item(idx).Range.End - 1;
    }
    if (pos === null) pos = doc.Content.End > 1 ? doc.Content.End - 1 : 0;
    // 插入点落在目录/域结果内部时自动挪到容器之外：写进去的内容会被下一次域重建清掉（实测）
    try {
      const container = wordUnsafeInsertContainer(doc, doc.Range(pos, pos));
      if (container) {
        warnings.push(`${container.reason}；插入点已自动移到该容器之后（位置 ${pos} → ${container.end}）。`);
        pos = container.end;
      }
    } catch (e) {}

    // 模板：{ref} / {page}
    const kind = crossReferenceType || "both";
    if (["ref", "pageref", "both"].indexOf(kind) < 0) {
      throw new Error(`未知的 crossReferenceType: ${kind}（支持 ref, pageref, both）`);
    }
    const template = kind === "both"
      ? String(crossReferenceTemplate || "{ref}（第 {page} 页）")
      : (kind === "ref" ? "{ref}" : "{page}");
    const segments = [];
    const re = /\{(ref|page)\}/g;
    let last = 0;
    let m;
    while ((m = re.exec(template)) !== null) {
      if (m.index > last) segments.push({ text: template.slice(last, m.index) });
      segments.push({ field: m[1] });
      last = m.index + m[0].length;
    }
    if (last < template.length) segments.push({ text: template.slice(last) });

    const inserted = [];
    const errors = [];
    if (prefixText) {
      try { doc.Range(pos, pos).InsertAfter(String(prefixText)); pos += String(prefixText).length; } catch (e) { errors.push(`前缀写入失败：${e.message}`); }
    }
    for (const seg of segments) {
      try {
        if (seg.text) {
          doc.Range(pos, pos).InsertAfter(seg.text);
          pos += seg.text.length;
          continue;
        }
        const type = seg.field === "ref" ? 3 : 37; // wdFieldRef / wdFieldPageRef
        const f = doc.Fields.Add(doc.Range(pos, pos), type, bookmark.name + " \\h", false);
        if (f === null || f === undefined) {
          errors.push(`${seg.field === "ref" ? "REF" : "PAGEREF"} 域插入返回 null（宿主未创建），该段被跳过`);
          continue;
        }
        // 关键：+1 跳过域结束标记，否则后续内容落在域内部，下次 Update 会被清掉（真机实测）
        pos = f.Result.End + 1;
        inserted.push({
          field: seg.field === "ref" ? "REF" : "PAGEREF",
          type: type,
          code: (f.Code.Text || "").replace(/[\r\n\x07]/g, " ").trim(),
          result: wordNormalizeFieldText(f.Result.Text),
          _field: f
        });
      } catch (e) {
        errors.push(`${seg.field ? (seg.field === "ref" ? "REF" : "PAGEREF") + " 域" : "文本段"}插入失败：${e.message}`);
      }
    }

    let updateOk = false;
    try { doc.Fields.Update(); updateOk = true; } catch (e) { warnings.push(`插入后 Fields.Update() 失败：${e.message}`); }

    // 读回：**优先读本次插入的那个域对象**。
    // 不能只按域码回查：同一书签可能被引用多次（域码完全相同），按码匹配会读到别人那一个
    // （本轮实测到过：after_paragraph 插入的 PAGEREF 回读成了先前那个、报出假的"未定义书签"）。
    for (const item of inserted) {
      const nativeField = item._field;
      delete item._field;
      try {
        if (nativeField) {
          item.resultAfterUpdate = wordNormalizeFieldText(nativeField.Result.Text);
          item.bookmarkStillExists = (() => { try { return doc.Bookmarks.Exists(bookmark.name); } catch (e) { return undefined; } })();
        }
      } catch (e) {
        warnings.push(`${item.field} 域对象在更新后不可读（${e.message}），改用域码回查`);
      }
      if (item.resultAfterUpdate === undefined) {
        try {
          for (let k = 1; k <= doc.Fields.Count; k++) {
            const f = doc.Fields.Item(k);
            const code = (f.Code ? f.Code.Text : "").replace(/[\r\n\x07]/g, " ").trim();
            if (code === item.code) {
              item.resultAfterUpdate = wordNormalizeFieldText(f.Result.Text);
              item.fieldIndex = k;
              break;
            }
          }
        } catch (e) {}
      }
      if (item.resultAfterUpdate === undefined) {
        item.resolved = false;
        warnings.push(`${item.field} 域在更新后未能重新定位，无法确认结果`);
      } else {
        item.resolved = item.resultAfterUpdate !== "" && !/错误|Error!|未定义书签/.test(item.resultAfterUpdate);
      }
    }
    const bookmarkSurvived = (() => { try { return doc.Bookmarks.Exists(bookmark.name); } catch (e) { return undefined; } })();
    if (bookmarkSurvived === false) {
      warnings.push(`书签 [${bookmark.name}] 在域更新后**已不存在**：宿主在重建域时销毁了它。请改用正文里的锚点文本（本工具会自动跳过目录/域结果内部的命中）。`);
    }

    return {
      success: true,
      documentName: doc.Name,
      action: "insert_cross_reference",
      bookmark: bookmark,
      bookmarkStillExistsAfterUpdate: bookmarkSurvived,
      insertLocation: loc,
      insertedFields: inserted,
      fieldsUpdateAfterInsert: updateOk,
      errors: errors.length ? errors : undefined,
      warnings: warnings,
      hostLimitations: [
        "本宿主没有 CrossReference 对象（Application/Document.CrossReference 实测 undefined），交叉引用只能按**书签**：不能按'标题/图表编号'这类 Word 内置引用类型插入。",
        "REF/PAGEREF 引用不存在的书签时结果恒为 '错误！未定义书签。'；空书签的 REF 结果为空字符串（均已实测）。",
        "域插入后立即就有结果，但页码要在文档完成分页后才准确；若结果可疑请再调 action='update' 刷新一次。"
      ],
      message: `已在 [${doc.Name}] 插入 ${inserted.length} 个交叉引用域（书签 [${bookmark.name}]${bookmark.created ? "，本次新建" : ""}）` +
        (errors.length ? `；有 ${errors.length} 条失败，见 errors` : "") +
        (warnings.length ? `；有 ${warnings.length} 条告警，见 warnings` : "")
    };
  }

  function wordUpdateFields(app, params) {
    const { documentName, action, scope, sectionIndex, tocMode, includeHeadersFooters } = params || {};
    const doc = getWordDocument(app, documentName);
    const act = action || "update";

    if (act === "insert_cross_reference") return wordInsertCrossReference(app, doc, params);
    if (act !== "update") {
      throw new Error(`未知的 Word 域操作: ${action}（支持 update、insert_cross_reference）`);
    }

    const scopeVal = scope || "all";
    if (["all", "toc", "section"].indexOf(scopeVal) < 0) {
      throw new Error(`未知的 scope: ${scopeVal}（支持 all、toc、section；scope='section' 时必须配 sectionIndex）`);
    }
    const mode = tocMode || "full";
    if (["full", "page_numbers"].indexOf(mode) < 0) {
      throw new Error(`未知的 tocMode: ${tocMode}（支持 full、page_numbers）`);
    }
    const warnings = [];
    const errors = [];
    const startedAt = Date.now();

    const sectionCount = (() => { try { return doc.Sections.Count; } catch (e) { return 0; } })();
    let sectionTargets = [];
    if (scopeVal === "section") {
      const idx = Number(sectionIndex);
      if (!idx || idx < 1 || idx > sectionCount) {
        throw new Error(`sectionIndex 越界：收到 ${sectionIndex}，当前文档共 ${sectionCount} 个节`);
      }
      sectionTargets = [idx];
    } else if (scopeVal === "all") {
      for (let i = 1; i <= sectionCount; i++) sectionTargets.push(i);
    }

    const fieldsBefore = wordFieldSnapshot(doc, 200);
    const tocBefore = wordTocSnapshot(doc);

    // 1) 正文域
    const bodyUpdate = { applied: false, scope: scopeVal };
    if (scopeVal === "all") {
      try {
        const t0 = Date.now();
        const ret = doc.Fields.Update();
        bodyUpdate.applied = true;
        bodyUpdate.elapsedMs = Date.now() - t0;
        bodyUpdate.hostReturnValue = typeof ret === "number" ? ret : undefined;
        bodyUpdate.note = "宿主 Fields.Update() 的返回值实测恒为 0，不能当计数用；本工具按快照差异统计。";
        bodyUpdate.coversTableOfContents = "宿主实测：doc.Fields.Update() 会连目录条目一起重建（新增标题会被收录），tocMode='full' 时无需再单独调 toc.Update()。";
      } catch (e) {
        errors.push(`正文域更新失败：${e.message}`);
      }
    } else if (scopeVal === "section") {
      for (const idx of sectionTargets) {
        const entry = { sectionIndex: idx, ok: false };
        try {
          const t0 = Date.now();
          const ret = doc.Sections.Item(idx).Range.Fields.Update();
          entry.fieldsCount = doc.Sections.Item(idx).Range.Fields.Count;
          entry.hostReturnValue = typeof ret === "number" ? ret : undefined;
          entry.elapsedMs = Date.now() - t0;
          entry.ok = true;
        } catch (e) { entry.error = e.message; errors.push(`第 ${idx} 节正文域更新失败：${e.message}`); }
        bodyUpdate.sections = (bodyUpdate.sections || []).concat([entry]);
      }
    }

    // 2) 目录
    const tocUpdate = [];
    if (scopeVal === "all" || scopeVal === "toc") {
      let count = 0;
      try { count = doc.TablesOfContents.Count; } catch (e) { errors.push(`读取目录数量失败：${e.message}`); }
      if (count === 0) {
        warnings.push("文档里没有任何目录（TablesOfContents.Count = 0），本次没有可更新的目录对象。");
      }
      for (let i = 1; i <= count; i++) {
        const item = { index: i, mode: mode, ok: false };
        if (scopeVal === "all" && mode === "full") {
          item.skippedBecause = "doc.Fields.Update() 已重建全部目录（含条目与页码），无需重复调用 toc.Update()";
          item.ok = true;
          tocUpdate.push(item);
          continue;
        }
        try {
          const t0 = Date.now();
          const toc = doc.TablesOfContents.Item(i);
          if (mode === "page_numbers") toc.UpdatePageNumbers();
          else toc.Update();
          item.elapsedMs = Date.now() - t0;
          item.ok = true;
        } catch (e) {
          item.error = e.message;
          errors.push(`第 ${i} 个目录更新失败：${e.message}`);
        }
        tocUpdate.push(item);
      }
    }

    // 3) 页眉/页脚域（逐节逐 story；doc.Fields 不含页眉页脚 story）
    const storyUpdates = [];
    if (includeHeadersFooters !== false && (scopeVal === "all" || scopeVal === "section")) {
      for (const idx of sectionTargets) {
        for (const kind of ["Headers", "Footers"]) {
          for (const slot of wordSectionStorySlots(doc, idx)) {
            storyUpdates.push(wordStorySnapshotAndUpdate(doc, idx, kind, slot, true, warnings));
          }
        }
      }
    }

    const fieldsAfter = wordFieldSnapshot(doc, 200);
    const tocAfter = wordTocSnapshot(doc);
    const diff = wordDiffFieldSnapshots(fieldsBefore, fieldsAfter, 60);
    const tocComparison = wordCompareTocs(tocBefore, tocAfter);
    const pageNumbersChanged = tocComparison.some(t => t.pageNumbersChanged);

    // 目录内部的域（PAGEREF/HYPERLINK）在目录重建后**域码里的 _Toc 书签名会整批变化**，
    // 按 `kind|code` 配对根本配不上（实测更新后 5 个条目域只配到 1 个 TOC 域）。
    // 因此目录内的变化改用目录快照对比来数，并合进 changes 明细，别让"只变了 1 个域"误导调用方。
    let tocEntryFieldsChanged = 0;
    const tocEntryChanges = [];
    for (const t of tocComparison) {
      tocEntryFieldsChanged += (t.changedPageCount || 0);
      for (const c of (t.changedPages || [])) {
        if (tocEntryChanges.length < 40) {
          tocEntryChanges.push({
            kind: "PAGEREF（目录条目）",
            code: `TablesOfContents.Item(${t.index})`,
            resultBefore: c.pageBefore === null ? "" : String(c.pageBefore),
            resultAfter: c.pageAfter === null ? "" : String(c.pageAfter),
            entry: c.entry
          });
        }
      }
    }
    const hfFieldCount = storyUpdates.reduce((sum, s) => sum + (Number(s.fieldCountAfter) || 0), 0);

    return {
      success: true,
      documentName: doc.Name,
      action: "update",
      scope: scopeVal,
      sectionIndex: scopeVal === "section" ? Number(sectionIndex) : undefined,
      tocMode: mode,
      elapsedMs: Date.now() - startedAt,
      fields: {
        countBefore: fieldsBefore.total,
        countAfter: fieldsAfter.total,
        // 宿主 Fields.Update() 不返回计数，按"更新后文档里实际存在的域"计（正文 + 页眉页脚 story）
        updatedFields: fieldsAfter.total + hfFieldCount,
        bodyFieldsAfter: fieldsAfter.total,
        headerFooterFieldsUpdated: hfFieldCount,
        comparedByCode: diff.compared,
        changedByCode: diff.changed,
        tocEntryFieldsChanged: tocEntryFieldsChanged,
        changedFields: diff.changed + tocEntryFieldsChanged,
        changes: diff.changes.concat(tocEntryChanges),
        changesTruncated: diff.changesTruncated || tocEntryFieldsChanged > tocEntryChanges.length
      },
      bodyUpdate: bodyUpdate,
      tocUpdate: tocUpdate,
      stories: storyUpdates,
      tablesOfContents: {
        countBefore: tocBefore.count,
        countAfter: tocAfter.count,
        pageNumbersChanged: pageNumbersChanged,
        comparison: tocComparison
      },
      warnings: warnings,
      errors: errors.length ? errors : undefined,
      hostLimitations: [
        "doc.Fields.Update() 的返回值在本宿主实测恒为 0，不能当'更新了几个域'的计数；本工具用更新前后逐域结果快照自己统计 updatedFields / changedFields。",
        "doc.Fields **不含页眉/页脚 story 里的域**，所以页眉页脚域由本工具逐节 Headers/Footers.Item(k).Range.Fields.Update() 单独更新（includeHeadersFooters=false 可关闭）。",
        "域更新无弹窗、无交互（实测 63 个域约 100ms），但目录页码依赖文档已完成分页：刚大批量改完内容时页码可能滞后，建议稍后重跑一次本工具。",
        "交叉引用只能按书签（宿主无 CrossReference 对象）；用 action='insert_cross_reference' 可先建书签再插 REF/PAGEREF。"
      ],
      message:
        `已更新 [${doc.Name}] 的域（scope=${scopeVal}，tocMode=${mode}）：更新后正文域 ${fieldsAfter.total} 个、页眉页脚域 ${hfFieldCount} 个，` +
        `其中 ${diff.changed + tocEntryFieldsChanged} 个结果发生变化（按域码配对 ${diff.changed} 个 + 目录条目域 ${tocEntryFieldsChanged} 个）；` +
        `${tocAfter.count} 个目录，页码${pageNumbersChanged ? "**已变化**" : "未变化"}` +
        (errors.length ? `；有 ${errors.length} 条错误，见 errors` : "") +
        (warnings.length ? `；有 ${warnings.length} 条告警，见 warnings` : "")
    };
  }

  // ===========================================================================
  // CAP-05：Word 内容控件全类型（RPC: word_manage_content_controls）
  //
  // 真机实测（WPS for Mac 12.x，2026-09-22，全部在临时新建文档上做，原文档只读）：
  //   - `doc.ContentControls.Add(Type, Range)` 可用；实测**可创建**：richText(0)、plainText(1)、
  //     picture(2)、comboBox(3)、dropdownList(4)、buildingBlockGallery(5)、date(6)、checkBox(8)。
  //   - **不支持**：group(7)、repeatingSection(9) —— `Add` 返回 null（不抛错），本工具显式报错。
  //   - **没有 `ListItems`**：`cc.ListItems` 实测 `undefined`（Word 桌面版的 ListItems 在 WPS for Mac 上不存在），
  //     下拉/组合框的选项要用 `cc.DropdownListEntries.Add(text, value)`；`Item(i).Delete()` 可删。
  //     新建的下拉/组合框自带一条**占位条目**（Text=占位文案、Value=""），本工具会先删掉它再写调用方的选项。
  //   - 值写入：plainText/richText/comboBox → `cc.Range.Text`；checkBox → `cc.Checked`（读回 Range.Text 是 ☐/☒）；
  //     date → `cc.Range.Text` + `cc.DateDisplayFormat` + `cc.DateDisplayLocale`；dropdownList → 必须
  //     `DropdownListEntries.Item(k).Select()`，**给 Range.Text 赋一个不在选项里的值是静默无效的**（实测无报错、值不变），
  //     本工具因此会核对选项并如实告警。
  //   - 占位符只能用 `cc.SetPlaceholderText(undefined, undefined, text)`；直接写 `cc.PlaceholderText.Text` 实测**无效**。
  //   - 嵌套：目标范围已在另一个内容控件内时 `Add` 返回 null（本宿主不支持嵌套控件），本工具显式报错。
  //   - 包裹语义：plainText/date 等把原范围**包起来**（文本保留）；dropdown/checkbox 是在原范围**之前**
  //     插入控件并留下原文本，本工具会把残留文本删掉（否则文档里会多出一份"标记原文"）。
  //   - `ContentControls.Item("标题")` 按名字取实测返回 null（无按 Title/Tag 查找），本工具改为遍历匹配。
  // ===========================================================================

  const WORD_CC_TYPES = {
    richText: 0, plainText: 1, picture: 2, comboBox: 3, dropdownList: 4,
    buildingBlockGallery: 5, date: 6, group: 7, checkBox: 8, repeatingSection: 9
  };
  const WORD_CC_TYPE_NAMES = {};
  for (const k in WORD_CC_TYPES) WORD_CC_TYPE_NAMES[WORD_CC_TYPES[k]] = k;
  // 实测不支持创建的类型（Add 返回 null）
  const WORD_CC_UNSUPPORTED = {
    group: "本宿主 WPS for Mac 实测 ContentControls.Add(7, range) 返回 null（不抛错），分组控件无法创建。",
    repeatingSection: "本宿主 WPS for Mac 实测 ContentControls.Add(9, range) 返回 null（不抛错），重复节控件无法创建。"
  };
  const WORD_CC_LIMITED = {
    picture: "可创建（Type=2），但本宿主没有图片填充通路，控件内容只是占位符号，不具备可填写语义。",
    buildingBlockGallery: "可创建（Type=5），但本宿主无法写入构建基块，控件内容只是占位文案，不具备可填写语义。"
  };

  function wordCcTypeNumber(type) {
    if (typeof type === "number" && WORD_CC_TYPE_NAMES[type] !== undefined) return type;
    const key = String(type === undefined || type === null ? "" : type).trim();
    if (WORD_CC_TYPES[key] !== undefined) return WORD_CC_TYPES[key];
    const asNum = Number(key);
    if (!isNaN(asNum) && WORD_CC_TYPE_NAMES[asNum] !== undefined) return asNum;
    throw new Error(`未知的内容控件类型: ${type}（可用：${Object.keys(WORD_CC_TYPES).join("、")}；其中 group、repeatingSection 本宿主不支持）`);
  }

  function wordCcEntries(cc) {
    const arr = [];
    try {
      const col = cc.DropdownListEntries;
      if (!col) return arr;
      for (let i = 1; i <= col.Count; i++) {
        const e = col.Item(i);
        arr.push({
          text: (() => { try { return e.Text; } catch (x) { return undefined; } })(),
          value: (() => { try { return e.Value; } catch (x) { return undefined; } })()
        });
      }
    } catch (e) {}
    return arr;
  }

  /** 单个内容控件的读回值（按类型取"当前值"这一项真正的语义）。 */
  function wordCcValue(cc, type) {
    if (type === 8) {
      return {
        checked: (() => { try { return Boolean(cc.Checked); } catch (e) { return undefined; } })(),
        text: (() => { try { return (cc.Range.Text || "").replace(/[\r\n\x07]/g, ""); } catch (e) { return undefined; } })()
      };
    }
    const v = {
      text: (() => { try { return (cc.Range.Text || "").replace(/[\r\n\x07]/g, ""); } catch (e) { return undefined; } })()
    };
    if (type === 6) {
      v.dateDisplayFormat = (() => { try { return cc.DateDisplayFormat; } catch (e) { return undefined; } })();
      v.dateDisplayLocale = (() => { try { return cc.DateDisplayLocale; } catch (e) { return undefined; } })();
    }
    if (type === 3 || type === 4) v.entries = wordCcEntries(cc);
    return v;
  }

  function wordCcDescribe(cc, index) {
    const type = (() => { try { return Number(cc.Type); } catch (e) { return undefined; } })();
    return {
      index: index,
      type: type,
      typeName: WORD_CC_TYPE_NAMES[type] || String(type),
      title: (() => { try { return cc.Title; } catch (e) { return undefined; } })(),
      tag: (() => { try { return cc.Tag; } catch (e) { return undefined; } })(),
      showingPlaceholder: (() => { try { return Boolean(cc.ShowingPlaceholderText); } catch (e) { return undefined; } })(),
      lockContentControl: (() => { try { return Boolean(cc.LockContentControl); } catch (e) { return undefined; } })(),
      lockContents: (() => { try { return Boolean(cc.LockContents); } catch (e) { return undefined; } })(),
      value: wordCcValue(cc, type),
      range: (() => { try { return { start: cc.Range.Start, end: cc.Range.End }; } catch (e) { return null; } })()
    };
  }

  function wordCcList(doc, limit) {
    const max = Number(limit) || 100;
    let count = 0;
    try { count = doc.ContentControls.Count; } catch (e) { return { count: 0, list: [], error: e.message }; }
    const list = [];
    for (let i = 1; i <= Math.min(count, max); i++) {
      try { list.push(wordCcDescribe(doc.ContentControls.Item(i), i)); }
      catch (e) { list.push({ index: i, error: e.message }); }
    }
    return { count: count, list: list, truncated: count > max };
  }

  /** 按 index / tag / title 定位控件；返回 {cc, index}。 */
  function wordCcLocate(doc, params) {
    const { index, tag, title } = params || {};
    const total = doc.ContentControls.Count;
    if (index !== undefined && index !== null && index !== "") {
      const i = Number(index);
      if (!i || i < 1 || i > total) throw new Error(`内容控件索引越界：收到 ${index}，当前共 ${total} 个内容控件`);
      return { cc: doc.ContentControls.Item(i), index: i };
    }
    const matches = [];
    for (let i = 1; i <= total; i++) {
      const c = doc.ContentControls.Item(i);
      let cTag = "", cTitle = "";
      try { cTag = String(c.Tag === undefined || c.Tag === null ? "" : c.Tag); } catch (e) {}
      try { cTitle = String(c.Title === undefined || c.Title === null ? "" : c.Title); } catch (e) {}
      if (tag !== undefined && tag !== null && tag !== "" && cTag === String(tag)) matches.push(i);
      else if (title !== undefined && title !== null && title !== "" && cTitle === String(title)) matches.push(i);
    }
    if (matches.length === 0) {
      throw new Error(`没有匹配的内容控件（tag=${tag === undefined ? "-" : tag}, title=${title === undefined ? "-" : title}）；当前共 ${total} 个。宿主 ContentControls.Item("名字") 实测不可用，本工具按 Title/Tag 遍历匹配。`);
    }
    return { cc: doc.ContentControls.Item(matches[0]), index: matches[0], matchedIndexes: matches };
  }

  function wordCcApplyListItems(cc, listItems, warnings) {
    const col = cc.DropdownListEntries;
    if (!col) { warnings.push("本宿主该控件没有 DropdownListEntries，选项未写入"); return; }
    // 清掉宿主自带的占位条目（Value 为空串），否则它会成为"第 1 个选项"
    if (col.Count >= 1) {
      try {
        const first = col.Item(1);
        const firstValue = first.Value === undefined || first.Value === null ? "" : String(first.Value);
        if (firstValue === "") first.Delete();
      } catch (e) { warnings.push(`占位选项清理失败：${e.message}`); }
    }
    for (const raw of listItems) {
      const text = typeof raw === "string" ? raw : String(raw && raw.text !== undefined ? raw.text : "");
      const value = typeof raw === "string" ? raw : (raw && raw.value !== undefined && raw.value !== null ? String(raw.value) : text);
      if (text === "") continue;
      try { col.Add(text, value); } catch (e) { warnings.push(`选项 [${text}] 写入失败：${e.message}`); }
    }
  }

  function wordManageContentControls(app, params) {
    const {
      documentName, action, type, index, tag, title, value, checked, dateDisplayFormat, dateDisplayLocale,
      listItems, placeholderText, lockContentControl, lockContents, location, markerText, markerOccurrence,
      paragraphIndex, initialText, clearListItems
    } = params || {};
    const doc = getWordDocument(app, documentName);
    const act = action || "list";
    const warnings = [];
    const hostSupport = {
      supportedTypes: ["richText", "plainText", "comboBox", "dropdownList", "date", "checkBox"],
      createOnlyWithoutContent: Object.keys(WORD_CC_LIMITED),
      unsupportedTypes: WORD_CC_UNSUPPORTED,
      listItemsNote: "本宿主没有 cc.ListItems（实测 undefined）；下拉/组合框选项请用 listItems 参数，落到宿主是 DropdownListEntries.Add(text, value)。",
      lookupNote: "宿主 ContentControls.Item(str) 按名字取实测返回 null；本工具的 set_value / delete 用 index 或 tag / title 遍历匹配。"
    };

    if (act === "list") {
      const ccList = wordCcList(doc, params.maxControls);
      return {
        success: true, documentName: doc.Name, action: "list",
        contentControls: ccList,
        hostSupport: hostSupport,
        message: `[${doc.Name}] 共有 ${ccList.count} 个内容控件`
      };
    }

    if (act === "add") {
      const typeNum = wordCcTypeNumber(type);
      const typeName = WORD_CC_TYPE_NAMES[typeNum];
      if (WORD_CC_UNSUPPORTED[typeName]) {
        throw new Error(`内容控件类型 '${typeName}' 本宿主未实现：${WORD_CC_UNSUPPORTED[typeName]}请改用 ${hostSupport.supportedTypes.join("/")}。`);
      }
      if (typeNum === 7 || typeNum === 9) {
        throw new Error(`内容控件类型 '${typeName}' 本宿主未实现（Add 返回 null），请改用 ${hostSupport.supportedTypes.join("/")}。`);
      }
      if (WORD_CC_LIMITED[typeName]) warnings.push(`${typeName}：${WORD_CC_LIMITED[typeName]}`);

      const loc = location || (markerText ? "marker" : "end");
      let range = null;
      if (loc === "marker") {
        if (!markerText) throw new Error("location='marker' 必须提供 markerText（要被替换成内容控件的标记文本，例如 '____'）");
        const meta = { skipped: [] };
        range = wordFindAnchorRange(doc, String(markerText), markerOccurrence, meta);
        if (!range) {
          throw new Error(
            `未在文档中找到可用的标记文本 [${markerText}]（第 ${Number(markerOccurrence) || 1} 处），未创建任何内容控件。` +
            (meta.skipped.length
              ? ` 已跳过 ${meta.skipped.length} 处命中：` + meta.skipped.map(s => `[${s.text}] 在${s.reason}`).join("；")
              : "")
          );
        }
        if (meta.skipped.length) {
          warnings.push(
            `有 ${meta.skipped.length} 处标记命中被跳过（不能把内容控件做进目录或域结果里，那里会被域重建清掉）：` +
            meta.skipped.map(s => `[${s.text}] 在${s.reason}`).join("；")
          );
        }
      } else if (loc === "start") {
        range = doc.Range(0, 0);
      } else if (loc === "selection") {
        const wordApp = getWordApp() || app;
        try {
          const sel = wordApp && wordApp.Selection && wordApp.Selection.Range;
          if (sel) range = sel;
        } catch (e) {}
        if (!range) { warnings.push("当前没有可用选区，内容控件回退创建到文档末尾。"); }
      } else if (loc === "after_paragraph") {
        const idx = Number(paragraphIndex);
        if (!idx || idx < 1 || idx > doc.Paragraphs.Count) {
          throw new Error(`段落索引越界：第 ${paragraphIndex} 段不存在（当前共 ${doc.Paragraphs.Count} 段）`);
        }
        const end = doc.Paragraphs.Item(idx).Range.End - 1;
        range = doc.Range(end, end);
      }
      if (loc === "end" || !range) {
        const e = doc.Content.End > 1 ? doc.Content.End - 1 : 0;
        range = doc.Range(e, e);
        if (initialText) {
          const text = String(initialText);
          range.InsertAfter(text);
          range = doc.Range(e, e + text.length);
        }
      }

      const target = { start: range.Start, end: range.End, markerText: markerText === undefined ? null : String(markerText) };
      let cc = null;
      try { cc = doc.ContentControls.Add(typeNum, range); }
      catch (e) { throw new Error(`ContentControls.Add(${typeNum}) 抛错：${e.message}`); }
      if (cc === null || cc === undefined) {
        throw new Error(
          `ContentControls.Add(${typeNum}, range) 返回 null，未创建任何控件。真机实测两类原因：` +
          `(1) 目标范围已落在另一个内容控件内部——本宿主不支持嵌套内容控件；` +
          `(2) 该类型本宿主不支持（group=7、repeatingSection=9）。` +
          `当前文档已有 ${doc.ContentControls.Count} 个内容控件，可先用 action='list' 查看范围。`
        );
      }

      // 非包裹型控件（下拉/组合框/复选框）会在原范围之前插入内容并留下原文，删掉残留
      let leftoverDeleted = null;
      try {
        const ccEnd = cc.Range.End;
        if (ccEnd < target.end) {
          const rest = doc.Range(ccEnd, target.end);
          const txt = (rest.Text || "").replace(/[\r\n\x07]/g, "");
          rest.Delete();
          leftoverDeleted = txt;
        }
      } catch (e) { warnings.push(`残留标记文本清理失败：${e.message}`); }

      if (title !== undefined) { try { cc.Title = String(title); } catch (e) { warnings.push(`Title 写入失败：${e.message}`); } }
      if (tag !== undefined) { try { cc.Tag = String(tag); } catch (e) { warnings.push(`Tag 写入失败：${e.message}`); } }
      if (placeholderText !== undefined) {
        // 实测：直接写 cc.PlaceholderText.Text 无效，必须走 SetPlaceholderText(BuildingBlock, Range, Text)
        try { cc.SetPlaceholderText(undefined, undefined, String(placeholderText)); }
        catch (e) { warnings.push(`占位符写入失败：${e.message}`); }
      }
      if (lockContentControl !== undefined) { try { cc.LockContentControl = Boolean(lockContentControl); } catch (e) { warnings.push(`LockContentControl 写入失败：${e.message}`); } }
      if (lockContents !== undefined) { try { cc.LockContents = Boolean(lockContents); } catch (e) { warnings.push(`LockContents 写入失败：${e.message}`); } }
      if (dateDisplayFormat !== undefined && typeNum === 6) { try { cc.DateDisplayFormat = String(dateDisplayFormat); } catch (e) { warnings.push(`日期格式写入失败：${e.message}`); } }
      if (dateDisplayLocale !== undefined && typeNum === 6) { try { cc.DateDisplayLocale = Number(dateDisplayLocale); } catch (e) { warnings.push(`日期区域写入失败：${e.message}`); } }

      if ((typeNum === 3 || typeNum === 4) && Array.isArray(listItems) && listItems.length) {
        wordCcApplyListItems(cc, listItems, warnings);
      }

      let valueApplied = null;
      if (checked !== undefined && typeNum === 8) {
        try { cc.Checked = Boolean(checked); valueApplied = Boolean(cc.Checked); }
        catch (e) { warnings.push(`Checked 写入失败：${e.message}`); }
      }
      if (value !== undefined && value !== null && String(value) !== "") {
        const v = String(value);
        if (typeNum === 4) {
          let hit = false;
          try {
            const col = cc.DropdownListEntries;
            for (let i = 1; i <= col.Count; i++) {
              const e = col.Item(i);
              const t = e.Text === undefined ? "" : String(e.Text);
              const val = e.Value === undefined ? "" : String(e.Value);
              if (t === v || val === v) { e.Select(); hit = true; break; }
            }
          } catch (e) { warnings.push(`下拉选中失败：${e.message}`); }
          valueApplied = hit;
          if (!hit) warnings.push(`下拉值 [${v}] 不在选项列表里：宿主对 Range.Text 赋非法值是**静默无效**的（实测无报错、值不变），本次未写入该值。`);
        } else if (typeNum === 8) {
          warnings.push("checkBox 的值请用 checked 参数，value 已忽略。");
        } else {
          try { cc.Range.Text = v; valueApplied = (cc.Range.Text || "").replace(/[\r\n\x07]/g, ""); }
          catch (e) { warnings.push(`值写入失败：${e.message}`); }
        }
      }
      if (typeNum === 6 && value !== undefined && value !== null && String(value) !== "") {
        try { cc.Range.Text = String(value); } catch (e) {}
      }

      // 读回：按控件自身对象读（创建后下标会随文档顺序变化，不能按 Count 取）
      let created = null;
      try { created = wordCcDescribe(cc, null); } catch (e) { warnings.push(`创建后读回失败：${e.message}`); }
      const listAfter = wordCcList(doc, 100);
      // 定位本次创建控件的真实下标（按 range.start 匹配）
      let createdIndex = null;
      if (created && created.range) {
        for (const item of listAfter.list) {
          if (item.range && item.range.start === created.range.start) { createdIndex = item.index; break; }
        }
      }
      if (createdIndex !== null) created.index = createdIndex;

      const tagApplied = tag === undefined ? null : (created && created.tag === String(tag));
      const titleApplied = title === undefined ? null : (created && created.title === String(title));
      if (tagApplied === false) warnings.push(`Tag 写入后读回不一致（期望 [${tag}]，读回 [${created && created.tag}]）：宿主可能丢弃了该属性，请以本返回的读回值为准。`);
      if (titleApplied === false) warnings.push(`Title 写入后读回不一致（期望 [${title}]，读回 [${created && created.title}]）。`);

      return {
        success: true, documentName: doc.Name, action: "add",
        type: typeName, typeCode: typeNum, location: loc,
        targetRange: target,
        leftoverMarkerTextDeleted: leftoverDeleted,
        valueApplied: valueApplied,
        tagApplied: tagApplied,
        titleApplied: titleApplied,
        created: created,
        contentControls: { count: listAfter.count, list: listAfter.list },
        hostSupport: hostSupport,
        warnings: warnings,
        message: `已在 [${doc.Name}] 创建 ${typeName} 内容控件（${loc}${markerText ? "：" + markerText : ""}），文档现有 ${listAfter.count} 个内容控件`
      };
    }

    if (act === "set_value") {
      const found = wordCcLocate(doc, params);
      const cc = found.cc;
      const typeNum = (() => { try { return Number(cc.Type); } catch (e) { return undefined; } })();
      const typeName = WORD_CC_TYPE_NAMES[typeNum] || String(typeNum);
      const applied = {};

      if (title !== undefined) { try { cc.Title = String(title); applied.title = cc.Title; } catch (e) { warnings.push(`Title 写入失败：${e.message}`); } }
      if (tag !== undefined) { try { cc.Tag = String(tag); applied.tag = cc.Tag; } catch (e) { warnings.push(`Tag 写入失败：${e.message}`); } }
      if (placeholderText !== undefined) {
        try { cc.SetPlaceholderText(undefined, undefined, String(placeholderText)); applied.placeholderApplied = true; }
        catch (e) { warnings.push(`占位符写入失败：${e.message}`); }
      }
      if (lockContentControl !== undefined) { try { cc.LockContentControl = Boolean(lockContentControl); applied.lockContentControl = cc.LockContentControl; } catch (e) { warnings.push(`LockContentControl 写入失败：${e.message}`); } }
      if (lockContents !== undefined) { try { cc.LockContents = Boolean(lockContents); applied.lockContents = cc.LockContents; } catch (e) { warnings.push(`LockContents 写入失败：${e.message}`); } }
      if (dateDisplayFormat !== undefined && typeNum === 6) { try { cc.DateDisplayFormat = String(dateDisplayFormat); applied.dateDisplayFormat = cc.DateDisplayFormat; } catch (e) { warnings.push(`日期格式写入失败：${e.message}`); } }
      if (clearListItems && (typeNum === 3 || typeNum === 4)) {
        try {
          const col = cc.DropdownListEntries;
          while (col.Count > 0) col.Item(col.Count).Delete();
          applied.listItemsCleared = true;
        } catch (e) { warnings.push(`选项清空失败：${e.message}`); }
      }
      if (Array.isArray(listItems) && listItems.length && (typeNum === 3 || typeNum === 4)) {
        wordCcApplyListItems(cc, listItems, warnings);
        applied.listItemsApplied = listItems.length;
      }
      if (checked !== undefined) {
        if (typeNum === 8) { try { cc.Checked = Boolean(checked); applied.checked = Boolean(cc.Checked); } catch (e) { warnings.push(`Checked 写入失败：${e.message}`); } }
        else warnings.push(`控件类型是 ${typeName}，checked 参数已忽略（只有 checkBox 支持）。`);
      }
      if (value !== undefined && value !== null && String(value) !== "") {
        const v = String(value);
        if (typeNum === 8) {
          warnings.push("checkBox 的值请用 checked 参数，value 已忽略。");
        } else if (typeNum === 4) {
          let hit = false;
          try {
            const col = cc.DropdownListEntries;
            for (let i = 1; i <= col.Count; i++) {
              const e = col.Item(i);
              const t = e.Text === undefined ? "" : String(e.Text);
              const val = e.Value === undefined ? "" : String(e.Value);
              if (t === v || val === v) { e.Select(); hit = true; break; }
            }
          } catch (e) { warnings.push(`下拉选中失败：${e.message}`); }
          applied.value = hit ? v : null;
          if (!hit) warnings.push(`下拉值 [${v}] 不在选项列表里：宿主静默无效（实测无报错、值不变），本次未写入。`);
        } else {
          try { cc.Range.Text = v; applied.value = (cc.Range.Text || "").replace(/[\r\n\x07]/g, ""); }
          catch (e) { warnings.push(`值写入失败：${e.message}`); }
        }
      }

      const readBack = wordCcDescribe(cc, found.index);
      return {
        success: true, documentName: doc.Name, action: "set_value",
        type: typeName, matchedIndex: found.index,
        matchedIndexes: found.matchedIndexes,
        applied: applied,
        contentControl: readBack,
        hostSupport: hostSupport,
        warnings: warnings,
        message: `已更新 [${doc.Name}] 第 ${found.index} 个内容控件（${typeName}）`
      };
    }

    if (act === "delete") {
      if (params.all === true) {
        const before = (() => { try { return doc.ContentControls.Count; } catch (e) { return 0; } })();
        const deleted = [];
        for (let i = before; i >= 1; i--) {
          try { doc.ContentControls.Item(i).Delete(); deleted.push(i); }
          catch (e) { warnings.push(`删除第 ${i} 个控件失败：${e.message}`); }
        }
        return {
          success: true, documentName: doc.Name, action: "delete", deleteAll: true,
          deletedCount: deleted.length, countBefore: before,
          contentControls: wordCcList(doc, 100),
          warnings: warnings,
          message: `已删除 [${doc.Name}] 的 ${deleted.length} 个内容控件（原有 ${before} 个）`
        };
      }
      const found = wordCcLocate(doc, params);
      const before = wordCcDescribe(found.cc, found.index);
      found.cc.Delete();
      const after = wordCcList(doc, 100);
      return {
        success: true, documentName: doc.Name, action: "delete",
        deleted: before,
        contentControls: { count: after.count, list: after.list },
        hostSupport: hostSupport,
        warnings: warnings,
        message: `已删除 [${doc.Name}] 第 ${found.index} 个内容控件（${before.typeName}，删除后剩余 ${after.count} 个）`
      };
    }

    throw new Error(`未知的内容控件操作: ${action}（支持 list, add, set_value, delete）`);
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
    let addError = null;
    try {
      chartShape = slide.Shapes.AddChart(typeCode, left, top, width, height);
    } catch (e) {
      addError = e.message;
    }
    if (!chartShape) {
      try {
        chartShape = slide.Shapes.AddChart2(-1, typeCode, left, top, width, height);
      } catch (e) {
        addError = (addError ? addError + "；" : "") + e.message;
      }
    }
    // 本机 WPS 实测：`AddChart` / `AddChart2` 都是 function，但**返回 null 且不创建任何形状**（问题台账 ISS-80）。
    // 原代码不检查返回值，随后 `chartShape.Chart` 抛错并被包装成"图表已创建，但数据配置未完成"——
    // 让调用方以为图已经建出来了，实际什么都没建。
    if (!chartShape) {
      throw new Error(
        "本宿主未能创建 PPT 原生图表：AddChart/AddChart2 未返回图表对象" +
        "（已实测本机 WPS 上二者返回 null 且不创建形状）" +
        (addError ? `；宿主返回：${addError}` : "") +
        "。替代方案：用矢量形状自行绘制，或改用 WPS 表格的原生图表。"
      );
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
            // 写后读回：本机 WPS 上 `Values = [...]` **不抛错也不生效**（问题台账 ISS-75），
            // 图表会显示宿主默认数据而调用方毫无察觉。这里核对是否真的落上。
            let readBackRaw = null;
            try { readBackRaw = sObj.Values; } catch (e) { readBackRaw = null; }
            if (readBackRaw !== null && readBackRaw !== undefined) {
              let readBack = [];
              try {
                const n = typeof readBackRaw.Count === "number" ? readBackRaw.Count : readBackRaw.length;
                for (let k = 0; k < n; k++) readBack.push(readBackRaw[k]);
              } catch (e) { readBack = []; }
              const want = sData.values.map(Number);
              const got = readBack.map(Number);
              const same = got.length === want.length &&
                want.every((value, index) => !Number.isFinite(value) || !Number.isFinite(got[index]) || value === got[index]);
              if (!same) {
                throw new Error(
                  `第 ${i + 1} 个数据系列的 Values 未生效：写入 ${JSON.stringify(want).slice(0, 70)}，读回 ${JSON.stringify(got).slice(0, 70)}。` +
                  `宿主忽略了数组赋值（实测本机 WPS 的 ChartData.Workbook 为 null）；请改用矢量形状绘制图表。`
                );
              }
            }
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
      // 配置失败时把**半成品形状**删掉再抛错：否则会留下一张数据空白/错乱的图表（问题台账 ISS-81）。
      let cleaned = false;
      try {
        chartShape.Delete();
        cleaned = true;
      } catch (delErr) {}
      throw new Error(
        `PPT 原生图表已创建但配置失败${cleaned ? "（已删除半成品形状）" : "（半成品形状删除失败，请手动清理）"}：${e.message}`
      );
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

  // 幻灯片页码校验：宿主对越界页码会抛内部 JS 错误（`Cannot read properties of null (reading 'Delete')`），
  // 调用方看不出是哪一页越界（问题台账 ISS-79）。这里统一前置校验并给中文上下文。
  function pptRequireSlideIndex(pres, slideIndex, actionLabel) {
    const total = pres.Slides.Count;
    const idx = Number(slideIndex);
    if (!slideIndex || !Number.isFinite(idx) || Math.floor(idx) !== idx) {
      throw new Error(`${actionLabel} 需要提供整数 slideIndex（1-based）；当前文稿共 ${total} 页`);
    }
    if (total === 0) {
      throw new Error(`${actionLabel} 失败：当前文稿一页都没有，请先用 action='add' 新增幻灯片`);
    }
    if (idx < 1 || idx > total) {
      throw new Error(`${actionLabel} 的 slideIndex=${idx} 越界：当前文稿 [${pres.Name}] 共 ${total} 页，有效范围 1~${total}`);
    }
    return idx;
  }

  function pptSavePresentation(app, params) {
    const { presentationName, filePath, format } = params || {};
    const pres = getPptPresentation(app, presentationName);
    const fmt = String(format || "").toLowerCase();
    const target = filePath ? String(filePath) : null;

    if (target && (fmt === "pdf" || target.toLowerCase().endsWith(".pdf"))) {
      // ppSaveAsPDF = 32；ExportAsFixedFormat(Path, FixedFormatType=2 表示 PDF)
      try {
        pres.ExportAsFixedFormat(target, 2);
      } catch (e) {
        pres.SaveAs(target, 32);
      }
      return {
        success: true,
        presentationName: pres.Name,
        savedPath: target,
        format: "pdf",
        message: `演示文稿 [${pres.Name}] 已导出为 PDF: ${target}`
      };
    }

    if (target) {
      try {
        pres.SaveAs(target);
      } catch (e) {
        pres.SaveAs(target, 24);
      }
      return {
        success: true,
        presentationName: pres.Name,
        savedPath: target,
        format: "pptx",
        message: `演示文稿 [${pres.Name}] 已另存为: ${target}`
      };
    }

    // 未命名的演示文稿：先判 Path 再决定要不要调 Save。
    // 实测（WPS for Mac 12.0）：对未命名文稿调用 `pres.Save()` **不抛错**，可能弹出"另存为"对话框；
    // 在宿主里弹模态框会卡住后续 RPC，因此这里不盲目调用，直接要求 filePath。
    let savePath = null;
    try { savePath = pres.Path ? String(pres.Path) : null; } catch (e) { savePath = null; }
    if (!savePath) {
      return {
        success: false,
        presentationName: pres.Name,
        savedPath: null,
        hostError: "Presentation.Path 为空：该文稿尚未保存到磁盘",
        attempts: [{ action: "save", ok: false, hostError: "文稿没有文件路径，未调用 Presentation.Save（未命名文稿的 Save 可能弹另存为对话框并阻塞自动化）" }],
        message: `演示文稿 [${pres.Name}] 尚未保存到磁盘，无法原地保存；请提供 filePath 走另存为（例如 /Users/.../方案.pptx）`
      };
    }

    pres.Save();
    const verified = (() => { try { return Number(pres.Saved) !== 0; } catch (e) { return null; } })();
    const fullName = (() => { try { return pres.FullName || null; } catch (e) { return null; } })();
    return {
      success: true,
      presentationName: pres.Name,
      savedPath: fullName || (savePath.replace(/[\\/]+$/, "") + "/" + pres.Name),
      format: "pptx",
      hostError: verified === false ? "保存后 Presentation.Saved 读回 false，落盘结果未确认" : undefined,
      verifiedSavedFlag: verified,
      message: `演示文稿 [${pres.Name}] 已原地保存（保存标记 Saved=${verified}）`
    };
  }

  function pptManageSlides(app, params) {
    const { presentationName, action, slideIndex, targetIndex, layoutIndex, backgroundColor, filePath, format, isVisible } = params || {};

    // 新建演示文稿：不需要预先存在的目标文稿，先处理
    if (action === "new_presentation") {
      const pptApp = getPptApp() || app;
      if (!pptApp) throw new Error("WPS 演示 (PowerPoint) 未就绪");
      const created = pptApp.Presentations.Add();
      try { if (isVisible !== false && created.Application) created.Application.Visible = true; } catch (e) {}
      let savedPath = null;
      if (filePath) {
        try { created.SaveAs(String(filePath)); savedPath = String(filePath); } catch (e) { savedPath = null; }
      }
      return {
        success: true,
        presentationName: created.Name,
        slideCount: created.Slides ? created.Slides.Count : 0,
        savedPath: savedPath,
        appended: false,
        message: savedPath
          ? `已新建演示文稿 [${created.Name}] 并另存为 ${savedPath}`
          : `已新建演示文稿 [${created.Name}]（尚未保存到磁盘；后续写入请显式传 presentationName）`
      };
    }

    const pres = getPptPresentation(app, presentationName);

    switch (action) {
      case "save":
      case "save_as":
        return pptSavePresentation(app, { presentationName: presentationName, filePath: filePath, format: format });
      case "add": {
        const idx = slideIndex ? Number(slideIndex) : (pres.Slides.Count + 1);
        if (slideIndex && (idx < 1 || idx > pres.Slides.Count + 1)) {
          throw new Error(`add 的 slideIndex=${idx} 越界：当前共 ${pres.Slides.Count} 页，可在 1~${pres.Slides.Count + 1} 之间插入`);
        }
        const lIndex = Number(layoutIndex) || 12;
        const newSlide = pres.Slides.Add(idx, lIndex);
        return {
          success: true,
          presentationName: pres.Name,
          slideIndex: newSlide.SlideIndex,
          slideCount: pres.Slides.Count,
          appended: true,
          idempotent: false,
          message: `已在位置 ${newSlide.SlideIndex} 新增幻灯片（追加型操作，重复调用会继续新增）`
        };
      }
      case "delete": {
        const idx = pptRequireSlideIndex(pres, slideIndex, "delete");
        const slide = pres.Slides.Item(idx);
        slide.Delete();
        return {
          success: true,
          presentationName: pres.Name,
          deletedIndex: idx,
          slideCount: pres.Slides.Count,
          message: `已成功删除第 ${idx} 页幻灯片（当前剩 ${pres.Slides.Count} 页）`
        };
      }
      case "move": {
        const idx = pptRequireSlideIndex(pres, slideIndex, "move");
        const to = Number(targetIndex);
        if (!targetIndex || !Number.isFinite(to) || to < 1 || to > pres.Slides.Count) {
          throw new Error(`move 的 targetIndex=${targetIndex} 越界：有效范围 1~${pres.Slides.Count}`);
        }
        const slide = pres.Slides.Item(idx);
        slide.MoveTo(to);
        return { success: true, presentationName: pres.Name, from: idx, to: to, slideCount: pres.Slides.Count, message: `幻灯片已从第 ${idx} 页移动到第 ${to} 页` };
      }
      case "duplicate": {
        const idx = pptRequireSlideIndex(pres, slideIndex, "duplicate");
        const slide = pres.Slides.Item(idx);
        slide.Duplicate();
        return {
          success: true,
          presentationName: pres.Name,
          originalIndex: idx,
          slideCount: pres.Slides.Count,
          appended: true,
          idempotent: false,
          message: `已克隆第 ${idx} 页幻灯片（追加型操作，重复调用会继续克隆）`
        };
      }
      case "set_background": {
        const idx = pptRequireSlideIndex(pres, slideIndex, "set_background");
        if (!backgroundColor) throw new Error("set_background 操作必须提供 backgroundColor");
        const slide = pres.Slides.Item(idx);
        const total = pres.Slides.Count;

        // 读回背景色的辅助：宿主返回的 RGB 是整数，统一成 "R,G,B"。
        // 注意 Office/WPS 的 RGB 整数约定是 **r + g*256 + b*65536**（红在低字节），
        // 与 hexToPptColor 的编码方向一致；解码若按相反方向会把红蓝显示反（真机实测踩到）。
        const readRgb = (target) => {
          try {
            const rgb = Number(target.Background.Fill.ForeColor.RGB);
            if (!Number.isFinite(rgb)) return null;
            const r = rgb & 0xff, g = (rgb >> 8) & 0xff, b = (rgb >> 16) & 0xff;
            return r + "," + g + "," + b;
          } catch (e) { return null; }
        };
        const neighbourIndex = idx === 1 ? Math.min(2, total) : idx - 1;
        const neighbourBefore = neighbourIndex !== idx ? readRgb(pres.Slides.Item(neighbourIndex)) : null;

        // 关键：该页若仍"跟随母版背景"，`slide.Background` 可能指向母版对象，写入会**串改全部页**
        // （问题台账 ISS-87，受控复现：设第 1 页后第 2 页也变红）。先显式断开与母版的关联再写。
        try { slide.FollowMasterBackground = false; } catch (e) {}
        slide.Background.Fill.Solid();
        slide.Background.Fill.ForeColor.RGB = hexToPptColor(backgroundColor);

        // 写后校验：目标页确实变了，且相邻页**没有被串改**
        const applied = readRgb(slide);
        const neighbourAfter = neighbourIndex !== idx ? readRgb(pres.Slides.Item(neighbourIndex)) : null;
        if (neighbourBefore !== null && neighbourAfter !== null && neighbourBefore !== neighbourAfter) {
          throw new Error(
            `设置背景时串改了相邻页：第 ${neighbourIndex} 页背景由 ${neighbourBefore} 变成了 ${neighbourAfter}。` +
            `已尝试先断开 FollowMasterBackground；请检查该稿的母版/版式是否被直接修改。`
          );
        }
        return {
          success: true,
          presentationName: pres.Name,
          slideIndex: idx,
          backgroundColor,
          appliedRgb: applied,
          neighbourChecked: neighbourIndex !== idx ? { slideIndex: neighbourIndex, before: neighbourBefore, after: neighbourAfter } : null,
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
      // 版式与占位符读回（ISS-44）：宿主目前无法区分矩形/圆角/椭圆（typeCode 都是 1），
      // 但占位符（typeCode 14）与所在版式名可读，用于核对"是否套用了正确版式"。
      layout: (() => {
        try {
          return {
            name: slide.CustomLayout ? slide.CustomLayout.Name : undefined,
            layoutIndex: Number(slide.Layout),
            placeholderCount: slide.Shapes.Placeholders ? slide.Shapes.Placeholders.Count : shapes.filter(s => s.typeCode === 14).length,
            placeholders: (() => {
              try {
                const list = [];
                const phCount = slide.Shapes.Placeholders.Count;
                for (let p = 1; p <= phCount; p++) {
                  const ph = slide.Shapes.Placeholders.Item(p);
                  list.push({
                    shapeId: ph.Id,
                    name: ph.Name,
                    placeholderType: (() => { try { return Number(ph.PlaceholderFormat.Type); } catch (e) { return undefined; } })(),
                    hasText: Boolean(ph.HasTextFrame && ph.TextFrame.HasText),
                    text: (() => { try { return ph.TextFrame.HasText ? String(ph.TextFrame.TextRange.Text).slice(0, 200) : ""; } catch (e) { return undefined; } })()
                  });
                }
                return list;
              } catch (e) { return []; }
            })()
          };
        } catch (e) {
          return { error: e.message };
        }
      })(),
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

        // 指定了 row/column 就**只刷那一格**；都没给才整表重刷（保持原有行为）。
        // 原实现无条件遍历整表，`row`/`column` 被静默忽略——于是"给某一格上色"永远做不到
        // （调用方以为设置成功了，实际被整表重刷覆盖）。
        const onlyRow = Number.isFinite(Number(row)) ? Number(row) : null;
        const onlyCol = Number.isFinite(Number(column)) ? Number(column) : null;

        const paintCell = (r, c, isHeader) => {
          const cell = tbl.Cell(r, c);
          cell.Shape.Fill.Solid();
          if (fillColor) {
            cell.Shape.Fill.ForeColor.RGB = hexToPptColor(fillColor);
          } else {
            cell.Shape.Fill.ForeColor.RGB = hexToPptColor(isHeader ? hColor : "#FFFFFF");
          }
          if (cell.Shape.HasTextFrame && cell.Shape.TextFrame.HasText) {
            const tr = cell.Shape.TextFrame.TextRange;
            if (isHeader && !fillColor) {
              tr.Font.Bold = true;
              tr.Font.Color.RGB = hexToPptColor("#FFFFFF");
            }
            if (fontColor) tr.Font.Color.RGB = hexToPptColor(fontColor);
            if (fontBold !== undefined) tr.Font.Bold = Boolean(fontBold);
            if (Number.isFinite(Number(fontSize))) tr.Font.Size = Number(fontSize);
          }
        };

        if (onlyRow !== null || onlyCol !== null) {
          const rStart = onlyRow !== null ? onlyRow : 1;
          const rEnd = onlyRow !== null ? onlyRow : tbl.Rows.Count;
          const cStart = onlyCol !== null ? onlyCol : 1;
          const cEnd = onlyCol !== null ? onlyCol : tbl.Columns.Count;
          if (rStart < 1 || rStart > tbl.Rows.Count || rEnd > tbl.Rows.Count) {
            throw new Error(`行号超出范围：表格共 ${tbl.Rows.Count} 行`);
          }
          if (cStart < 1 || cStart > tbl.Columns.Count || cEnd > tbl.Columns.Count) {
            throw new Error(`列号超出范围：表格共 ${tbl.Columns.Count} 列`);
          }
          for (let r = rStart; r <= rEnd; r++) for (let c = cStart; c <= cEnd; c++) paintCell(r, c, r === 1);
        } else {
          for (let r = 1; r <= tbl.Rows.Count; r++) for (let c = 1; c <= tbl.Columns.Count; c++) paintCell(r, c, r === 1);
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
    const total = pres.Slides.Count;
    let idx = Number(slideIndex);
    if (!idx || isNaN(idx)) {
      try {
        if (pptApp.ActiveWindow && pptApp.ActiveWindow.Selection && pptApp.ActiveWindow.Selection.SlideRange) {
          idx = pptApp.ActiveWindow.Selection.SlideRange.SlideIndex;
        }
      } catch (e) {}
      if (!idx) idx = 1;
    }
    if (total === 0) {
      throw new Error(`capture_slide_preview 失败：演示文稿 [${pres.Name}] 一页都没有，没有可导出的幻灯片`);
    }
    if (Math.floor(idx) !== idx || idx < 1 || idx > total) {
      throw new Error(`capture_slide_preview 的 slideIndex=${slideIndex} 越界：当前文稿 [${pres.Name}] 共 ${total} 页，有效范围 1~${total}`);
    }
    const slide = pres.Slides.Item(idx);

    const primaryPath = params.outputPath ? String(params.outputPath) : null;
    if (!primaryPath) throw new Error("缺少 Bridge 指定的预览输出路径（outputPath）");
    // 同一 API 用脚本 `slide.Export(path,"PNG",1280,720)` 实测可用，但工具路径历史上 100% 报
    // "PPT 未生成预览"（问题台账 ISS-76/ISS-86）：桥接侧把"文件没落盘"折成了兜底文案，
    // 宿主原始错误被丢掉。这里改为**两条导出路径依次尝试**，并把每次尝试的宿主原始报错全部回传。
    const attempts = [];
    const tryExport = (path, withSize) => {
      try {
        if (withSize) slide.Export(path, "PNG", 1280, 720);
        else slide.Export(path, "PNG");
        attempts.push({ path: path, scale: withSize ? "1280x720" : "default", ok: true });
        return true;
      } catch (e) {
        attempts.push({ path: path, scale: withSize ? "1280x720" : "default", ok: false, hostError: String(e && e.message ? e.message : e) });
        return false;
      }
    };

    const exportedPrimary = tryExport(primaryPath, true) || tryExport(primaryPath, false);
    const exported = [primaryPath];

    // 可选第二落点：调用方显式指定的 outputPath（网关转发的临时路径为 userOutputPath，
    // 两者不同时一并写出，便于调用方自行取图）
    const altPath = params.userOutputPath ? String(params.userOutputPath) : null;
    if (altPath && altPath !== primaryPath) {
      tryExport(altPath, true) || tryExport(altPath, false);
      exported.push(altPath);
    }

    const hostErrors = attempts.filter(a => !a.ok).map(a => `${a.path}（${a.scale}）: ${a.hostError}`);
    // 加载项无文件系统访问，无法确认落盘；把原始返回完整回传，由桥接侧判定文件是否存在（ISS-86）
    return {
      success: exportedPrimary,
      presentationName: pres.Name,
      slideIndex: idx,
      slideCount: total,
      imagePath: primaryPath,
      exportedPaths: exported,
      hasImage: exportedPrimary,
      hostError: hostErrors.length ? hostErrors.join(" | ") : undefined,
      attempts: attempts,
      message: exportedPrimary
        ? `已导出第 ${idx} 页幻灯片预览图至: ${primaryPath}`
        : `第 ${idx} 页幻灯片导出未成功；宿主对 ${attempts.length} 次导出尝试的原始报错见 hostError 与 attempts`
    };
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
      // 心跳 register 与 connection.js 的首次 register 是**两处**发报文的地方，
      // 字段必须一致：漏了 buildFingerprint 会让桥接永远判不出"进程里跑的是哪一版"（ISS-59）。
      sendPacket({
        type: "register",
        client: host === "word" ? "wps-word-addon" : host === "ppt" ? "wps-ppt-addon" : "wps-et-addon",
        version: "2.1.0",
        buildFingerprint: typeof ADDON_BUILD_FINGERPRINT === "string" ? ADDON_BUILD_FINGERPRINT : null,
        summary: getWorkspaceSummary(getApp())
      });
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
