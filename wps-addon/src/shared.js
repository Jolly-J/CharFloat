  // ---------------------------------------------------------------------------
  // shared.js — 配置常量与运行态变量、日志/状态 UI/原生弹窗、宿主组件探测与文档定位、颜色换算、工作区摘要
  // 本文件是 addon-core.js 的构建片段：由 scripts/build-wps-addon.mjs 按固定顺序拼进外层 IIFE。
  // 文本原样搬迁，因此保留 2 空格基础缩进；请勿在此文件内写 import/export。
  // ---------------------------------------------------------------------------
/**
 * WPS Bridge - WPS 内部加载项核心运行时 (专业美学与多功能版)
 * 运行在 WPS Office 进程内部 (JSA 环境)
 */

  const ADDON_VERSION = "2.2.0";
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

  /**
   * ISS-126 残留目标锁的**自愈**。
   *
   * 场景：锁定的文档被关掉（或宿主重启）后，`lockedTargets.<组件>` 仍指向一个不存在的名字，
   * 于是本会话**后续所有带锁工具全部失败**，而报错只说"未找到目标"，不指向"该重新锁目标"。
   *
   * 两条判据必须同时成立才自愈：
   *   ① 目标名字在宿主里确实不存在；
   *   ② 这个名字是**锁隐式生效**的，不是调用方显式传进来的。
   *      —— 调用方显式传了错名字时**绝不改动他的锁**，只报错，否则会悄悄换掉他要操作的文档。
   *
   * 返回 true 表示"已自愈，调用方继续走活动文档分支"。
   */
  function healStaleLock(component, targetName, explicitlyPassed, openList) {
    if (explicitlyPassed) return false;          // 调用方显式传名：不碰锁
    if (targetName !== lockedTargets[component]) return false;
    lockedTargets[component] = null;             // 锁失效 → 清掉，避免后续调用继续撞墙
    try {
      if (typeof console !== "undefined" && console.warn) {
        console.warn(`[Bridge] 目标锁 [${component}=${targetName}] 指向的文档已不存在（当前打开: ${openList.join(", ") || "无"}），已自动解除该锁并回退到活动文档。`);
      }
    } catch (e) {}
    return true;
  }

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
      if (!healStaleLock("word", targetName, !!docName, openList)) {
        throw new Error(`未在 WPS 中找到目标 Word 文档 [${targetName}]。当前已打开: ${openList.join(", ") || "无"}`);
      }
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
      if (!healStaleLock("ppt", targetName, !!presName, openList)) {
        throw new Error(`未在 WPS 中找到目标演示文稿 [${targetName}]。当前已打开: ${openList.join(", ") || "无"}`);
      }
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
      // ISS-126：锁隐式生效且目标已不存在 → 自愈（清锁 + 回退活动工作簿）
      if (!healStaleLock("excel", targetName, !!workbookName, openList)) {
        throw new Error(`未在 WPS 中找到目标工作簿 [${targetName}]。当前已打开: ${openList.join(", ") || "无"}`);
      }
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
