// ── 模块: src/ui.js — 任务窗格 UI 渲染：运行日志、活动流、状态徽标 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 16-56 行（原样搬迁，未改写）。
  function log(msg, data) {
    const timestamp = new Date().toLocaleTimeString();
    const formatted = `[${timestamp}] ${msg}` + (data ? ` ${JSON.stringify(data)}` : "");
    console.log(formatted);
    const logEl = document.getElementById("logArea");
    if (logEl) {
      logEl.innerText = (formatted + "\n" + logEl.innerText).slice(0, 5000);
    }
  }

  function addActivityItem(actionTitle) {
    const feed = document.getElementById("activityFeed");
    if (!feed) return;
    const empty = feed.querySelector(".feed-empty");
    if (empty) empty.remove();
    const now = new Date().toLocaleTimeString("zh-CN", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
    const item = document.createElement("div");
    item.className = "feed-item";
    item.innerHTML = `<span class="feed-item-action">${actionTitle}</span><span class="feed-item-time">${now}</span>`;
    feed.prepend(item);
    while (feed.children.length > 8) {
      feed.removeChild(feed.lastChild);
    }
  }

  function updateStatusUI(connected, text) {
    const badge = document.getElementById("connectionBadge");
    const connectionText = document.getElementById("connectionText");
    const docTitle = document.getElementById("docTitle");
    const activeSheetEl = document.getElementById("activeSheet");
    const activeSelEl = document.getElementById("activeSelection");

    if (badge && connectionText) {
      badge.className = "status-badge " + (connected ? "connected" : text === "连接中…" ? "connecting" : "error");
      connectionText.innerText = connected ? "已就绪" : (text || "未连接");
    }
    if (docTitle) docTitle.innerText = activeWorkbookName;
    if (activeSheetEl) activeSheetEl.innerText = activeSheetName;
    if (activeSelEl) activeSelEl.innerText = activeSelectionAddress;
  }

