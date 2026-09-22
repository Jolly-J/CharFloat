// ── 模块: src/bootstrap.js — DOM 事件绑定与启动引导（必须最后拼接） ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1774-1808 行（原样搬迁，未改写）。
  function bindDomEvents() {
    const btnReconnect = document.getElementById("btnReconnect");
    if (btnReconnect) {
      btnReconnect.addEventListener("click", () => {
        window.location.reload(true);
      });
    }

    const btnToggleLog = document.getElementById("btnToggleLog");
    const logDrawer = document.getElementById("logDrawer");
    const logArrow = document.getElementById("logArrow");
    if (btnToggleLog && logDrawer) {
      btnToggleLog.addEventListener("click", () => {
        logDrawer.classList.toggle("collapsed");
        if (logArrow) {
          logArrow.style.transform = logDrawer.classList.contains("collapsed") ? "rotate(0deg)" : "rotate(180deg)";
        }
      });
    }

    const btnClearLog = document.getElementById("btnClearLog");
    const logArea = document.getElementById("logArea");
    if (btnClearLog && logArea) {
      btnClearLog.addEventListener("click", () => {
        logArea.innerText = "";
      });
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bindDomEvents);
  } else {
    bindDomEvents();
  }

