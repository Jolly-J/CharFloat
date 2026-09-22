// ── 模块: src/bootstrap.js — DOM 事件绑定与启动引导（必须最后拼接） ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 1774-1808 行。

  /**
   * 真正重新加载任务窗格。
   *
   * 原实现用 `window.location.reload(true)`：`true` 是已废弃的 forceGet 参数，
   * Excel for Mac 的 WKWebView 里不保证重新拉取子资源（实测点了「刷新连接」后
   * 窗格仍跑旧代码，见 docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/16-cap-ms-shapes.md）。
   * 改成"清 query 再 reload"：地址变了就必然重新请求 taskpane.html 与 taskpane.js，
   * 不依赖 forceGet 的实现细节。
   */
  function reloadTaskPane() {
    try {
      const url = new URL(window.location.href);
      url.searchParams.set("_reload", String(Date.now()));
      window.location.replace(url.toString());
    } catch (e) {
      window.location.reload();
    }
  }

  function bindDomEvents() {
    const btnReconnect = document.getElementById("btnReconnect");
    if (btnReconnect) {
      btnReconnect.addEventListener("click", () => {
        reloadTaskPane();
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

    // CAP-08 形状能力自检：隐藏入口（地址带 #selftest 时自动跑，或在控制台调 window.cap08ShapeSelfTest()）。
    const btnSelfTest = document.getElementById("cap08SelfTest");
    if (btnSelfTest) {
      btnSelfTest.addEventListener("click", () => {
        window.cap08ShapeSelfTest({});
      });
    }
    window.cap08ShapeSelfTest = async function (options) {
      try {
        const result = await handleShapeSelfTest(options || {});
        console.log("[CAP-08] 形状自检结果", result);
        if (typeof log === "function") log(`CAP-08 形状自检：${result.passed}/${result.total} 通过${result.failed && result.failed.length ? "，失败：" + result.failed.join("；") : ""}`);
        return result;
      } catch (e) {
        console.error("[CAP-08] 形状自检执行失败", e);
        throw e;
      }
    };
    if (window.location.hash === "#selftest") {
      window.cap08ShapeSelfTest({});
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bindDomEvents);
  } else {
    bindDomEvents();
  }

