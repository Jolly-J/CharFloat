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
