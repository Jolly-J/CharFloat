// ── 模块: src/state.js — 全局常量与跨模块共享状态 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 7-15 行（原样搬迁，未改写）。
  const ADDON_VERSION = "2.0.0";
  // 临时诊断（CAP-08 探测用，收尾删除）：记录 run_script 的到达与执行结果。
  const DIAG = { rpcSeen: 0, lastMethod: null, seen: {}, runScriptSeen: 0, runScriptOk: 0, runScriptError: null, lastAt: null, lastCodeLen: 0, bootAt: new Date().toISOString() };
  const DEFAULT_WS_URL = "wss://localhost:19891/office-addon";
  let ws = null;
  let reconnectTimer = null;
  let isConnected = false;
  let activeWorkbookName = "—";
  let activeSheetName = "—";
  let activeSelectionAddress = "—";

