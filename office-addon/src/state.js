// ── 模块: src/state.js — 全局常量与跨模块共享状态 ──
// 拼接片段（非独立 ES 模块）：由 scripts/build-office-addon.mjs 按固定顺序拼入 IIFE；初始迁移自 taskpane.js 第 7-15 行（原样搬迁，未改写）。
  const ADDON_VERSION = "2.0.0";
  // 部署探针：区分"磁盘上已构建的新能力"与"Excel 当前跑的那份窗格代码"。
  // 任务窗格是宿主 WebView 里长期驻留的页面，改完源码不重开窗格就不会生效；
  // 这个探针让任意 host=microsoft 的读工具都能回答"窗格里有没有本轮新分支"。
  const ADDON_CAPABILITIES = { cap08Shapes: true, shapeMethods: ["add_shape", "list_shapes", "update_shape", "group_shapes", "ungroup_shapes", "set_shape_zorder", "get_active_shape", "export_shape_image", "shape_self_test", "insert_image", "reload"] };
  const DEFAULT_WS_URL = "wss://localhost:19891/office-addon";
  let ws = null;
  let reconnectTimer = null;
  let isConnected = false;
  let activeWorkbookName = "—";
  let activeSheetName = "—";
  let activeSelectionAddress = "—";

