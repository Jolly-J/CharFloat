# 辅助脚本与副作用

## 文件地图

- [dev.ts](dev.ts)：启动 Vite 和 Electron；启动前依次构建主进程与**加载项部署入口**（`build:main` + `build:addons`）。
- [inspect.ts](inspect.ts)：启动服务并探测宿主。
- [verify-all-mcp-tools.ts](verify-all-mcp-tools.ts)：访问真实 Microsoft 宿主的工具验证。
- [verify_all_tiers.ts](verify_all_tiers.ts)：真实 Bridge 集成操作。
- [build_beautified_sheet.ts](build_beautified_sheet.ts)：真实表格写入与排版。
- [create-icons.mjs](create-icons.mjs)：生成图标文件。
- [check-agents.mjs](check-agents.mjs)：只读检查协作导航。
- [snapshot-tools.ts](snapshot-tools.ts)：只读导出工具契约与能力快照（`npm run snapshot:tools`）。
- [check-capability-claims.ts](check-capability-claims.ts)：只读核对 skill、官网与**加载项页面**中的能力数量表述与元数据源（`npm run check:claims`）。
- [build-win-icon.mjs](build-win-icon.mjs)：由 `build/icon.png` 生成 Windows 用的**多尺寸** `build/icon.ico`（`npm run build:win-icon`）；纯 Node 实现（自带 PNG 解码/缩放/编码与 DIB 封装），写完自检格式，不通过不落盘。
- [build-wps-addon.mjs](build-wps-addon.mjs)：由 `wps-addon/src/**` 生成部署入口 `wps-addon/addon-core.js`（`npm run build:wps-addon`，`--check` 只校验）。
- [build-office-addon.mjs](build-office-addon.mjs)：由 `office-addon/src/**` 生成部署入口 `office-addon/public/taskpane.js`（`npm run build:office-addon`，`--check` 只校验）。

## 定位与联动

先看脚本入口、目标、连接地址和写操作，再决定能否运行。命令入口见 [package.json](../package.json)。

## 修改边界

文件名带 inspect/verify 不代表只读；实际宿主脚本不能作为普通单元测试执行。不要沿用示例里写死的端口、认证和目标。新增脚本说明输入、输出与副作用，避免打印凭据。

## 验证

导航检查用 `npm run check:agents`；能力声明一致性用 `npm run check:claims`（两者都只读）。加载项构建用 `npm run build:addons`（**会覆写部署入口**，属写操作但只写生成物）。其余按实际副作用选择隔离环境和验证方式。

## 避坑

Windows 包 exe 图标显示空白 → 只看"`RT_GROUP_ICON` 条目数够不够"会得出错误结论：**条目数量正确、图标画面正确，仍然可能不显示** → 已证实原因：Windows 只对 **256×256** 支持 PNG 内嵌条目，小于 256 的条目必须是 **BMP(DIB)**；全部用 PNG 时资源管理器的小/中/大图标视图拿不到可解码的条目 → 正确做法：<256 写 32bpp DIB（`BITMAPINFOHEADER` 40 字节 + 自下而上 BGRA + 1bpp AND 掩码），256 写 PNG，与原版 `electron.exe` 的资源结构一致 → 由 [build-win-icon.mjs](build-win-icon.mjs) 生成并在写完自检，`npm run dist` 已串联该步 → 校验方法：用 `resedit` 读 exe 的 `RT_ICON`，逐条判断是 DIB 还是 PNG（对照 [win-icon.test.ts](../tests/win-icon.test.ts)）。另注意 **PE 版本资源是 UTF-16LE**，`strings -a` 只找 ASCII 会漏判。

直接手改 `wps-addon/addon-core.js` 或 `office-addon/public/taskpane.js` → 下次构建即被覆盖，且源码与部署产物不一致 → 两者都是生成物（顶部有"请勿手改"） → 改 `*/src/**` 后跑对应构建 → `--check` 模式可检出漂移。

能力文案检查只看 `skills/` 与官网 → **加载项页面这类手写文件里的数量文案会漏网** → 已证实：`office-addon/public/taskpane.html` 长期显示"42 项能力已就绪"，与四个口径（路由 30 / 对外工具 91 / WPS 26 / Microsoft 27）都对不上，并随发布包出厂；生成器只产 `taskpane.js`，检查器既不扫该目录、模式也不含"`N 项…能力`" → 扫描范围加入 `office-addon/public` 与 `wps-addon`，模式加入 `(\d+)\s*项[^，。；、\n]{0,16}?能力`；文案里的宿主必须写明（写 `Microsoft Excel`，不能只写"Excel"），否则按"未指明宿主"报错 → 反向验证：把 27 改回 42 即 `exit 1`。

直接执行 verify 脚本 → 可能修改真实文件 → 先检查请求列表和目标 → 使用明确测试文档并读回，不能靠脚本名称判断安全性。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
