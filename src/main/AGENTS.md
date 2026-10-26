# 桌面主进程与安装

## 文件地图

- [index.ts](index.ts)：窗口、托盘、IPC、服务控制及更新自动检查/每版本一次弹窗；弹窗仅唤起并定位客户端更新区域。
- [update-notice-window.ts](update-notice-window.ts)：带品牌形象的独立更新提醒窗口及受控 IPC；只允许关闭或打开客户端更新区域。
- [desktop-service.ts](desktop-service.ts)：只持有本 GUI 启动的 Bridge 句柄，取消启动竞态并等待退出；不接管既有后台。
- [update-protocol.ts](update-protocol.ts)、[update-trust.ts](update-trust.ts)：签名清单/版本/源验证与构建时信任配置；正式源未配置时不访问网络。
- [update-service.ts](update-service.ts)：检查、后台下载、取消/重试、尺寸/哈希核对与安装历史；下载状态与客户端全局忙碌状态分离。
- [update-install.ts](update-install.ts)、[update-handoff.ts](update-handoff.ts)：私有签名交接、就绪后一次性授权、失败撤销、旧 GUI 退出等待与提交前消费；失败旧票据不可重用，不接受 renderer 指定路径。
- [update-mac.ts](update-mac.ts)、[update-helper.ts](update-helper.ts)：独立 CJS helper（Windows 从既有程序的专用进程运行，不进入安装向导）、Mac 安全解压、同盘替换/回滚与离线缺失路径恢复。
- Mac 和 Windows ZIP 流共用 [zip-stream.mjs](../../windows-bootstrap/zip-stream.mjs)，必须覆盖 STORE/DEFLATE、空条目及真实发布 ZIP，不能只用压缩小样本判断更新助手可用。
- [update-config.ts](update-config.ts)：Windows 更新后的当前品牌 MCP 路径迁移，保留自定义参数与禁用状态，备份配置。部署/故障流程见[更新指南](../../UPDATES.md)。
- [installer-engine.ts](installer-engine.ts)：客户端检测、状态判定与 MCP/skill 安装编排。
- [client-registry.ts](client-registry.ts)：AI 客户端注册表（国内/国外分类、配置格式、路径、检测路径、手动接入说明）；加客户端只改这一处。
- [mcp-adapters.ts](mcp-adapters.ts)：按 JSON/JSONC、Codex TOML、OpenCode 和 VS Code 格式合并及读回 MCP 配置；保留停用和无关条目。
- [config-file.ts](config-file.ts)、[dsh-mcp.ts](dsh-mcp.ts)：有界配置读写、并发修改检查、备份；DSH YAML 节点合并，不执行表达式。
- [skill-links.ts](skill-links.ts)：稳定托管技能源、旧副本备份、Mac 符号链接/Windows junction、读回与失败回退。
- [desktop-diagnostics.ts](desktop-diagnostics.ts)：后台离线也能运行的只读检查；区分加载项连接、配置保存与实际客户端加载。两端路径及依据见 [CLIENTS.md](../../CLIENTS.md)。
- [legacy-brand.ts](legacy-brand.ts)：集中维护旧配置/加载项的识别与迁移标识，不用于用户文案或新条目。
- [addon-installer.ts](addon-installer.ts)：办公加载项部署；Office 清单按内容核对更新与状态，兼容同版本品牌变更；写入失败时返回结构化 `permissionIssue` 而非裸错误。
- [addon-autoupgrade.ts](addon-autoupgrade.ts)：启动时对齐加载项构建——比对**构建指纹**（版本号在多次构建间不变，不可作依据），不一致则重新部署并触发加载项重载，免重启 WPS。
  部署检查完成后状态包携带 `addonStartupChecked`；renderer 据此重新读取环境，避免卡片停留在初始检测结果。
- [permissions.ts](permissions.ts)：macOS 完全磁盘访问权限引导（设置深链、当前运行模式该授权的 App 路径）。
- [permission-detect.ts](permission-detect.ts)：权限错误的**纯识别逻辑**，不依赖 electron，可单元测试。
- [permission-window.ts](permission-window.ts)：权限引导的**置顶浮窗**（复用 `index.html#permission-guide`，含 App 图标拖拽 `startDrag`）。

## 定位与联动

窗口行为从 index.ts 开始；客户端接入从 installer-engine.ts 开始（客户端清单见 client-registry.ts，写入格式见 mcp-adapters.ts）；加载项状态从 addon-installer.ts 开始。IPC 联动 [preload](../preload/AGENTS.md) 与 [界面](../renderer/AGENTS.md)。

## 修改边界

后台不由窗口持有；关闭窗口隐藏，明确退出等待本 GUI 拥有的后台及其辅助进程停止并暂停 MCP 自动拉起，既有独立后台保留。外部配置合并保留其他条目，解析失败不得覆盖原文件。读取状态不应顺便安装或写配置。安装路径依运行配置，避免写死开发机器路径。

## 验证

`node --import tsx --test tests/platform.test.ts tests/client-integration.test.ts tests/lifecycle.test.ts`；IPC 修改执行 typecheck/build，窗口交互还需实际 Electron 验证。

更新和退出追加 `tests/update.test.ts`、`tests/update-install.test.ts`、`tests/desktop-lifecycle.test.ts`（需先构建 CLI）；安装测试仅临时目录，真实 GUI/Windows 安装分别报告。新增 helper 构建入口必须同步 tsup 配置与发布白名单；不能把测试公钥植入正式包。

## 避坑

Windows 更新后进程/Bridge 正常但主窗口隐藏 → 给 GUI 继承后台 helper 的 `windowsHide: true` → Windows STARTUPINFO 会压掉第一次 ShowWindow → 只在后台进程保留隐藏标记，更新器启动新版/恢复旧版 GUI 均使用正常显示启动 → `tests/update-gui-launch.test.ts` 执行实际 launch 函数核对参数；最终必须以未点托盘/未强制 ShowWindow 的 Windows 实机画面验收。

Windows Electron Node 22 的 path stat.dev 为 0、fstat.dev 为实际卷号 → 直接要求 dev 严格相等误判同一包被换 → 实机身份/大小/哈希一致且独立验票复现 → 仅实际 Windows 的未知 dev 兼容，保留 inode、独立文件、大小、句柄哈希/mtime；POSIX 和非零卷号仍严格比较。验票前失败走受限 stderr，不向未验证 argv 目录写回执；父进程先读 nonce 匹配的 failed，已知退出的 ready 不授权，撤销保留具体 cause → update-install 身份矩阵、Windows 实文件和失败回执/死助手回归。

Windows 更新交接使用短目录名 → 把 `fs.realpathSync` 与 `fs.promises.realpath` 的不同结果当作目录链接 → 后者调用 native 解析并展开 8.3 名称 → 校验使用 `promisify(fs.realpath)`，保留真实 symlink/junction 拒绝；一次性消费使用 `wx` 排他创建，不能依赖允许替换目标的 rename → `tests/update-install.test.ts` 的短路径、链接父级和多进程竞争回归，Windows 另跑离线预检及真实升级。

修复配置时全量覆盖 → 丢失其他客户端设置 → 使用合并与备份 → [platform.test.ts](../../tests/platform.test.ts) 验证损坏配置和无关条目保留。

安装加载项失败时只抛裸错误或只给一句"请去系统设置开启"→ 用户看不懂也点不动，且错误停在会自动消失的 toast 里 → **FDA 是 TCC 里唯一没有申请 API、也没有系统弹窗的类别**，系统不会替应用弹框，只能检测失败后主动引导 → 权限类错误统一经 `detectPermissionIssue` 转成结构化 `permissionIssue`，界面弹不自动消失的引导（打开设置面板 + 显示**当前运行模式**该授权的 App + 重试） → [permission-guide.test.ts](../../tests/permission-guide.test.ts)；界面交互仍需实际 Electron 验证。

引导要"拖动图标进设置列表"却做成应用内弹窗 → 弹窗被系统设置窗口盖住，用户无从下手 → 拖拽必须与设置面板**同时可见** → 用置顶独立 BrowserWindow（`alwaysOnTop: 'floating'`），复用同一渲染包按 `location.hash === '#permission-guide'` 渲染浮窗内容，不新增构建入口 → `permission-window.ts` + `PermissionFloat.tsx`。

用 `new URL(import.meta.url).pathname` 取模块目录 → **不解码 `%20`**，仓库/安装路径含空格时 preload 路径失效，`window.api` 直接不存在 → 渲染层若又用 `api?.` 可选链，整座桥失效会表现成"点了没反应"，排查成本极高 → 路径一律用 `path.dirname(fileURLToPath(import.meta.url))`（与 index.ts 一致）；渲染层的桥调用**不用**可选链，桥缺失时显式降级提示；主进程监听 `preload-error` 写日志 → 全仓 grep 同类写法为空；`permission-window.ts` 已修正并加 `preload-error` 诊断。

Electron 拖拽文件出窗口用 `mousedown` + `ipcRenderer.invoke` → **拖不动，且没有任何报错** → `startDrag` 必须在拖拽会话内调用；`invoke` 的异步往返返回时系统拖拽会话已结束 → 用 HTML5 `dragstart`（`draggable` 元素）触发，经 `ipcRenderer.send` 单向通知主进程调用 `startDrag`；图标用 `app.getFileIcon(appPath,{size:'large'})` 取真实 Finder 图标，不要自绘方块 → 全站唯一实现见 `permission-window.ts` 的 `startAppDrag` + `PermissionFloat.tsx` 的 `onDragStart`。

给新窗口设 `transparent: true` 又叠加 `titleBarStyle: 'hiddenInset'` → **点"重新部署/修复"后整个应用 SIGTRAP 闪退**，终端报 `bootstrap_look_up …MachPortRendezvousServer.1: Permission denied (1100)` / `No rendezvous client, terminating process` → 透明窗在新渲染进程里走特殊合成路径，与自定义标题栏组合会让该进程连不上主进程的 Mach rendezvous 服务，属**不可捕获的原生 abort**（try/catch 无效） → 浮窗只用保守选项：不设 `transparent`、不设 `titleBarStyle`、给普通 `backgroundColor`，靠 `alwaysOnTop` + `setAlwaysOnTop(true,'floating')` 保持置顶 → 见 `permission-window.ts` 的注释；此类改动**必须在真实桌面点一次**，不能只看构建。

dev 与打包版的授权对象不同 → 引导里显示错 App，用户拖了也没用 → `npm run dev` 跑的是 `node_modules/electron/dist/Electron.app`，打包版是 `字浮 CharFloat.app`，FDA 按 App 授权、两者互不相通；且 TCC 按代码签名记忆，重签后可能需重新授权 → 引导按 `app.isPackaged` 显示当前该授权的那个，并提示重启应用 → `appToAuthorize()`。

只比 `manifest.xml` 的版本号判断"加载项是否需要升级" → **版本号 2.1.0 在多次构建之间不变**，"版本号相同"完全不能说明"跑的是同一份构建" → 装了旧构建也会被判成最新，AI 拿到的一直是旧代码 → 比**构建指纹**（产物头部注入的 `ADDON_BUILD_FINGERPRINT`） → `checkBuildFreshness()`；回归见 [addon-freshness.test.ts](../../tests/addon-freshness.test.ts)。

升级加载项后不触发重载 → 磁盘是新的、WPS 进程里跑的还是旧 JS（ISS-59），用户以为修好了 → 部署完调桥接的 `wps_reload_addon`；且该工单必须**逐个组件**发送（只发一次实测只送达一个组件，word/ppt 仍是旧构建） → 配合部署时写入 `index.html` 的 `<script src="./addon-core.js?v=<指纹>">`，重载才会真正取到新文件（否则 WPS 只重跑缓存） → `ensureAddonUpToDate()`。

部署失败仍去重载 → 重载只会再次加载旧文件，制造"已修复"的假象 → 安装器返回 `success: false` 时**不触发重载**并如实报告 → 同上测试的"部署失败时不触发重载"用例。

给各 AI 客户端套同一配置格式或把应用缓存目录当用户配置 → 文件写入成功但客户端不读取 → 不同厂商及不同入口的契约不同 → 按注册表的每个 `configTargets.kind` 适配并读回，官方依据与两端矩阵见 [CLIENTS.md](../../CLIENTS.md) → [client-integration.test.ts](../../tests/client-integration.test.ts)。

只看到 MCP 条目存在就显示连接成功，或直接改客户端审批记录 → 无效路径、停用和用户授权被忽略 → 配置保存与运行时加载是两层状态 → 检查当前命令/参数/环境、停用记录，保留策略；应用内接入明确引导，不猜全局路径，不代用户授权 → 同上测试与各客户端实际工具列表。

Windows 技能链接指向应用版本目录 → 更新后仍读取旧技能 → 版本目录不可变且旧目录保留 → 链接到用户数据下稳定托管源，Windows 使用 junction，旧副本先备份；更新仅刷新已管理源 → 同上链接与回退用例，Windows 实机单独验收。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
