# 桌面主进程与安装

## 文件地图

- [index.ts](index.ts)：窗口、托盘、IPC 与服务控制。
- [installer-engine.ts](installer-engine.ts)：客户端检测及 MCP/skill 安装。
- [addon-installer.ts](addon-installer.ts)：办公加载项部署；写入失败时返回结构化 `permissionIssue` 而非裸错误。
- [permissions.ts](permissions.ts)：macOS 完全磁盘访问权限引导（设置深链、当前运行模式该授权的 App 路径）。
- [permission-detect.ts](permission-detect.ts)：权限错误的**纯识别逻辑**，不依赖 electron，可单元测试。
- [permission-window.ts](permission-window.ts)：权限引导的**置顶浮窗**（复用 `index.html#permission-guide`，含 App 图标拖拽 `startDrag`）。

## 定位与联动

窗口行为从 index.ts 开始；客户端接入从 installer-engine.ts 开始；加载项状态从 addon-installer.ts 开始。IPC 联动 [preload](../preload/AGENTS.md) 与 [界面](../renderer/AGENTS.md)。

## 修改边界

后台不由窗口持有；退出 GUI 不等于停止服务。外部配置合并保留其他条目，解析失败不得覆盖原文件。读取状态不应顺便安装或写配置。安装路径依运行配置，避免写死开发机器路径。

## 验证

`node --import tsx --test tests/platform.test.ts tests/lifecycle.test.ts`；IPC 修改执行 typecheck/build，窗口交互还需实际 Electron 验证。

## 避坑

修复配置时全量覆盖 → 丢失其他客户端设置 → 使用合并与备份 → [platform.test.ts](../../tests/platform.test.ts) 验证损坏配置和无关条目保留。

安装加载项失败时只抛裸错误或只给一句"请去系统设置开启"→ 用户看不懂也点不动，且错误停在会自动消失的 toast 里 → **FDA 是 TCC 里唯一没有申请 API、也没有系统弹窗的类别**，系统不会替应用弹框，只能检测失败后主动引导 → 权限类错误统一经 `detectPermissionIssue` 转成结构化 `permissionIssue`，界面弹不自动消失的引导（打开设置面板 + 显示**当前运行模式**该授权的 App + 重试） → [permission-guide.test.ts](../../tests/permission-guide.test.ts)；界面交互仍需实际 Electron 验证。

引导要"拖动图标进设置列表"却做成应用内弹窗 → 弹窗被系统设置窗口盖住，用户无从下手 → 拖拽必须与设置面板**同时可见** → 用置顶独立 BrowserWindow（`alwaysOnTop: 'floating'`），复用同一渲染包按 `location.hash === '#permission-guide'` 渲染浮窗内容，不新增构建入口 → `permission-window.ts` + `PermissionFloat.tsx`。

用 `new URL(import.meta.url).pathname` 取模块目录 → **不解码 `%20`**，仓库/安装路径含空格时 preload 路径失效，`window.api` 直接不存在 → 渲染层若又用 `api?.` 可选链，整座桥失效会表现成"点了没反应"，排查成本极高 → 路径一律用 `path.dirname(fileURLToPath(import.meta.url))`（与 index.ts 一致）；渲染层的桥调用**不用**可选链，桥缺失时显式降级提示；主进程监听 `preload-error` 写日志 → 全仓 grep 同类写法为空；`permission-window.ts` 已修正并加 `preload-error` 诊断。

Electron 拖拽文件出窗口用 `mousedown` + `ipcRenderer.invoke` → **拖不动，且没有任何报错** → `startDrag` 必须在拖拽会话内调用；`invoke` 的异步往返返回时系统拖拽会话已结束 → 用 HTML5 `dragstart`（`draggable` 元素）触发，经 `ipcRenderer.send` 单向通知主进程调用 `startDrag`；图标用 `app.getFileIcon(appPath,{size:'large'})` 取真实 Finder 图标，不要自绘方块 → 全站唯一实现见 `permission-window.ts` 的 `startAppDrag` + `PermissionFloat.tsx` 的 `onDragStart`。

给新窗口设 `transparent: true` 又叠加 `titleBarStyle: 'hiddenInset'` → **点"重新部署/修复"后整个应用 SIGTRAP 闪退**，终端报 `bootstrap_look_up …MachPortRendezvousServer.1: Permission denied (1100)` / `No rendezvous client, terminating process` → 透明窗在新渲染进程里走特殊合成路径，与自定义标题栏组合会让该进程连不上主进程的 Mach rendezvous 服务，属**不可捕获的原生 abort**（try/catch 无效） → 浮窗只用保守选项：不设 `transparent`、不设 `titleBarStyle`、给普通 `backgroundColor`，靠 `alwaysOnTop` + `setAlwaysOnTop(true,'floating')` 保持置顶 → 见 `permission-window.ts` 的注释；此类改动**必须在真实桌面点一次**，不能只看构建。

dev 与打包版的授权对象不同 → 引导里显示错 App，用户拖了也没用 → `npm run dev` 跑的是 `node_modules/electron/dist/Electron.app`，打包版是 `Office Agent Bridge.app`，FDA 按 App 授权、两者互不相通；且 TCC 按代码签名记忆，重签后可能需重新授权 → 引导按 `app.isPackaged` 显示当前该授权的那个，并提示重启应用 → `appToAuthorize()`。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
