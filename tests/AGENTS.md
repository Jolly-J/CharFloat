# 测试与验收

> `docs/` 为维护者本地资料，不随 GitHub 仓库提供；下文本地资料路径仅供已有资料时参考，不是构建或使用前提。

## 文件地图

- [brand.test.ts](brand.test.ts)：品牌门禁的正反向夹具、历史例外与生成物检查；只操作临时目录。
- [version-policy.test.ts](version-policy.test.ts)：版本一致性门禁及漏改、缺失、重复声明、锁文件根版本、格式错误的反向夹具；只操作临时目录，不构建或访问宿主。
- [release-protection.test.ts](release-protection.test.ts)：发布副本混淆后的 WPS/Office RPC 行为、构建指纹和泄露门禁；不修改开发加载项或真实文档。
- [bytecode-exports.test.ts](bytecode-exports.test.ts)：用实际 Electron 运行时编译并加载字节码，验证 PDF/图表文件核验、缺失文件拒绝和宿主调用次数；只用临时文件和模拟导出，不连接办公软件。
- [addon.test.ts](addon.test.ts)：WPS 模拟 RPC。
- [wps-addon-modules.test.ts](wps-addon-modules.test.ts)：递归构建清单的合法多层目录及未登记、缺失、重复反向夹具；只操作临时目录。
- [wps-excel-pure.test.ts](wps-excel-pure.test.ts)：矩阵归一化与列名进位，可直接导入无宿主依赖的 `excel/pure.js`。
- [wps-excel-rpc.test.ts](wps-excel-rpc.test.ts)：压缩产物的 Excel 领域模拟 RPC、宿主写入失败、选区去重与注册/心跳指纹；可用 `OAB_TEST_WPS_ARTIFACT` 选择隔离产物，用 `CHARFLOAT_EXCEL_SNAPSHOT` 输出响应与宿主调用记录。
- [addon-lifecycle.test.ts](addon-lifecycle.test.ts)：WPS 侧栏与连接分离、单实例和关闭后的 RPC；Office 共享运行时、自动启动与失败降级。
- [failure-routing.test.ts](failure-routing.test.ts)：失败分类与跨通道回退故障注入。
- [contracts-boundary.test.ts](contracts-boundary.test.ts)：契约层边界与依赖回环守卫。
- [session-isolation.test.ts](session-isolation.test.ts)：两个 stdio 客户端的真实链路会话隔离与审计归属（模拟宿主，需先 `npm run build`）。
- [contract-consistency.test.ts](contract-consistency.test.ts)：两类契约漂移与审计归属断言（只覆盖网关分支存在，不覆盖宿主实现与参数转换）。
- [fixtures/ppt-style/](fixtures/ppt-style/)：固定SHA来源、带哈希的旧读取函数与工具契约夹具，测试不依赖Git历史；layout-failure.ts以假页面运行真实布局函数产生部分失败响应。
- [ppt-read-style.test.ts](ppt-read-style.test.ts)：默认响应兼容、样式读取失败、有界递归与实际网关参数转发。
- [bridge-client.test.ts](bridge-client.test.ts)、[bridge-client.test.py](bridge-client.test.py)：Python CLI 的假 HTTP、参数文件、会话隔离及 PNG 输出，无真实服务。
- [tool-protocol.test.ts](tool-protocol.test.ts)：注入传输/runtime 执行真实 HTTP/MCP 回调，验证错误经 JSON 到 Python 不重放；不启动服务。
- [tool-error-response.test.ts](tool-error-response.test.ts)：协议错误白名单和执行状态保守透传。
- [ppt-layout.test.ts](ppt-layout.test.ts)：页面缩放、容量及脚本参数。
- [win-icon.test.ts](win-icon.test.ts)：Windows 图标生成器的 PNG 解码/缩放/编码与 ICO 条目格式（含"小于 256 的条目必须是 DIB"的回归项）。
- [windows-bootstrap.test.ts](windows-bootstrap.test.ts)：用户目录、真实 ZIP/逐文件哈希、慢写入字节进度、全局合并、错误哈希不可计完成、原生开屏像素、诊断日志白名单/脱敏/上限与入口静态约束、越界与链接防护、未知文件保留、并发锁、快捷方式适配器、最低五秒阶段和启动环境；只操作临时目录，Windows Shell 与 PowerShell 启动入口仍需实机验收。
- [windows-bootstrap-electron.cjs](windows-bootstrap-electron.cjs)：显式用完整 Electron 主进程与真实 ZIP/manifest 验证物理 `app.asar`、逐文件核验、重复配置和全局进度合并；只写独立临时目录，用无副作用快捷方式适配器，30 秒截止，不启动真实应用。此项不能以 Electron Node 模式或普通 Node 替代，运行命令见双平台构建文档。
- [windows-bootstrap-sharing.ps1](windows-bootstrap-sharing.ps1)：由 Windows 条件单测启动，验证真实 Node 写句柄打开时共享读取、UTF-8 半行、去重、暂态失败保留与恢复；只访问测试临时文件，不改执行策略。非 Windows 明确跳过。
- [windows-bootstrap-ui.cjs](windows-bootstrap-ui.cjs)：完整 Electron 静态四态截图、关闭图标中心与标注/按钮断言；按指定 zoom 同比例放大窗口模拟显示缩放，记录实际 devicePixelRatio，不把模拟写成 Windows DPI 实测。无 preload、无安装操作，输出拒绝覆盖。
- [office.test.ts](office.test.ts)：Office.js 参数转换。
- [office-background-routing.test.ts](office-background-routing.test.ts)：多工作簿后台连接、精确路由、歧义拒绝、改名和 WPS/Office 通道隔离。
- [service.test.ts](service.test.ts)：HTTP/MCP、鉴权、审计。
- [lifecycle.test.ts](lifecycle.test.ts)：后台与客户端生命周期。
- [platform.test.ts](platform.test.ts)：安装配置、目标隔离、原生进程。
- [native-contracts.ps1](native-contracts.ps1)：原生脚本契约。
- [windows-office.ps1](windows-office.ps1)：Windows 实机测试。

## 定位与联动

按修改模块选文件，查测试夹具与断言后运行。完整 Node 回归使用根命令 npm test。

## 修改边界

临时目录、模拟宿主与测试服务须隔离，不附着用户文档。测试关注行为及风险，不只比对实现文本。浮点几何使用合理误差。真实办公验证先确认脚本创建和关闭的实例归属。

## 验证

`npm test`；局部使用 `node --import tsx --test tests/<文件>.test.ts`，不要将尖括号占位命令原样执行。`lifecycle.test.ts` 与 `session-isolation.test.ts` 通过 stdio 拉起 `dist/bridge/cli.cjs`，**运行前必须先 `npm run build`**，否则测的是旧产物。Windows 验收按 validation.md（本地 `docs/validation.md`）。

## 避坑

服务测试出现 listen EPERM → 沙箱拒绝本地监听而非业务断言失败 → 在授权环境重新运行 → 保留实际错误和结果。

只测"客户端退出不误停服务"就以为覆盖了多客户端 → 跨客户端的**会话状态**（目标锁、审计归属）会串而不被发现 → 进程内 `AsyncLocalStorage` 只在单进程生效，经 HTTP 转发到后台时必须显式携带 `sessionId`，漏传会让后台把所有 stdio 客户端归入默认会话 `http-local` → 用两个独立 stdio 客户端进程分别锁定再互查 → [session-isolation.test.ts](session-isolation.test.ts)（复现证据见 evidence（本地 `docs/acceptance/2.1.0-p0p1/evidence/p2.2-session-isolation-regression.txt`））。

改了 `cli.ts` 之后只跑 `node --import tsx --test` 验证 → 实际执行的是 `dist/` 里的旧 CLI 产物，改动没生效却"测试通过" → stdio 类测试拉起的是构建产物 → 先 `npm run build` 再跑 → 见本页"验证"一节。

浮点比例严格相等导致假失败 → 使用容差 → 参见 ppt-layout.test.ts。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
