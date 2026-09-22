# 测试与验收

## 文件地图

- [release-protection.test.ts](release-protection.test.ts)：发布副本混淆后的 WPS/Office RPC 行为、构建指纹和泄露门禁；不修改开发加载项或真实文档。
- [addon.test.ts](addon.test.ts)：WPS 模拟 RPC。
- [failure-routing.test.ts](failure-routing.test.ts)：失败分类与跨通道回退故障注入。
- [contracts-boundary.test.ts](contracts-boundary.test.ts)：契约层边界与依赖回环守卫。
- [session-isolation.test.ts](session-isolation.test.ts)：两个 stdio 客户端的真实链路会话隔离与审计归属（模拟宿主，需先 `npm run build`）。
- [contract-consistency.test.ts](contract-consistency.test.ts)：两类契约漂移与审计归属断言（只覆盖网关分支存在，不覆盖宿主实现与参数转换）。
- [ppt-layout.test.ts](ppt-layout.test.ts)：页面缩放、容量及脚本参数。
- [win-icon.test.ts](win-icon.test.ts)：Windows 图标生成器的 PNG 解码/缩放/编码与 ICO 条目格式（含"小于 256 的条目必须是 DIB"的回归项）。
- [office.test.ts](office.test.ts)：Office.js 参数转换。
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

`npm test`；局部使用 `node --import tsx --test tests/<文件>.test.ts`，不要将尖括号占位命令原样执行。`lifecycle.test.ts` 与 `session-isolation.test.ts` 通过 stdio 拉起 `dist/bridge/cli.cjs`，**运行前必须先 `npm run build`**，否则测的是旧产物。Windows 验收按 [validation.md](../docs/validation.md)。

## 避坑

服务测试出现 listen EPERM → 沙箱拒绝本地监听而非业务断言失败 → 在授权环境重新运行 → 保留实际错误和结果。

只测"客户端退出不误停服务"就以为覆盖了多客户端 → 跨客户端的**会话状态**（目标锁、审计归属）会串而不被发现 → 进程内 `AsyncLocalStorage` 只在单进程生效，经 HTTP 转发到后台时必须显式携带 `sessionId`，漏传会让后台把所有 stdio 客户端归入默认会话 `http-local` → 用两个独立 stdio 客户端进程分别锁定再互查 → [session-isolation.test.ts](session-isolation.test.ts)（复现证据见 [evidence](../docs/acceptance/2.1.0-p0p1/evidence/p2.2-session-isolation-regression.txt)）。

改了 `cli.ts` 之后只跑 `node --import tsx --test` 验证 → 实际执行的是 `dist/` 里的旧 CLI 产物，改动没生效却"测试通过" → stdio 类测试拉起的是构建产物 → 先 `npm run build` 再跑 → 见本页"验证"一节。

浮点比例严格相等导致假失败 → 使用容差 → 参见 ppt-layout.test.ts。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
