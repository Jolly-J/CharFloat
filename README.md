# Office Agent Bridge

让 AI 连接本机正在打开的办公文档。优先面向 Excel / WPS 表格，提供结构化操作、连接诊断和单元格修改记录。

## 当前支持范围

| 宿主 | 实现 | 验证状态 |
| --- | --- | --- |
| macOS WPS | 加载项、表格工具、Word/PPT 基础工具 | 1.x 曾在本机使用；2.0 服务与模拟宿主测试通过，新加载项仍需实机回归 |
| Windows WPS | 加载项安装分支及同一工具协议 | 待 Windows 实机验收 |
| Windows Microsoft Excel | PowerShell COM 结构化表格适配 | 代码已实现，待真实 Office 验收 |
| macOS Microsoft Office | JXA 原生脚本与状态通道 | 结构化 Excel 适配暂不支持 |
| Microsoft Word / PowerPoint | 原生脚本，PowerPoint 预览 | 不与 WPS 结构化工具等同，待实机确认 |

PPT 适合基础读取、文本、形状、表格和页面操作。复杂图表、嵌入数据编辑、母版、动画和高保真复刻不作完整支持承诺。

## 运行方式

```text
桌面管理器 ─┐
stdio MCP ──┼─ 本机认证 HTTP / MCP ─ 独立 Bridge 后台
HTTP MCP ───┘                           ├─ WPS 加载项 WebSocket
                                      └─ Windows Office COM / macOS JXA
```

- 管理器启动后台，窗口关闭或管理器退出后后台仍可运行。
- stdio MCP 自动启动或复用同一个后台；退出一个 AI 客户端不停止其他连接。
- HTTP/SSE 配置只连接已有服务，不能自行启动。可从管理器启动，或设置安装版登录启动。
- “停止服务”主动断开所有 AI。不会关闭办公软件或替用户保存文档。
- 服务只监听 `127.0.0.1`，HTTP 和加载项都需要当前安装的凭据。

## 源码运行

需要 Node.js 22 LTS 或更高版本。

```sh
npm ci
npm run typecheck
npm run build
npm start
```

在概览中选择“安装 / 修复加载项”，再在合适时机重启对应 WPS 组件。安装器只合并本项目条目，保留其他插件与备份，不会自动结束 WPS。

在“AI 接入”勾选客户端后配置。支持生成 stdio 配置，不依赖另装 Node（安装版使用随应用交付的 Electron 运行时）。设置中可切换主题、检查状态和查看日志。

## CLI / Agent

```sh
node dist/bridge/cli.cjs --start
node dist/bridge/cli.cjs --status
node dist/bridge/cli.cjs --doctor
node dist/bridge/cli.cjs --repair-addon
node dist/bridge/cli.cjs --stop
```

无参数运行 `cli.cjs` 即为 stdio MCP。`--repair-addon` 仅用于明确要部署/修复加载项时。安装管理器不会在每次启动时重写 AI 配置。

源码配置工具默认只检查：

```sh
npm run setup
npm run setup -- --agent=cursor --skills
npm run setup -- --addon
```

配套主技能位于 [skills/office-agent-bridge](skills/office-agent-bridge/SKILL.md)，包含 Python 标准库客户端、启动说明、故障分类和能力边界。AI 接入后应先调用 `bridge_get_capabilities`。

统一表格工具使用 `excel_*`，必须传 `host: "wps"` 或 `"microsoft"`；旧 `wps_*` 名称继续仅操作 WPS。原生脚本属于高级通道，不能据此推断结构化工具或宿主 API 全部可用。

## 数据与限制

默认运行目录：macOS `~/.wps-bridge`；Windows `%LOCALAPPDATA%/WPSBridge`。

`WPS_BRIDGE_HOME`、`WPS_BRIDGE_PORT` 可配置隔离环境；默认端口 19890。`installation.json` 保存可执行程序、CLI 与资源路径，`token` 保存本机凭据。不要分享 token 或把本机服务暴露到公网。

- 单元格 `patch_cells` 保存值和公式快照，回滚前检查是否存在后续修改。
- 样式、图表、工作表结构、Word/PPT、原生脚本不具备统一回滚。
- 写入成功不代表已保存。超时或断线意味着结果可能未知，先读回，勿自动重放。
- Windows COM 需要同一桌面用户/权限会话，PowerShell 脚本运行策略须允许本地安装资源；程序不改变系统执行策略。
- macOS WPS 表格截图仍依赖 Swift/Cocoa 剪贴板提取；预览可能改变剪贴板或当前激活对象。
- Microsoft Excel COM 通道附着注册的运行实例；多实例工作簿不一定都可枚举，找不到指定文件时明确报错，不创建隐藏替代文档。

## 验证与打包

```sh
npm test
npm run dist
```

本机测试使用临时目录、随机端口和模拟宿主，不修改用户文档。CI 在 macOS / Windows 运行构建和协议测试；Windows 额外解析 PowerShell 脚本。CI 配置已提供，未将其写入等同于已在远端运行。

真实 Windows Office 验收：在安装 Excel 的 Windows 桌面会话执行以下命令。它创建自己的 Excel 实例与临时工作簿，不附着或关闭已有用户文档：

```powershell
powershell -NoProfile -File tests/windows-office.ps1 -RunOffice
```

WPS 和客户端验收步骤见 [验证矩阵](docs/validation.md)。构建成功和模拟宿主通过不等于 Windows 实机验收通过。

默认打包输出在 `release/`：macOS DMG / ZIP，Windows NSIS / portable。安装包签名与公证需要发布者凭据，源码不包含签名密钥。
