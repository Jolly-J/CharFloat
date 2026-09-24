# Changelog

本项目的所有显著变更将记录在此文件中。

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/) 格式，
版本号采用 [Semantic Versioning](https://semver.org/lang/zh-CN/) 语义化版本规范。

## [Unreleased]

## [2.2.0] - 2026-09-24

### Added (新增)
- 增加豆包工作连接器配置指引、配置状态识别和桌面端快捷唤起。
- 增加 WorkBuddy MCP 自动授信和专属技能自动同步。
- 增加宣传片拍摄脚本、演示道具及 2.2.0 工具清单快照。

### Changed (变更)
- 品牌统一为「字浮 CharFloat」，更新桌面客户端、WPS 与 Microsoft Office 加载项的名称、图标及连接文案。
- MCP 服务名与客户端配置切换到 `charfloat`；技能目录由 `office-agent-bridge*` 更名为 `charfloat*`，安装流程清理旧名称并同步新技能。
- 官网更新品牌视觉、交互演示、使用案例和 Windows/macOS 下载入口。
- 发布包名称统一为 `字浮-CharFloat-2.2.0-<平台>-<架构>.zip`。

### Fixed (修复)
- 修复独立检查 Windows 发布包时未放行其明文 CLI 兜底文件的问题。
- 更新豆包与 WorkBuddy 配置流程及跨平台测试覆盖。

## [2.1.0] - 2026-09-22

### Added (新增)
- **GUI 控制台组件化重构**：
  - 新增 `TopBar`（控制台顶栏导航）
  - 新增 `TopologyView`（Agent 与 Office 实时连接拓扑视图）
  - 新增 `LiveWorkspaceCard`（实时文档工作区状态卡片）
  - 新增 `SafetyTimeMachine`（安全审计与历史版本时间机器）
  - 新增 `AgentHub`（Agent 接入管理中心）
  - 新增 `DiffModal`（文档改动差异对比弹窗）
  - 新增 `SettingsDrawer`（全局参数与插件配置抽屉）
- **官方宣传与演示网站 (`website/`)**：
  - 引入基于 Vite + React + Tailwind 的独立官网模块
  - 包含 `ArchitectureFlow`（架构演进图）、`ComparisonTable`（AI 工具对比表）、`LiveWorkspace`（实时表格展示）等交互组件
- **多平台品牌图标与托盘资源**：
  - 补充 Windows / macOS 托盘高清图标及 WPS 插件功能按钮图标

### Changed (变更)
- **重构插件自动化安装器 (`src/main/addon-installer.ts`)**：
  - 优化 WPS/Office Addon 的配置读取、静默注册与跨平台日志输出
- **增强 Bridge 运行时与 WebSocket 状态通信 (`src/bridge/`)**：
  - 完善长连接日志监控与实时数据流动推送

### Fixed (修复与测试)
- 扩展自动化单元测试 (`tests/platform.test.ts` & `tests/service.test.ts`)，提升跨平台兼容性保障
