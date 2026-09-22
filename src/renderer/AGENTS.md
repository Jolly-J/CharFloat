# 桌面界面

## 文件地图

- [App.tsx](App.tsx)：状态与操作编排。
- [index.css](index.css)：公共样式。
- [components/AgentHub.tsx](components/AgentHub.tsx)：客户端接入。
- [components/LiveWorkspaceCard.tsx](components/LiveWorkspaceCard.tsx)：工作区显示。
- [components/SafetyTimeMachine.tsx](components/SafetyTimeMachine.tsx)：审计界面。
- [components/DiffModal.tsx](components/DiffModal.tsx)：差异详情。
- [components/PermissionFloat.tsx](components/PermissionFloat.tsx)：macOS 完全磁盘访问权限引导**浮窗内容**（置顶、可拖拽图标、设置直达、重试）。
- [components/SettingsDrawer.tsx](components/SettingsDrawer.tsx)：设置。
- [components/TopologyView.tsx](components/TopologyView.tsx)：连接拓扑。

## 定位与联动

从 App.tsx 找状态来源和事件，再进入对应组件；接口追到 [preload](../preload/index.ts)，不要直接在 UI 调宿主实现。

## 修改边界

区分服务在线、加载项连接、文件打开和操作成功。保留加载、失败和部分完成状态；不能用静态文案冒充实时能力。订阅需解绑。

## 验证

`npm run typecheck`、`npm run build:renderer`；受影响交互需实际界面或截图核验。

## 避坑

只看构建就声称交互完成 → IPC 与真实状态未被覆盖 → 实际触发操作并读回状态 → 区分静态验证与界面验收。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
