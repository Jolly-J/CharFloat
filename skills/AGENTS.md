# 面向使用者的技能

## 文件地图

- [office-agent-bridge/SKILL.md](office-agent-bridge/SKILL.md)：共享上下文、任务执行与恢复。
- [office-agent-bridge-chart-style/SKILL.md](office-agent-bridge-chart-style/SKILL.md)：表格、图表与矢量。
- [office-agent-bridge-ppt-design/SKILL.md](office-agent-bridge-ppt-design/SKILL.md)：页面批次与视觉验收。
- [office-agent-bridge-word-batch-edit/SKILL.md](office-agent-bridge-word-batch-edit/SKILL.md)：结构定位与改稿。
- [office-agent-bridge/references/native-scripting.md](office-agent-bridge/references/native-scripting.md)：脚本对象模型与示例。
- [office-agent-bridge/scripts/bridge_client.py](office-agent-bridge/scripts/bridge_client.py)：本地 HTTP 辅助客户端。

## 定位与联动

按任务只读对应 skill。说明与工具不符时核对 [gateway](../src/bridge/gateway.ts)、[catalog](../src/bridge/catalog.ts) 及宿主实现。

**能力元数据源（P2.6）**：`catalog.getTools()` 与 `catalog.capabilities()` 是唯一权威来源，MCP 的 `bridge://capabilities` 资源、`bridge_get_capabilities` 工具、HTTP `/api/v1/capabilities` 三处都读它。skill 与官网的人工说明必须与它一致：**不要在文档里另写能力数量**，需要引用时改为描述性表述或标注"以 `bridge_get_capabilities` 为准"。数量表述由 `npm run check:claims` 自动核对；定性表述（宿主支持范围、验证程度）仍需人工核对。

## 修改边界

AGENTS.md 管开发，SKILL.md 管用户任务，不能混入发布技能。共享流程只写一份，子 skill 复用上下文。只读必要字段、按可恢复批次修改，避免逐对象往返。新增能力同步描述、示例和验证边界；不承诺跨宿主 API 等同。

## 验证

检查 YAML frontmatter、相对链接、示例与真实 schema；运行 `npm run check:agents`、`npm run check:claims`。复杂流程用代表性任务验证调用与结果，格式通过不代表执行效率已验证。

## 避坑

没有专用工具就拒绝绘图 → 工具清单不是原生 API 能力上限 → 只读探测后使用脚本 → 核验实际对象与预览。重复连接检查拖慢任务 → 每个子 skill 重复初始化 → 共享上下文，仅异常时刷新。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
