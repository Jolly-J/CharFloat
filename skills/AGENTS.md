# 面向使用者的技能

## 文件地图

- [charfloat/SKILL.md](charfloat/SKILL.md)：共享上下文、任务执行与恢复。
- [charfloat-chart-style/SKILL.md](charfloat-chart-style/SKILL.md)：表格、图表与矢量。
- [charfloat-ppt-design/SKILL.md](charfloat-ppt-design/SKILL.md)：页面批次与视觉验收。
- [charfloat-word-batch-edit/SKILL.md](charfloat-word-batch-edit/SKILL.md)：结构定位与改稿。
- [charfloat/references/native-scripting.md](charfloat/references/native-scripting.md)：脚本变量绑定、WPS JS API 与 VBA 差异、Excel 构图与 Word 示例。
- [charfloat/references/enumeration.md](charfloat/references/enumeration.md)：枚举常量速查（对齐、形状、图表、线型、颜色、域类型），逐项标注证据等级。
- [charfloat/scripts/bridge_client.py](charfloat/scripts/bridge_client.py)：本地 HTTP 辅助客户端。

## 定位与联动

按任务只读对应 skill。说明与工具不符时核对 [gateway](../src/bridge/gateway.ts)、[catalog](../src/bridge/catalog.ts) 及宿主实现。

**能力元数据源（P2.6）**：`catalog.getTools()` 与 `catalog.capabilities()` 是唯一权威来源，MCP 的 `bridge://capabilities` 资源、`bridge_get_capabilities` 工具、HTTP `/api/v1/capabilities` 三处都读它。skill 与官网的人工说明必须与它一致：**不要在文档里另写能力数量**，需要引用时改为描述性表述或标注"以 `bridge_get_capabilities` 为准"。数量表述由 `npm run check:claims` 自动核对；定性表述（宿主支持范围、验证程度）仍需人工核对。

## 修改边界

AGENTS.md 管开发，SKILL.md 管用户任务，不能混入发布技能。共享流程只写一份，子 skill 复用上下文。只读必要字段、按可恢复批次修改，避免逐对象往返。新增能力同步描述、示例和验证边界；不承诺跨宿主 API 等同。

## 验证

检查 YAML frontmatter、相对链接、示例与真实 schema；运行 `npm run check:agents`、`npm run check:claims`。复杂流程用代表性任务验证调用与结果，格式通过不代表执行效率已验证。

## 避坑

没有专用工具就拒绝绘图 → 工具清单不是原生 API 能力上限 → 只读探测后使用脚本 → 核验实际对象与预览。重复连接检查拖慢任务 → 每个子 skill 重复初始化 → 共享上下文，仅异常时刷新。

脚本里按组件只有一个变量有值 → 文档写"`doc` 已自动绑定"，Excel 场景 `doc` 实为 `null` → 宿主按形参名逐个解析，找不到即 `null` 且不报错 → 表格用 `wb`、文字用 `doc`、演示用 `pres`，批次开头自检 → 实测 `doc.Worksheets` 抛 `Cannot read properties of null`（ISS-52）。

格式化枚举传字符串 → 不报错但被静默吞成默认值（`'center'` 落成左对齐 `-4131`）→ 宿主对枚举型属性不做类型校验 → 一律传整数并写完读回断言 → 取值见 [references/enumeration.md](charfloat/references/enumeration.md)（ISS-11）。

Word 排版后想截图自查 → 反复调 `wps_word_capture_preview`、`ExportAsFixedFormat` → 前者**未注册**（报"未知工具"），后者不落盘 → 改用脚本读回真实属性并如实报告视觉验证缺口 → [native-scripting.md](charfloat/references/native-scripting.md)（ISS-71）。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
