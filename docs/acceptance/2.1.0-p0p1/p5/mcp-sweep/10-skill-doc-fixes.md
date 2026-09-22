# 10 skill 文档补强（写入台账）

- 开始：2026-09-22 · 执行者：skill 文档补强员
- 输入证据：[issues.md](../../issues.md) §ISS-11/12/13/15/26/52/71/74、[01-inventory.md](01-inventory.md)、[03-word.md](03-word.md)、[04-sheet-advanced.md](04-sheet-advanced.md)、[README.md](README.md)（成品清单与三个成品的元素构成）
- 写区：`skills/**`；本文件与 `skills/**` 之外不改（`docs/**` 仅新增本台账，不改他人文件）
- 不含：`src/**`、`wps-addon/**`、`office-addon/**`、`tests/**`、其他 `docs/**`

## 逐条：编号 → 文件 → 补什么

| 编号 | 目标文件 | 补什么 | 证据来源 | 状态 |
|---|---|---|---|---|
| ISS-52 | `office-agent-bridge/SKILL.md`、`references/native-scripting.md` | 删掉"`doc` 已自动绑定"式表述；按组件写清绑定变量：Excel → `wb`（实测 `doc` 为 `null`，`doc.Worksheets` 报 `Cannot read properties of null`）、Word → `doc`、PPT → `pres`，其余为 `null`；附 `dispatch.js:283-318` 的形参顺序作为机制说明 | issues.md §ISS-52；`wps-addon/src/dispatch.js:283-318` | 待写 |
| ISS-12 | `references/native-scripting.md` | 新增一节「WPS JS API ≠ VBA 差异清单」：`TextFrame.TextRange` 为 `undefined`（须 `Characters()`）、`ParagraphFormat` 不存在、`ws.Cells(r,c)` 不是函数（须 `ws.Range("A1")` 地址） | issues.md §ISS-12/§ISS-26；native-scripting.md 现例（已用 `Characters()`） | 待写 |
| ISS-26 | 同上（并入差异清单） | `ws.Cells(r,c)` 不存在；`AddChart2` + `SetSourceData` 默认按行取系列，须显式 `PlotBy=2` | issues.md §ISS-26 | 待写 |
| ISS-11 | `references/native-scripting.md`、新增 `references/enumeration.md` | 明确"枚举必须传整数"：字符串 `'center'` 不报错但静默落在左对齐（读回 `-4131`），居中须 `-4108`；给出实测对齐取值与"读完再断言"的做法 | issues.md §ISS-11 | 待写 |
| ISS-13 | 新增 `references/enumeration.md` | 枚举速查表：Excel 单元格对齐、边框线型、RGB 颜色、条件格式规则类型与运算符；PPT 形状类型 / 文本框类型 / 图表类型 / 版式枚举；Word 边框索引、`Orientation`、`HeightRule`、域类型、样式类型。**逐项标注证据等级**，未实测的明确写"未实测" | issues.md §ISS-13；`wps-addon/src/{excel,ppt,word}.js`；`04-sheet-advanced.md`；`.scratch/ppt3/21-layoutindex-align.json` | 待写 |
| ISS-15 | `references/native-scripting.md` | 完整信息图/流程图范例（标题条 + 三栏分组框 + 流程带箭头 + 应用卡 + KPI 条 + 页脚），复用已有成品的元素构成；写清画布换算：默认列宽 48pt、行高 16pt，`A1:U42` = 1008×672pt；给出锚点 `Range.Left/Top` + 列宽/行高推算坐标的写法与批次返回结构 | issues.md §ISS-15；mcp-sweep/README.md（成品 134 元素构成）；`docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/shapes-infographic-final.png` | 待写 |
| ISS-74 | `references/native-scripting.md` | 新增 Word 脚本示例节：新建/应用段落与字符样式（判定读 `Style.NameLocal`）、分节 + 横向页 + 断开页眉链接、书签创建与范围读回、`Fields.Add` 的 PAGE/NUMPAGES/REF/PAGEREF（code 里不得重复写域名）、表格行高须先设 `HeightRule`、底纹是 RGB 不是 BGR、表内公式不支持 | issues.md §ISS-74；`03-word.md` §2/§3/§4/§5 | 待写 |
| ISS-71 | `office-agent-bridge-word-batch-edit/SKILL.md`、`office-agent-bridge/SKILL.md` 或 `references/native-scripting.md` | 如实写明 `wps_word_capture_preview` 调用报"未知工具"，Word 侧**当前没有可用的视觉预览通路**；替代做法：native 属性读回 + `ComputeStatistics(2)` 分页统计 + OOXML 解包/工具读回；`ExportAsFixedFormat` 不落盘，不能当预览通道 | issues.md §ISS-71；`03-word.md` §7.4/§7.5/§9；`src/bridge/gateway/word.ts`（处理器存在） | 待写 |
| 联动 | `skills/AGENTS.md` | 文件地图加入 `references/enumeration.md`；避坑补"字符串枚举静默失效"与"组件变量"两条 | skills/AGENTS.md 同步维护规则 | 待写 |

## 准确性口径（本轮遵守）

1. 表中事实均来自真实宿主实测；**代码可证**的（`wps-addon/src/**` 常量、`dispatch.js` 绑定形参）标注为"代码可证"，不冒充实测。
2. 本轮**未连接宿主**，不新增任何未实测的宿主行为断言；不确定项在文档里标"未实测"。
3. 不写"WPS 支持 X"这类泛化结论；只写"在 macOS WPS 12.1.28496 + 加载项 2.1.0 实测/未实测"。
4. 文档不写能力数量（`check:claims` 口径）。

## 收尾检查

- [ ] `npm run check:claims` → 0
- [ ] `npm run check:agents` → 0
- [ ] 相对链接可达（新增 reference 的链接在 SKILL.md / native-scripting.md / skills/AGENTS.md 三处一致）
- [ ] 三个子 skill 里"Word 预览不可用"的口径与 `03-word.md` 一致

## 未覆盖（留给后续）

- ISS-89 危险成员清单（深度反射致 WPS 崩溃）——本轮只写"避免盲目逐成员求值"的通用约束，**具体崩溃清单需安全扫法先产出**，未写入枚举表。
- 多元素构图范例**未经真实宿主复跑**（本轮无宿主权限），标注为"依据已有成品结构整理，未在本轮复跑"。
