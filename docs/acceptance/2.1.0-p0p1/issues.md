# 问题台账（总验收期间累积，逐条修复）

开始记录：2026-09-22 · 来源：P5 真实宿主业务矩阵 + 全方位 MCP 盘点 + 独立审计

**用法**：每条问题一个编号，记录证据、影响面、建议修法与验证方式。修复时把状态改成「修复中 → 已修待验 → 已修已验证」，并在"验证"列写清用什么证明的。**不删条目**；判定不修的写进"状态"并说明理由。

状态口径：`待修` / `修复中` / `已修待验` / `已修已验证` / `不修` / `撤销·非缺陷`（经原始证据复核，原现象**不成立**——不是宿主缺陷也不是产品缺陷，通常是验证方法本身出错；按"不删条目"保留原记录与撤稿原因）

> **本文件只管"现有能力做错了"。** "宿主能做、但 MCP 还没让 AI 用上"的**能力缺口**单独维护在
> [capability-backlog.md](capability-backlog.md)（能力补强专项台账，37 条，待排期）——两者不要混。

---

## 决策结论（2026-09-22，使用者确认）

| 决策 | 结论 | 影响 |
|---|---|---|
| ISS-01 / DP7 | **方案 A**：`formulas` 中的空项视为"不动该单元格" | 改 `wps-addon/src/excel.js` 的 `patchCells` + 回归测试 + 工具说明 |
| ISS-04 重打包 | **方案 B**：不重打，登记为已知差异 | 已交付三包 SHA 不变；包内文案待下次统一打包时修正 |
| "假成功"整类（ISS-17/29/38/39/42） | **方案 A**：写工具执行后**主动读回校验**，未生效即报错 | 需要改适配层，是本次最大的架构性修复 |
| ISS-09（D8/D9） | **方案 B**：要求处理 | 重写 WPS PPT 描述；让 `capabilities` 覆盖 Word/PPT |
| ISS-44/45（工具缺口） | **补读回工具** | 条件格式 / 数据验证 / 冻结窗格的读回能力；`clear_range` 按 DP3 再评估 |
| 测试工作表 | 保留 3 张 `sweep*` 成品表；**删除** `__MCP验收测试__` | 待第二轮子代理结束后执行 |
| 工具说明变更 | **允许**（属有意变更，非契约漂移） | 改完重新留档契约快照并列出差异 |
| 未提交改动 | **已提交**：`fbb0e83`（`main`，未推送） | 基线恢复实测 22/22 仍通过（锚点与 HEAD 解耦已验证） |
| DP2 签名/公证 | **暂缓**，后期再议 | 不影响内部验收 |

---

## 汇总

| 编号 | 标题 | 严重度 | 状态 | 归属 |
|---|---|---|---|---|
| ISS-01 | `excel_patch_cells` 混用 `values`+`formulas` 会静默清空数据 | **高** | **已修已验证** | 宿主实现 + 工具说明 |
| ISS-02 | `wps_get_workspace_summary` 的 `hasOpenWorkbook` 与真实打开状态不符 | **高** | **已修已验证** | 目标锁语义 + 工具说明 |
| ISS-03 | 参数约定不一致：多数 excel 工具要 `host`/`sheetName`，个别不要 | 中 | **已修已验证** | 工具契约 |
| ISS-04 | 已交付包内 `taskpane.html` 仍显示"42 项能力已就绪" | 中 | 源码已修、**包未重打** | 发布物 |
| ISS-05 | 能力文案检查漏掉加载项页面（扫描范围 + 模式双重漏网） | 中 | 已修已验证 | 检查脚本 |
| ISS-06 | P3.2/P3.3 加载项等价性证据（acorn 多重集、diff、20 条 RPC 序列）无留档 | 中 | **已修已验证**（三项重跑，数字全部对上） | 验收证据 |
| ISS-07 | 台账与证据文件存在多处数字漂移 | 低 | 已修已验证 | 文档 |
| ISS-08 | `macOS x64` 无依据、DP1 决定不扩大（保留登记） | 低 | 不修（DP1 已决议） | 发布矩阵 |
| ISS-09 | D8/D9 能力声明残留（WPS PPT 支持范围、capabilities 不能作为完整来源） | 中 | 待使用者裁决 | 能力声明 |
| ISS-10 | 9 条未确认项（实机可用性、部署版本一致性、超时阈值等） | 中 | **已修已验证**（9 条逐条定状态：6 已澄清 / 3 仍保留） | 验证缺口 |
| ISS-11 | 绘图对齐传字符串被**静默吞掉**成左对齐，不报错 | **高** | **已修已验证**（文档类） | 宿主 API 语义 + skill 说明 |
| ISS-12 | `TextFrame.TextRange` / `ParagraphFormat` 在 WPS JS API 中不存在，说明与 VBA 文档混用会直接中断 | 中高 | **已修已验证**（文档类） | skill 说明 |
| ISS-13 | `wps_inspect_api` 只反射方法名，**不返回枚举常量**，而绘图/格式化全靠枚举 | 中高 | **已修已验证**（文档类） | 工具能力 + skill 说明 |
| ISS-14 | `Shapes.Range([...]).Group()` 不报错但产出损坏的分组对象 | 中 | **已修已验证**（文档类，见 skill 差异清单） | 宿主 API 语义 |
| ISS-15 | `native-scripting.md` 缺多元素构图范例与画布换算 | 中 | **已修已验证**（文档类） | skill 说明 |
| ISS-16 | 多任务并发时 `save_workbook` 会把**整个工作簿**的内存状态落盘（含他人在途结果） | 中 | 台账已补实测场景 + 规避；**说明文案待工具说明批次** | 并发语义 + 说明 |
| ISS-17 | `add_chart` 的 `chartType` **静默降级**（`scatter`/`area`/`column`/`bar` 全变成柱状/条形） | **高** | **已修已验证** | 宿主实现 + 说明 |
| ISS-18 | `add_chart` 定位参数被静默忽略（4 张图叠在 360/40），而说明写"100% 完美的行级锁定" | **高** | **已修已验证** | 宿主实现 + 说明 |
| ISS-19 | `delete_chart` 的 `leftCell` 是**像素邻近批量匹配**，一次删掉 14 张图 | **高** | **已修已验证** | 宿主实现 + 说明 |
| ISS-20 | `update_chart` 在 WPS 不可用，但工具说明未标注 | 中 | **已修已验证**（契约快照已复核） | 工具说明 |
| ISS-21 | `patch_cells` 静默把字符串日期（`"2026-01"`）转成序列号 | 中 | **已修已验证**（文档类） | 宿主实现 + 说明 |
| ISS-22 | `format_cells` 的 `rowHeight` 不生效 | 低中 | **已修已验证** | 宿主实现 |
| ISS-23 | `get_charts` 的 `leftCell` 与实际位置不自洽，`detail` 不返回类型枚举名与颜色 | 中 | **已修已验证** | 宿主实现 |
| ISS-24 | `seriesColors` 对单系列图表是**逐点染色**，导致彩虹柱 | 中 | **已修已验证** | 宿主实现 + 说明 |
| ISS-25 | `capture_sheet_preview` 的参数优先级与返回字段未说明 | 低中 | **已修已验证**（契约快照已复核） | 工具说明 |
| ISS-26 | 脚本 API 差异：`ws.Cells(r,c)` 不存在；`AddChart2` 默认按行取系列须 `PlotBy=2` | 中 | **已修已验证**（文档类） | skill 说明 |
| ISS-27 | 缺少**选型指引**：27 对 `excel_*`/`wps_*` 完全同构，`tools/list` 里无法区分 | **高** | **已修已验证**（契约快照已复核） | 工具说明 + skill |
| ISS-28 | `wps_rollback` 说"原地恢复表格"，实际只覆盖 `patch_cells` 的值/公式 | **高** | **已修已验证**（契约快照已复核） | 工具说明 |
| ISS-29 | `search_cells` 声称能搜公式，实测 0 命中**却返回 `success:true`** | **高** | **已修已验证** | 宿主实现 + 说明 |
| ISS-30 | 审计族说明近乎空白（15–17 字 + 9 个参数全无说明） | **高** | **已修已验证**（契约快照已复核） | 工具说明 |
| ISS-31 | 11 个参数缺 JSON Schema `type`，1 个用非标准类型 | 中 | **已修已验证**（契约快照已复核） | 工具契约 |
| ISS-32 | 装配层正则误伤说明文本，出现**残句**（`office_execute_script`） | 中 | **已修已验证**（契约快照已复核） | 工具说明 |
| ISS-33 | "Windows Microsoft Excel 尚待实机验收"措辞把范围写窄，读者会以为 macOS 已验证 | 中 | **已修已验证**（契约快照已复核） | 工具说明 |
| ISS-34 | 枚举成员与参数说明不齐（如 `preset` 含 `academic` 却未解释） | 中 | **已修已验证**（契约快照已复核） | 工具说明 |
| ISS-35 | 17 个写工具**零必填参数**（连目标都不是必填） | 中 | **已修已验证**（说明层；schema 的 anyOf 强制待评估） | 工具契约 |
| ISS-36 | 说明营销化，6549 字符里约 13% 是样板套话，挤占有效信息 | 低中 | **已修已验证**（契约快照已复核） | 工具说明 |
| ISS-37 | 错误提示不给允许值清单，参数集不一致时尤其费轮次 | 低 | **已修已验证**（契约快照已复核） | 工具契约 |
| ISS-38 | `set_filter_and_sort` **假成功**：返回 `success` + `sortedRuleCount:1`，数据根本没排序 | **高** | **已修已验证** | 宿主实现 + 说明 |
| ISS-39 |  `auto_fit_columns` 不传 `columnRules` 时静默 no-op；`address` 参数被完全忽略 | 中高 | **已修已验证** | 宿主实现 + 说明 |
| ISS-40 |  筛选范围不受 `range` 约束，自动扩展到最后一个已用行，把合计行卷进筛选区 | 中高 | **已修已验证** | 宿主实现 + 说明 |
| ISS-41 | `freeze_panes` 的 `freezeColumnIndex` 错位 | **高** | **结论存疑，待实测** | 宿主实现 |
| ISS-42 |  `format_cells` 的 `merge` 失败被吞仍返回 success；`borders:'none'` 文档说可去边框但实现跳过 | 中 | **已修已验证**（merge 部分） | 宿主实现 + 说明 |
| ISS-43 | `manage_cell_comments` 的 `author` 参数不生效 | 低中 | **已修已验证** | 宿主实现 |
| ISS-44 | 条件格式 / 数据验证 / 冻结窗格**没有读回工具**，只能靠脚本探测 | 中 | **已修已验证** | 工具能力缺口 |
| ISS-45 | `clear_range` 属 `declaredNotCallable`，清空只能写 `null` 矩阵，极易残留 | 中 | **已修已验证** | 工具能力缺口 |
| ISS-46 | 回滚的"后续修改"判定基于**内容相等**，重建的同内容数据会被旧 `auditId` **清空**（真实数据破坏） | **高** | **已修已验证** | 回滚判定逻辑 |
| ISS-47 |  回滚文案未说明范围，被误读为"整表还原"（**原判"样式未恢复=缺陷"经复核不成立**，见下） | 中 | **已修已验证** | 回滚文案 |
| ISS-48 | 格式/条件格式/冻结/行列/图表类工具**不返回 auditId**，完全不可回滚且说明未写 | 中高 | **已修已验证** | 审计覆盖 |
| ISS-49 | 两个不同的 MCP 客户端在审计里都记成 `MCP Agent`，**无法区分归属** | 中高 | **已修待验**（桥接侧，静态验证） | 审计归属 |
| ISS-50 | 后台未运行时错误不可操作（`fetch failed` / 原始 urlopen 错误，不给服务名/端口/恢复方式） | 中 | **已修待验**（桥接侧，静态验证） | 错误文案 |
| ISS-51 | HTTP `/mcp` 会话硬上限 64 且**无空闲过期**，累积后 429，需手动 DELETE 释放 | 中 | **已修已验证** | 会话生命周期 |
| ISS-52 | `SKILL.md` 称脚本里 `doc`"已自动绑定"，实测 Excel 场景 `doc` 为 `null`，须用 `wb` | 中 | **已修已验证**（文档类） | skill 说明 |
| ISS-53 | `generate_deck` 标题页背景**超框 33%**（1280×720 / 960×540），根因待干净实验 | 中高 | **现象确认，根因待定** | 宿主实现 |
| ISS-54 |  `generate_deck` 的 `content` 布局缺 `bulletPoints` 时**静默只出标题**，无警告 | 中高 | **已修已验证** | 宿主实现 |
| ISS-55 | `generate_deck` 坐标基准与说明不符，调用方被迫自己再换算一次 | 中 | **已修已验证**（文档类） | 宿主实现 + 说明 |
| ISS-56 | `wps_execute_script` 返回值**嵌套超过两层就丢属性**（变 undefined 且不报错） | 中 | **已修已验证** | 脚本返回值序列化 |
| ISS-57 | `wps_word_page_layout_and_watermark` 不接 `pageNumberFormat`（422），页码只能退回脚本插域 | 中高 | **已修已验证** | 宿主实现 + 说明 |
| ISS-58 | 专用水印只落在**正文层第 1 页**，跨页水印需自行改页眉层 | 中 | **不修**（WPS 宿主不支持跨页页眉水印，已如实告警） | 宿主实现 + 说明 |
| ISS-59 | **磁盘上是新构建、WPS 进程里跑的是旧构建**：部署后不重载，且没有任何指纹能判断"运行中的是哪一版" | **高（方法学）** | **已修已验证** | 部署流程 + 版本可观测性 |
| ISS-60 | 行隐藏 `hide` 返回 success 但没生效（`Range("5:6").Hidden=true` 静默 no-op） | **高** | **已修已验证** | 宿主实现 |
| ISS-61 | 批注 `author` 无效：宿主 `Comment.Author` 恒为 jolin，返回体里的 author 是**入参回显冒充读回** | 中 | **已修已验证** | 宿主实现 |
| ISS-62 | 公式查找替换静默不支持（`totalFound:0` + `success`） | 中 | **已修已验证** | 宿主实现 + 说明 |
| ISS-63 | 透视表：建表瞬间是空骨架须手动 `Refresh`；`destSheetName` 必须已存在；`RecordCount` 恒为 1 | 中 | **已修已验证** | 宿主实现 + 说明 |
| ISS-64 | 条件格式无图标集/公式规则参数入口，也无读取/清除工具 | 中 | **部分完成**（读回已补；图标集/公式规则**入口**未加，属能力补强 CAP-06） | 工具能力缺口 |
| ISS-65 | `find_and_replace` 的 `results[].row/col` 是**区域相对偏移**，说明未写 | 低 | 台账已补记录；**说明文案待工具说明批次** | 工具说明 |
| ISS-66 | 批注成功消息拼接 bug（`function Address() { [native code] }`） | 低 | **已修已验证** | 宿主实现 |
| ISS-67 | ~~`wps_word_write_content` **吞掉所有小写字母 `a`**~~ → **【撤销·非缺陷】原结论错误**：现象是**验证探针自己造的**——探针写 `.replace(/[\r\a]/g, "")`，而 **JS 正则里 `\a` 是恒等转义 = 字母 `a`**（BEL 必须写 `\x07`，那是 C/PCRE/.NET 的语义），于是探针**先把读回文本里每个小写 `a` 删掉再比对**，必然"少 a"。宿主写入一直是正确的：原始输出里 `strippedMatch:[148]`（共 149 段）证明内容**正确落在文末** | **撤销**（原判"高"是误判） | **已澄清·非缺陷**（原始会话逐字取证；全仓库 1161 条正则穷举 **0 条**能删 `a`；无裸 `\a` 转义）。兜底仍保留并升级为**逐字比对**（等长改写也拦，超出原"比长度"） | **验证方法出错**（既非宿主实现，也非本仓库代码） |
| ISS-68 | `find_and_replace` 只查找时 `matchCount` **恒为 0**，文案还谎称"已应用格式化" | 中高 | **已修已验证** | 宿主实现 |
| ISS-69 | 导出 PDF 返回 `success` + `savedPath` 但**磁盘无文件** | 中高 | **已修已验证** | 宿主实现 |
| ISS-70 | `write_content` 的 `location:"bookmark"` 没插到书签处，还污染了另一个书签范围 | 中 | **已修已验证** | 宿主实现 |
| ISS-71 | `wps_word_capture_preview` **已实现但未注册**（调用报"未知工具"），Word 视觉验收有缺口 | 中 | **已修已验证** | 死分支（DP3 实例） |
| ISS-72 | `page_layout_and_watermark` 的 header/footer/watermark **只作用于第 1 节**且不提示 | 中 | **已修已验证** | 宿主实现 + 说明 |
| ISS-73 | Word 侧无样式/书签/内容控件/交叉引用/分节/行列尺寸/文档属性工具，全需脚本 | 中 | **部分完成**（读回已加，工具化待做） | 工具能力缺口 |
| ISS-74 | `native-scripting.md` 只有 Excel/PPT 示例，Word 表格/分节/样式零示例 | 中 | **已修已验证**（文档类） | skill 说明 |
| ISS-75 | `generate_deck` **原生图表数据写不进去**：`ser.Values=` 不生效也不抛错，图表显示宿主默认值，却返回 success | **高** | **已修已验证** | 宿主实现 |
| ISS-76 | `wps_ppt_capture_slide_preview` 恒 422"PPT 未生成预览"，但同一 API 用脚本可成功 → **真实错误被兜底文案替换** | 中高 | **已修已验证** | 宿主实现 |
| ISS-77 | **目标锁语义不一致**：宿主 `lockedTargets` 是加载项**进程级全局**，网关 `TargetLockStore` 是 `sessionId:host` 级 → 不传目标时行为不可预测 | **高** | **已修已验证** | 架构不一致（解释 ISS-02） |
| ISS-78 | PPT 无「新建/保存」工具；追加型工具**不幂等**且说明未写 | 中 | **已修已验证** | 工具能力缺口 + 说明 |
| ISS-79 | 越界页码报宿主内部 JS 错误（`Cannot read properties of null (reading 'Delete')`），缺中文上下文 | 低 | **已修已验证** | 错误文案 |
| ISS-80 | `insert_native_chart` **100% 失败**且文案误导：`AddChart/AddChart2/AddOLEObject` 都是 function 但返回 `null`、不建形状，报的却是"数据配置未完成" | **高** | **已修已验证** | 宿主实现 |
| ISS-81 | `generate_deck` 的 chart 布局同样失败，且**已插入的页不回滚**，留半成品 | 中高 | **已修已验证** | 宿主实现 |
| ISS-82 | `layoutIndex` 实为 **ppLayout 枚举**而非版式序号，越界（12）不报错 | 中 | **已修已验证**（契约快照已复核） | 工具说明 |
| ISS-83 | `manage_table` 的 enum 里有 `set_table_data`，**宿主没有该操作** → 报"未知的表格操作" | 中 | **已修已验证**（契约快照已复核） | 契约漂移 |
| ISS-84 | `set_table_data` 的 `data` 含数字就报 `arguments.data[1][2]: 类型不正确` | 中 | **已修已验证**（契约快照已复核） | 工具契约 |
| ISS-85 | **PPT 无保存工具**（`wps_ppt_save_presentation` → 未知工具） | 中高 | **已修已验证** | 工具能力缺口 |
| ISS-86 | `capture_slide_preview` 吞掉 addon 原始返回、且不允许传 `outputPath` | 中 | **已修已验证** | 宿主实现 |
| ISS-87 | **`set_background` 会串改全部页**（受控复现：设第 1 页后第 2 页也变红） | **高** | **已修已验证** | 宿主实现 |
| ISS-89 | 深度属性反射会让 **WPS 进程崩溃**（3 次崩溃报告，2 次栈指向 jsetapi→etcore） | **高** | **已修已验证**（真机确认：原崩溃点 `Cells` 现返回 223 个成员名不崩；`evaluate:true, maxMembers:30` 精确求值 30、跳过 110 并逐条给原因） | 宿主 API 安全性 + 工具护栏 |
| ISS-90 | 崩溃后宿主组件掉线，桥接无疑似崩溃信号与恢复指引 | 中高 | **已修已验证** | 可观测性 |
| ISS-91 | 读操作被别名到写函数（`list_conditional_formats` → 新增条件格式）；**但该 RPC 无工具暴露，不可达，属潜在风险** | **中（潜在）** | **已修·静态验证** | Office.js 路由 |
| ISS-92 | 批注 `action: "clear_all"` 静默成功（宿主不支持该 action）；`list_comments` 别名到默认 add | **高** | **已修已验证**（MS 实机） | Office.js 路由 + 处理函数 |
| ISS-93 | `normalizer.ts` 只适配 14 个方法，**12 个 `excel_*` 在 host=microsoft 参数错位**；21 个 Word/PPT 工具在该宿主是死路 | **高** | **已修已验证**（MS 实机 10/10） | 跨宿主适配 |
| ISS-94 | **C 类·只写不读共 7 项**：条件格式、冻结窗格、数据有效性、筛选状态、工作表保护/标签色、Word 页眉页脚/水印、透视表 | 中高 | **已修已验证** | 工具能力缺口 |
| ISS-95 | **B 类·宿主能做但没暴露**：清空区域 / 读原表设计语言 / Word 页面预览（**3 条零成本死分支**）＋超链接、命名区域、文档属性、区域复制、图片形状、结构化表格 | 中高 | **已修已验证** | 工具能力缺口 |
| ISS-96 | Office.js 的 `capture_sheet_preview` 是**伪渲染**：无图表时用 Canvas 按 `cellW=110/rowH=28` 硬编码合成假图 → **AI 视觉自检会得出与真实文件不符的结论** | **高** | **已修已验证**（MS 实机） | 证据可信度 |
| ISS-97 | adapter 把整张 `EXCEL_METHODS`(30) 当 COM 回退白名单，而 COM 实际只覆盖 28/30 → **白名单过度声明** | 中高 | **已修待验**（桥接侧，静态验证） | 回退策略 |
| ISS-98 | `excel_find_and_replace`（host=microsoft）字段名错位 → `text=""` → 每个非空单元格都命中并 `replaceAll("")`，**可能破坏内容** | **高** | **已修·静态验证**（MS 通道未连，无法运行时验证） | 跨宿主契约 |
| ISS-99 | **响应转换丢字段**：`office/normalizer.ts` 的白名单把加载项新返回的 `warnings` / `dataRangesApplied` / `seriesCount` / `yAxisApplied`（capture 的 `kind`/`renderedBy`）丢掉 → "哪些参数没生效"的如实提示到不了调用方 | 中高 | **已修已验证** | 响应转换 |
| ISS-100 | MS 侧行列插入"**执行成功却返回失败**"（裸读未 load 的 `sheet.name`） | 中 | 已修（MS 侧） | 宿主实现 |
| ISS-101 | MS 侧 ISS-96 的报错被"属性 name 不可用"盖掉（`load` 排在 `sync` 之后） | 中 | 已修（MS 侧） | 宿主实现 |
| ISS-102 | MS 侧 `dataRanges` **多段实际只生效 1 段**（`setData` 是替换不是追加） | 中 | 已修（改为只取首段 + 如实 warnings） | 宿主实现 |
| ISS-103 | `Range.Format` 在本机 WPS **不存在**（undefined）→ `write_content` 的 alignment / 行距 / 缩进**一直在静默失败** | **高** | 已修（改用 `Range.ParagraphFormat`） | 宿主实现 |
| ISS-104 | WPS 的 `Headers.Item(1).Shapes.AddTextEffect` **静默把形状加到正文层**（header.Shapes.Count 恒 0）→ 跨页页眉水印**在 WPS 上无法实现**（宿主限制，已如实告警） | 中 | 不修（宿主不支持，已记录） | 宿主限制 |
| ISS-105 | 网关 `ppt.ts:154` 用 `res?.error \|\| "PPT 未生成预览"` **覆盖掉宿主真实错误** | 中高 | **已修已验证** | 响应转换 |
| ISS-106 | 网关 `ppt.ts` 的 manageSlides **未转发 `filePath`/`format`**，schema 里的字段到不了宿主 | 中 | **已修已验证** | 参数转发 |
| ISS-107 | `wps_reload_addon` **不能加载新构建**：宿主实现是 `window.location.reload()`，只重跑已缓存的 JS，不重新读盘（实测部署后调用，`ADDON_BUILD_FINGERPRINT` 仍 `undefined`） | 中高 | **已修**（说明改为如实告知：必须彻底退出 WPS 再打开） | 宿主实现 + 说明 |
| ISS-108 | **本轮修复引入的回归**：`format_cells` 的 `rowHeight` 分支误写 `targetRange`（该函数里叫 `range`）→ 抛 `ReferenceError`。**`npm test` 未覆盖该路径，只有真机调用才暴露** | **高** | **已修已验证** | 宿主实现（真机自查） |
| ISS-109 | `patch_cells` 的 `formulas` 矩阵 schema 只允许 `string`，而宿主把 `null` 也当"空项跳过" → 说明承诺的空项写法有一半走不通（`null` 被 validateArgs 拦下） | 中 | **已修已验证** | 工具定义 |
| ISS-110 | **构建指纹心跳丢失**：`register` 有两处（`connection.js` 首发 + `bootstrap.js` 每 5 秒心跳），只改了前者 → 桥接收到的始终是心跳报文，`buildFingerprint` 恒为 null | 中高 | **已修已验证** | 宿主实现 |
| ISS-111 | **Word 页面预览恒不落盘**：宿主返回 `hasPdf:true` + `pdfPath`，但磁盘上没有文件。根因＝桥接把预览输出到 `~/.wps-bridge/previews/`，而 **WPS 是沙箱应用写不进去**（实测：该目录/`os.tmpdir()`/`~/Downloads` 子目录 均 ❌，只有 WPS 容器 tmp ✅） | **高** | **已修已验证** | 桥接预览路径 + 落盘校验 |
| ISS-112 | **宿主方法被"摘出来"调用导致 this 丢失**：`const f = cell.Address; f()` 抛 "Address called on null or undefined"，被 catch 吞掉后表现为**地址为空**（成功消息变成"已在单元格  添加批注"）。这是 ISS-66 的真正根因，前一版修复（把取地址提前）方向对但没解决 this 问题 | 中 | **已修已验证** | 宿主实现（真机定位） |
| ISS-113 | **MS 侧矢量形状：分组已修好；直线是宿主限制（结论已定）**。**分组**：根因是 `loadShapeDetail` 一阶段读不到 `type`（`type` 也要 load 才有值）→ `isLine`/`isGroup` 恒为 false → 对 Group load `fill`/`textFrame` → sync 抛错。改**两阶段加载**后真机通过：`组 ✔`、**解组释放 `["m1","m2"]`**（证明成员真的在组里）。顺带修 `memberCount` 恒为 0（`group.group.shapes.load()` 不带上下文 → 返回值不可读；改为先取代理再 load，读不到时**如实告警**而不是给个 0 让人以为组是空的）。**直线**：`addLine` 真机报「当前对象不允许此操作」，**判定为宿主限制**。**关键澄清**：ms-tester 曾据任务窗格自检称「addLine 真的建成了」——核实后**不成立**：自检里的 "line" 用的是 `addGeometricShape` + 设置 `lineFormat`，**从未调用 `addLine`**，所以自检通过不能证明 `addLine` 可用。替代方案已验证：用**细矩形**当线（`height≈4`）可正常建成。 | 中 | **分组已修已验证；直线确认为宿主限制** | Office.js 宿主 |
| ISS-117 | ~~筛选写入静默不生效~~ **误报，已撤销**。复核真因：验证脚本用错了 patch_cells 参数（传 `cells:[{address,value}]`，正确是 `address` + 二维 `values`），数据从未写进表；空区域上 `Range.AutoFilter()` 自然 no-op。**根本教训是脚本没检查写入结果**——已加「造数据后立即读回核对」 | — | **撤销**（非产品缺陷） | 验证方法 |
| ISS-118 | **透视表创建报 `CreatePivotTable` 为 null**。复核确认**双因**：① 源区域为空（同 ISS-117 的脚本错误）；② 非活动表上 `PivotCaches().Create(1, srcRange)` 返回 null。补防御性 `Activate()` 后真机通过（`refreshedBy: RefreshTable`、`tableRange: $D$1:$E$4`） | 中 | **已修已验证** | WPS 宿主实现 |
| ISS-119 | **结构化表格上设表级筛选会假成功**：目标区域若已是 ListObject（结构化表格），它自带筛选器，此时 `Range.AutoFilter()` 不生效、`sheet.AutoFilterMode` 仍为 false，但工具此前返回 `success:true`，调用方会以为筛选已开。**由最终验收测试发现**（CAP-33 在 CAP-19 建表之后失败）。已修：写入后读回核对，不生效即写入 warnings，并指明是哪个表格导致的 | 中 | **已修已验证** | WPS 宿主实现 |
| ISS-120 | ~~图表数据源不支持跨工作表~~ **已修**。真相：宿主**支持**跨表——`chart.SetSourceData(跨表 Range 对象)` 真机实测可拿到 2 个系列，只是**不接受带表名的字符串**（交 `AddChart2` 会报 `Parameter type error source`）。而桥接侧 `addChart` 里**根本没有 `wb` 变量**（只有 `sheet`），先前误用 `wb.Worksheets` 抛 ReferenceError 后被 catch 吞掉、**静默回退到当前表**，拿一张空表的区域去 SetSourceData → 报成「宿主不支持」，**看起来像宿主限制，其实是变量名写错**。已改用 `sheet.Parent.Worksheets` 解析表名，并**去掉静默回退**（找不到表就如实报错，不能悄悄把数据源换掉）。真机验证：在空表上建图、数据源指向另一张表 → 系列数 2 ✔ | 中高 | **已修已验证** | WPS 宿主 + 我们的实现 |
| ISS-122 | **图表数据源必须是交叉表，否则每行都变成一个系列**：把平铺明细（月份/区域/营收 18 行）直接当 `dataRange` → 宿主生成 **18 个系列**，图例挤成一团、颜色随机、完全不可读。用户一眼看出"图表不对"。正确做法：先做**月份 × 区域 的交叉表**再画图 → 3 个系列。**这条不是产品缺陷，是使用方法**——已写进能力验证脚本与 skill 提示 | 中 | **已修正用法** | 使用方式 |
| ISS-123 | **图表数据标签关不掉却报成功**：`update_chart` 传 `showDataLabels=false` 返回 success 但图表上标签仍在——宿主 `chart.ApplyDataLabels(0)` 不报错也关不掉。**由开关测试发现**。已修：改为**逐系列设 `HasDataLabels`** 再读回每个系列的真实状态，并返回 `dataLabelsMatchRequest`（命名直说含义：实际是否与请求一致），不一致时写 warnings | 中 | **已修已验证** | WPS 宿主实现 |
| ISS-124 | **`wps_update_chart` schema 与宿主实现不符**：宿主支持 `chartIndex`/`chartType`/`hasLegend`/`showDataLabels`，schema 里没有，反而有个宿主不读的 `legendPosition` → 数据标签等能力**调不到**。已对齐并补网关转发。**暴露 `check:params` 的盲区**：它只查 schema→处理器（声明了没读），查不出反方向的"处理器实现了、schema 没暴露" | 中 | **已修** | 契约一致性 |
| ISS-125 | 🔴 **`wps_word_page_layout_and_watermark` 写水印导致 WPS 主进程崩溃（SIGSEGV）**。真机：调用 `action:apply, watermarkText:内部机密严禁外传, watermarkColor:#C0C0C0` 返回 success 并给出「水印落在正文层」告警，**~1 秒后 WPS 整体退出**，Word/PPT/Excel 三组件同时掉线（code 1001）。崩溃报告 `~/Library/Logs/DiagnosticReports/wpsoffice-2026-09-22-204903.ips` 已由 Lead 独立核实：`EXC_BAD_ACCESS/SIGSEGV`，故障线程栈 `wpsapi +3272704` **连续重复 6 帧**（无限递归），调用链 `ksojscore → jswpsapi → wpsapi`，即我们的 JS API 调用触发；当天其它 5 次 wpsoffice 崩溃签名均不同。**影响面**：宿主崩溃会带走用户未保存的数据。**处置：立即禁止在该工具上传 watermarkText/watermarkColor；禁止受控复现**（代价太高）。另有一个独立问题：实现把水印写到了**正文层**而非页眉层，本身就不对 | 最高 | **已修（禁用）**：watermarkText 在任何宿主写操作前硬拒绝、删除 AddTextEffect 实现、read 通路保留；真机验证「被拒且不崩」 |
| ISS-126 | **残留目标锁：已修并在三个组件上生效**。真机复核（锁指向已关闭的 `agent-word.docx`、当前无 Word 文档）：**修复前**报「未在 WPS 中找到目标 Word 文档 [agent-word.docx]。当前已打开: 无」——既没有自动清理、也没有可操作提示；**修复后**报「**当前未打开任何 Word 文档，请先在 WPS 文字中打开目标文档**」。**关键更正**：此前（bridge-fixer 首轮）的修复只判**桥接侧**会话锁，而真机查证 `wps_get_locked_status` 返回 `gatewayLocks: {}` —— **桥接侧是空的**，那个死名字实际来自**加载项侧的 `lockedTargets`**（两套独立的锁）。本轮把自愈做在**加载项侧**（锁真正所在的地方），Word/PPT/Excel 三个解析器全部覆盖。**设计判据（两条同时成立才自愈）**：① 目标在宿主里确实不存在；② 该名字是**锁隐式生效**而非调用方显式传入 —— **显式传错名字时绝不改锁**，否则会悄悄换掉调用方要操作的文档（真机验证：显式传错名字仍如实报错且锁不变）。真机 2/2 通过 | 中高 | **已修已验证** | 加载项 + 桥接 |
| ISS-127 | **`wps_word_read_document` 的 `scope="paragraphs"` 与工具描述不符**：描述说返回"仅连续段落文本数组"，实测返回的是 `paragraphDetails` + `previewText`，**没有 `paragraphs` 数组**。调用方按描述取 `paragraphs` 会拿到 undefined。待复核 `maxParagraphs` 是否一并失效。**由 word-ppt-tester 发现** | 中 | **已修** | 补 `paragraphs` 数组（待真机复核 Word 侧） | agent 测试 |
| ISS-128 | **越界错误的呈现不统一**：`wps_word_format_document` 传 `paragraphIndex=999`（越界）报内部错误 `Cannot read properties of null (reading 'Range')`；而同类越界在 `write_content` 里是**中文可读报错**。错误信息不统一会让使用者分不清"我传错了"还是"工具有 bug" | 低 | **已修** | 段落越界改中文可读报错（含有效范围提示） | agent 测试 |
| ISS-129 | **`check:params` 的 `KNOWN_EXCEPTIONS` 声明了却从未被使用**——登记例外也不生效，等于这个逃生口是假的。已修复：problems 循环里真正过滤。**由本轮修复 readOnly 例外时发现** | 中 | **已修** | 检查脚本自身 |
| ISS-130 | **`manage_named_range` 的两条宿主限制已证实并写进工具描述**：① `comment` **宿主不存储**（写后读回恒为空）——现在如实给 warnings 而不是静默忽略；② **名字不能看起来像单元格地址**（如 `fz1` 会被宿主拒绝）。 | 低 | **已修（如实告知）** | WPS 宿主 |
| ISS-131 | **`create_workbook` 失败：宿主 `Workbooks.Add()` 返回 null（环境/宿主状态问题，非代码回归）**。真机现象：连续两次调用均返回 `success:false`「已新建工作簿 [null]（? 张表，首表 null）」，**宿主侧工作簿数不变**（只有 AgentExcelProbe.xlsx）→ 实际没建出来。直接探测证实：`app.Workbooks.Add()` **返回 null**，Add 前后 `Workbooks.Count` 都是 1，`DisplayAlerts=true`（不是对话框阻塞）。**判定**：这是**宿主状态**问题——同一段代码今日早前建簿成功过多次（H-4 验证、重点验收等），说明不是本次改动引入的回归。**待确认**：重启 WPS 后是否恢复；若恢复则属"长会话累积状态"，需在 `add_workbook` 路径加**返回值校验与可操作提示**（当前报错文案「已新建工作簿 [null]」容易被误读为代码 bug） | 中高 | **待修（待重启确认根因）** | WPS 宿主 |
| ISS-121 | **能力验证脚本的断言太弱**：只断言"调用成功"，没断言"结果对"。图表建出来是空的（跨表数据源无效）却判定为通过，是靠截图肉眼才发现。已在脚本里补"系列数 > 0"这类结果断言 | 中 | **已修**（补结果断言） | 验证方法 |
| ISS-114 | MS 侧形状自检**存在级联依赖**：后续检查引用前面创建的 `cap08_line`/`cap08_svg`，一旦前面失败就报「请求的资源不存在」，把「分组/删除不支持」误报成结论（修正为各项独立后，真实结果是 10/13 而非 5/12） | 中高 | **已修**（每项自查前置形状） | 测试设计 |
| ISS-115 | Office.js 的 `getTargetSheet` 从不 `load("name")`，而返回体普遍写 `sheet.name` → 读未 load 的属性直接抛「属性 name 不可用」，**形状工具全套因此全挂** | 中高 | **已修**（共享助手统一 load） | Office.js 宿主实现 |
| ISS-116 | MS 侧热重载失效：`window.location.reload(true)` 的 forceGet 参数在 Excel for Mac 的 WKWebView 中不生效（信号发出、返回成功、代码不变） | 中高 | **已修已验证**（改为 `location.replace` + 时间戳查询串强制不命中缓存） | 桥接 + 宿主 |
| ISS-88 | `swap_shapes` 只换 Top；`align_shapes` 实际是"对齐到首个形状"，且 `shapeIds` 先按 Id 再按索引 | 中 | **已修已验证**（契约快照已复核） | 工具说明 |

> 本表随盘点和修复推进持续追加。下面每节写清证据与修法。

---

## ISS-01 `excel_patch_cells` 混用 `values`+`formulas` 会静默清空数据

**严重度**：高（调用方看到"成功"，实际是目标区域被清空）

**现象**：同时提供 `values` 与 `formulas` 两个矩阵，`formulas` 里用 `''` 表示"该单元格只写值"时，第二步把第一步刚写的值覆盖成空。

**复现**（2026-09-22 实测，真实 WPS 宿主）：

| 步骤 | 操作 | 观察 |
|---|---|---|
| 1 | `excel_patch_cells` 写入 `A40:D40 = ["原始A","原始B",100,"原始D"]` | 读回正确 |
| 2 | 对同一区域 patch：`values=[["新A","新B",200,"新D"]]` + `formulas=[["","","",""]]` | 返回 `success:true`、**"已成功修改 4 个单元格"** |
| 3 | 读回 `A40:D40` | **`[null,null,null,null]`**（新旧值都没了） |
| 4 | `wps_rollback`（第 2 步返回的 `auditId`） | ✔ 恢复为原值 |

**根因**：`wps-addon/src/excel.js:360-361`

```js
if (values !== undefined && values !== null) range.Value2 = values;
if (formulas !== undefined && formulas !== null) range.Formula = formulas;  // '' 即清空
```

**影响面**：任何"部分单元格写值、部分单元格写公式"的调用都会丢数据；`diff` 里能看到 `newValue: null`，但**面向调用方的文案是"已成功修改 N 个单元格"**。

**规避（当前版本）**：公式直接放进 `values`（实测可行：`values=[["=1+1","=2*3"]]` → 读回公式 `=1+1`、值 `2`），不要同时传 `formulas`。

**建议修法（= DP7，待决策）**：A（推荐）`formulas` 中空项视为"不动该单元格"；B 混用且含空项时直接报错；C 只改文档。落到 `patchCells` + `tests/addon.test.ts` 回归 + 工具说明。

**验证方式**：补一条测试——写值后对同区域用混合矩阵 patch，断言原值仍在；并在真实宿主复跑本文件 §复现表。

**证据**：[p5.4-macos-wps-business-matrix.md](p5/p5.4-macos-wps-business-matrix.md) §3

---

### ✅ ISS-01 修复记录（2026-09-22，方案 A）

**改动**：`wps-addon/src/excel.js` 的 `patchCells` —— `formulas` 矩阵中的空项（`''`/`null`/`undefined`）改为**跳过**，不再参与赋值；只有整张矩阵都没有空项时才走原来的整批 `range.Formula = formulas`（保住批量性能）。空项路径改为逐单元格 `range.Cells.Item(r, c).Formula = ...`。

**回归测试**：`tests/addon.test.ts` 新增 3 项，宿主模拟**复刻了 WPS"给 Formula 赋空字符串即清空单元格"的真实行为**：
- `formulas` 全空 → 同批写入的值必须保留；
- 混合矩阵（`=1+1` + 空项）→ 公式落位、空项那格的值保留；
- 无空项 → 值写入行为不变（防止修复误伤正常路径）。

**修复前复现（证明测试有效）**：临时用 `git show HEAD:wps-addon/src/excel.js` 回退实现并重建 `addon-core.js`，重跑测试 → **正好这 2 项失败**；恢复修复版后 5/5 通过。全量回归 **76/76**（原 73 + 3）。

**待验**：实现只在仓库里，**运行中的 WPS 加载项仍是旧版**——需重新部署加载项后，按本文件 §复现表在真实宿主复跑一遍才算验证通过。

**剩余（归入批次 A）**：`excel_patch_cells` 的工具说明尚未写明"空项不改公式、清空请用 `values` 的 `null`"，待工具说明批次一并改。

---

## ISS-02 `hasOpenWorkbook` 与真实打开状态不符

**严重度**：高（会让 AI 直接放弃任务）

**现象**：工作簿明明在 WPS 里开着，`excel_get_workspace_summary` / `/api/v1/status` 返回 `hasOpenWorkbook: false`，并附一句"未在 WPS 中找到目标工作簿 [X]。当前已打开: Y"。

**根因**：该字段反映的是**目标锁指向的文档**是否存在，不是"当前实际打开了什么"。锁来自更早的会话，指向一个已关闭的文件。

**影响面**：AI 只看这个字段会判断"用户没打开表格"，于是报错收场；实际必须**显式传 `workbookName`** 才能读到真实工作区。

**建议修法**：把"已打开的工作簿列表"与"当前锁定目标是否存在"拆成两个字段（如 `openWorkbooks` + `lockTarget`），并修正文案；工具说明里写明二者的区别。

**验证方式**：在锁目标未打开、但有其他工作簿打开的情况下，断言 `hasOpenWorkbook` 为真且 `lockTarget.missing` 为真。

**证据**：[p5.4-macos-wps-business-matrix.md](p5/p5.4-macos-wps-business-matrix.md)（另：2026-09-22 盘点期间实测）

---

## ISS-03 参数约定不一致

**严重度**：中（每次调用都要试探，易错）

**现象**：`excel_*` 系列大多接受 `host` + `workbookName` + `sheetName`，但：
- `wps_rollback` 只接受 `auditId`（传 `host`/`workbookName` 报"未知参数"）；
- `excel_save_workbook` 接受 `host`/`workbookName` 但**不接受** `sheetName`（同一批调用统一加 `sheetName` 时会 422）。

**影响面**：调用方按同一套约定批量调用会踩坑；本次盘点期间连踩两次，都是靠读错误信息才发现。

**建议修法**：要么统一"所有 excel/wps 工具都接受 host/workbookName/sheetName（不适用则忽略）"，要么在工具说明里显式写明各自接受的参数。推荐前者（`additionalProperties: false` 目前会直接拒绝，说明它是有意收紧的——那更要在说明里写清楚）。

**验证方式**：契约测试断言"同一域内工具的参数集合一致，或有显式豁免清单"。

---

## ISS-04 已交付包内仍显示"42 项能力已就绪"

**严重度**：中（用户可见的错误声明）

**现状**：源码 `office-addon/public/taskpane.html:38` 已改为 `Microsoft Excel · 27 项能力已就绪`；但**已交付的三个包内仍是旧文案**。

**影响面**：已经过使用者构建物验收的三个包，内容与当前源码不一致。

**建议修法**：二选一——A 重新打包三份包（SHA 全变，构建物验收需重做）；B 不重打，登记为已知差异，并入后续 P5.4/P5.5 完成后的一次统一打包。

**验证方式**：解包 `app.asar.unpacked/office-addon/public/taskpane.html` 搜索 `42` 应为 0 命中。

**证据**：[final-review.md](final-review.md) §10.1、§11

---

## ISS-05 能力文案检查漏掉加载项页面

**严重度**：中（防御性缺口，已导致 ISS-04 长期未被发现）

**根因**：`check:claims` 的 `TARGET_DIRS` 只有 `skills/` 与 `website/src/`；`CLAIM_PATTERNS` 也不含"`N 项…能力`"。而 `taskpane.html` 是**手写文件**（`build:office-addon` 只生成 `taskpane.js`），两重漏网。

**现状**：**已修已验证**——扫描范围加入 `office-addon/public`、`wps-addon`；新增模式 `(\d+)\s*项[^，。；、\n]{0,16}?能力`；宿主必须写明（写 `Microsoft Excel` 而非裸 `Excel`）。反向验证：把 27 改回 42 → `exit 1`。

**验证方式**：`npm run check:claims`（27 个文件）exit 0；注入 42 后 exit 1。

---

## ISS-06 加载项等价性证据无留档

**严重度**：中（P3 的核心保证无法复核）

**现象**：P3.2/P3.3 声称的 "acorn 顶层语句多重集 120/121 逐字节相同"、"diff 证明 0 行删除/修改、仅新增 62 行"、"桩测试 20 条 RPC 序列快照逐行一致"，**只在台账正文里，验收目录无对应文件**。

**建议修法**：重跑这三项比对并把输出留档到 `docs/acceptance/2.1.0-p0p1/evidence/`；Office.js 侧 20 条序列快照优先（其余 57 个分支目前只有"原样搬迁 + 语法检查"级保证）。

**验证方式**：文件存在且结论与台账一致；台账引用改为指向文件。

---

### ✅ ISS-06 修复记录（2026-09-22，三项比对已重跑并留档）

**结论**：三项比对**全部重跑，数字与台账声称逐项吻合**（详见下表）。

| # | 台账声称 | 重跑实测 | 证据文件 |
|---|---|---|---|
| A | acorn 顶层语句多重集 **120/121** 逐字节相同 | 改造前 **121** 条，其中 **120** 条可在改造后逐字节找到 → **120/121**；改造前独有恰好 1 条 = `fitGeneratedPptShapes`；改造后独有 5 条 = 该函数 + **4 个新抽出的纯函数**（`pptScaleForPage`/`pptCountTextUnits`/`pptEstimateLines`/`pptFitFontSize`） | [evidence/iss-06-acorn-multiset.txt](evidence/iss-06-acorn-multiset.txt) |
| B | diff 证明 **0 行删除/修改、仅新增 62 行** | 删除 **0** / 修改 **0**；新增 **62** = 注释 60 + 空行 2，**非注释非空的代码行新增 0 条**。附带声明同步复核一致：`dispatchExcelTool` 77 case、58 个函数定义、`context.sync()` 61、`.load(` 40、`Excel.run(` 46，前后完全相同 | [evidence/iss-06-diff-stat.txt](evidence/iss-06-diff-stat.txt) |
| C | 桩测试 **20 条 RPC 序列快照逐行一致** | **20/20 条**序列逐行一致、响应签名逐条一致、不一致 **0** 条。**额外加固**：全部 **77** 个 `dispatchExcelTool` 分支同样 **77/77** 一致（脚本 `--all`） | [evidence/iss-06-rpc-sequence.txt](evidence/iss-06-rpc-sequence.txt) |

**比对输入（两个文件都是仓库内既有文件，未做 git 操作、未构建、未改源码）**：

| 角色 | 文件 | 行数 | sha256 | 为何是它 |
|---|---|---|---|---|
| A/B/C 改造前 | `.scratch/evidence/prior/wps-addon/addon-core.js`（由 `release/_legacy-build/…/wps-addon/addon-core.js` + `baseline/baseline-prior.patch` 的该文件补丁段用系统 `patch` 应用得到） | 5170 | `6450c171…` | ★ 与 `baseline/MANIFEST.sha256` 的该条目**逐字符相同** —— 被证明就是 P0 基线（P3.2 改造前）的字节内容 |
| A 改造后 | `release/2.1.0/win/win-unpacked/resources/app.asar.unpacked/wps-addon/addon-core.js` | 5247 | `c30e548a…` | ★ 与 `p5/p5.1-candidate-and-prebuilt-state.md` 的"构建前指纹"一致；5247 行与 P3.2 声称的一致 |
| B/C 改造前 | `.scratch/evidence/prior/office-addon/public/taskpane.js`（取自 `release/_legacy-build/…`） | 1809 | `ca86ef85…` | 1809 行与 P3.3 声称的一致 |
| B/C 改造后 | `release/2.1.0/win/win-unpacked/resources/app.asar.unpacked/office-addon/public/taskpane.js` | 1871 | `5ef08e87…` | ★ 与 `p5.1` 的"构建前指纹"一致；1871 行与 P3.3 声称的一致 |

> **为什么不用仓库当前的两个入口**：它们已被 P5 各项修复推到 5631 / 1899 行（`28cbb199…` / `99d87f8d…`）。ISS-06 要复核的是 **P3.2/P3.3 改造当时**的等价性，因此取当时的那份生成物；`release/2.1.0` 是它在本机唯一被记录过指纹的落点。

**重跑方式**（脚本随证据一起留档在 `evidence/`，可原样复跑；按约定不放仓库根目录）：

```bash
node docs/acceptance/2.1.0-p0p1/evidence/iss-06-acorn-multiset.mjs \
  .scratch/evidence/prior/wps-addon/addon-core.js \
  release/2.1.0/win/win-unpacked/resources/app.asar.unpacked/wps-addon/addon-core.js
diff -u .scratch/evidence/prior/office-addon/public/taskpane.js \
  release/2.1.0/win/win-unpacked/resources/app.asar.unpacked/office-addon/public/taskpane.js
node docs/acceptance/2.1.0-p0p1/evidence/iss-06-rpc-sequence.mjs \
  .scratch/evidence/prior/office-addon/public/taskpane.js \
  release/2.1.0/win/win-unpacked/resources/app.asar.unpacked/office-addon/public/taskpane.js [--all]
```

**如实登记的两点口径**：
1. P3.3 当时所用的那份"20 条 RPC 名单"**没有随证据留档**，本次是按脚本内固定名单**重建**的 20 条（覆盖读/写/脚本/图表/表格/批注/预览）。名单本身不影响结论——结论是"同一名单、同一参数下改造前后逐行一致"；为此另跑了全部 **77** 个分支，同样 77/77 一致。
2. 序列快照与 diff 结论**互为印证而非重复**：diff 已证明可执行代码逐字节未变，序列快照是行为层面的独立复核。

**仍未覆盖（保留，不在本条修复范围）**：`issues.md` ISS-06 原文提到的"其余 57 个分支目前只有原样搬迁 + 语法检查级保证"——本次 `--all` 已把桩测试扩到 **77/77**，该限制**已被解除**；但**真实宿主上的等价性仍未验证**（两个入口都没有在 P3 改造后做 WPS/Excel 实机加载对比，见 P3.8"未做真实宿主回归"）。

**证据**：[evidence/iss-06-acorn-multiset.txt](evidence/iss-06-acorn-multiset.txt)、[evidence/iss-06-diff-stat.txt](evidence/iss-06-diff-stat.txt)、[evidence/iss-06-rpc-sequence.txt](evidence/iss-06-rpc-sequence.txt)；脚本 [evidence/iss-06-acorn-multiset.mjs](evidence/iss-06-acorn-multiset.mjs)、[evidence/iss-06-rpc-sequence.mjs](evidence/iss-06-rpc-sequence.mjs)；执行记录 [p5/mcp-sweep/14-evidence-fixes.md](p5/mcp-sweep/14-evidence-fixes.md)

---

## ISS-07 台账与证据文件数字漂移

**严重度**：低（结论未变，但引用会误导）

**现状**：**已修已验证**。修正项：P1.4 故障注入 "10 项 / pass 8 / fail 2" → `tests 11 / pass 9 / fail 2`（共 6 处文件）；证据文件 SHA 陈旧；`gateway.ts` 285 → 280；`any` 口径标注为未复现；阶段台账 `test 57/57` 无出处 → 改为"test 全部通过"；`baseline.md` 33 vs 35 统一；"从 HEAD 恢复" → "从锚点提交恢复"；Windows 环境表述澄清；交接备忘与 P2.3 状态矛盾统一。

**验证方式**：全仓 grep 旧数字无残留；文档相对链接全部可达；`check:agents` exit 0。

**证据**：[final-review.md](final-review.md) §10

---

## ISS-09 D8/D9 能力声明残留

**严重度**：中

- **D8**：`wps.ppt` 描述"有限支持：基础页面、文字、形状、表格；复杂图表与母版需实机验证"，但 WPS 侧实际暴露 `wps_ppt_insert_native_chart` 且有原生图表实现——**声明偏窄，漏报已存在的能力**。
- **D9**：`capabilities().hosts` 只给 excel 的 `implemented` 数组，Word/PPT 仅一句描述，因此 `bridge_get_capabilities` **不能作为"支持哪些工具"的完整来源**。

**建议修法**：D8 按实际能力面重写描述；D9 让 capabilities 覆盖 Word/PPT（或明确标注"仅 Excel 完整"并在说明里指向 `bridge_get_tools`）。

**验证方式**：`check:claims` + 人工核对；D9 需断言 `bridge_get_capabilities` 的覆盖范围与 `tools/list` 一致。

---

## ISS-10 9 条未确认项

**严重度**：中（不能静态推断，需实机）

见 `tool-inventory.md` 第 180 行起：实机可用性未确认、部署版本与源码是否一致未确认、"42 项"来源未确认、各工具宿主超时阈值未确认、无读回工具的写操作失败判定未确认等。

**建议修法**：并入 P5.4 业务矩阵逐项实测。

---

### ✅ ISS-10 处理记录（2026-09-22，9 条逐条定状态）

**产出**：**[p5/open-questions.md](p5/open-questions.md)** —— 每条给出「当前状态 / 结论 / 证据指向」，仍未确认的写清**下一步怎么确认**（具体到"打开什么、调哪个工具、看哪个字段"）。

**状态分布：9 条中 6 条已澄清、3 条仍未确认。**

| # | 未确认项 | 状态 | 结论一句话 |
|---|---|---|---|
| 1 | 实机可用性未确认 | ✅ 已澄清 | 第一轮的"纯静态"局限已被真实宿主盘点解除：3 个真实成品、e2e 跨组件一致性 18/18；结论是**能跑通成品，但大量工具"返回 success 却没生效"**（ISS-38/39/17/29/11/42） |
| 2 | 部署版本与源码是否一致未确认 | ✅ **已澄清（结论：不一致）** | 实测查明"磁盘上是新构建、进程里跑的是旧构建"，升级为 **ISS-59（高）**；版本号相同、构建不同，现有版本号比对拦不住 |
| 3 | "42 项"的来源未确认 | ✅ 已澄清 | 来源是两处手写文案（`catalog.ts` 的 validation 字段 + `taskpane.html:38`），与 `EXCEL_METHODS`=30 / `excel_*`=27 都对不上；两处均已改写。残留：已交付三包内仍是 42（ISS-04，决策方案 B 不重打） |
| 4 | 各工具的宿主超时阈值未确认 | ❌ **仍未确认** | 只有 1 个实测锚点（`Presentations.Add()+SaveAs` 21 s 超时，且文案正确标"结果未知、不要重放"）；逐工具阈值表仍未建立 |
| 5 | 无专用读回工具的写操作失败判定未确认 | ✅ 已澄清 | 缺口成立并升级为 ISS-44/ISS-64/ISS-94（C 类只写不读 7 项）；修法已决策（补读回工具）；判定路径已用脚本读回实测走通（ISS-41） |
| 6 | host=microsoft 下审计快照缺失的实际后果未确认 | ❌ **仍未确认** | Microsoft 通道自始未连接，属外部前置条件；静态结论已知（无快照 → 回滚被拒），但落库形态与文案未验 |
| 7 | Microsoft 侧 Word/PPT 能力边界未确认 | ✅ 已澄清 | `mcp-sweep/08-ms-vs-wps.md`（552 行双通道对照）+ ISS-93（21 个 Word/PPT 工具在该宿主是死路）；"Office.js 原生通道"说法已判不实并改写 |
| 8 | 8 行的"部分成功语义"未逐函数核对 | ✅ **已澄清 8/8** | 8 个函数全部拿到实测或代码复核结论：`freeze_panes` 两条判"不修"（ISS-41）；`set_data_validation` 两条实测"清除原校验却返回 success"；`wps_word_write_content` → ISS-67（**该条已复核撤销·非缺陷**，见文末"撤稿记录"）；`page_layout_and_watermark` → ISS-57/58/72；`ppt_add_business_cards` **可用**（超栏有明确报错）；`ppt_insert_native_chart` → ISS-80/81 |
| 9 | 审计记录条数上限（500）的副作用未确认 | ❌ **仍未确认** | 代码事实已复核（`audit-store.ts:47-48` 静默淘汰），但从未有测试触及上限（P5 期间只 34 条留痕）；淘汰后的回滚文案、是否影响 ISS-46 判定，均未验 |

**说明**：标记"已澄清"**不等于"问题已解决"**——#1/#2/#5/#8 澄清出来的都是真实缺陷，已各自登记为独立 ISS 条目。"澄清"指的是"这条未知项不再是未知"。#4/#6/#9 的确认依赖外部前置条件（宿主会话 / Microsoft 通道），已写明可执行的下一步，**不当作已完成**。

**证据**：[p5/open-questions.md](p5/open-questions.md)（含逐条证据文件索引表）

---

## 第一轮全方位盘点发现（4 个并行子代理）

> 方式：只给**成品目标**，不给步骤；由子代理自己从工具清单与 skill 说明判断该用什么工具。目的就是反推"说明够不够用"。
> 产出目录：[mcp-sweep/](p5/mcp-sweep/)

### ISS-11 绘图对齐传字符串被静默吞掉

**严重度**：高（**静默错误**——不报错、结果不对、极难自查）

**现象**：`wps_execute_script` 里设 `HorizontalAlignment = 'center'`，赋值**不报错**，但实际被吞成左对齐（枚举值 `-4131`）；必须传整数 `-4108` 才是居中。子代理第一版所有居中文字全跑到左上角，**全量重绘一次**才修好。

**影响面**：任何用字符串枚举做格式化的脚本都会静默产出错误排版；调用方拿不到任何异常。

**建议修法**：在 `skills/office-agent-bridge/references/native-scripting.md` 明确"枚举必须用整数"，并附**对齐 / 形状类型 / 箭头样式**的取值表；更好的做法是在脚本执行层加一层校验，字符串枚举无法识别时**直接报错**而不是静默回退。

**验证方式**：脚本里传字符串枚举应抛出可读错误；文档枚举表与实测取值一致。

**来源**：`shapes` 子代理，2026-09-22

---

### ISS-12 WPS JS API 与 VBA 的差异没有说明

**严重度**：中高（会让 AI 直接中断）

**现象**：`TextFrame.TextRange` 在 WPS JS API 中是 `undefined`（VBA 里有），必须改用 `Characters()`；`ParagraphFormat` 也不存在（空成员表）。子代理首轮探测即抛 `Cannot read properties of undefined (reading 'Font')` 并整体中断。

**影响面**：AI 的默认先验是 VBA/ Office.js 文档，直接用会踩空；且 `native-scripting.md` 完全没提这些差异。

**建议修法**：在该文档里给一节 **"WPS JS API ≠ VBA" 差异清单**（可用的文本操作方式、不存在的成员、替代写法），并给出可运行的探测方法。

**验证方式**：按文档写的示例能在真实 WPS 上直接跑通。

**来源**：`shapes` 子代理

---

### ISS-13 `wps_inspect_api` 不返回枚举常量

**严重度**：中高（关键信息无处可查）

**现象**：该工具只能反射出**方法名**（如 `Shapes` 集合 34 个方法），但绘图与格式化**大量依赖枚举常量**（形状类型、对齐、箭头样式、线型），这些值既不在文档里、也问不到，只能靠试和查外部资料。

**影响面**：把"能不能用脚本画图"变成"能不能猜对枚举值"；与 ISS-11 叠加后，猜错还不报错。

**建议修法**：让 `wps_inspect_api` 支持列出常见枚举常量表；或在 skill 里附**枚举速查表**（至少覆盖形状类型、对齐、箭头、线型、颜色常量）。

**验证方式**：盘点的多元素构图任务，不需要外部资料即可完成。

**来源**：`shapes` 子代理

---

### ISS-14 `Shapes.Range([...]).Group()` 静默产出损坏的分组对象

**严重度**：中

**现象**：调用不报错，但组对象**名称错乱**、`GroupItems` 不可枚举——即产出了一个"看起来成功、实际不可用"的对象。

**建议修法**：文档明确标注该能力在 WPS JS API 下不可用；若宿主支持，最好在适配层检测并报错。

**验证方式**：文档有明确结论；或脚本调用后能检出不可用并报错。

**来源**：`shapes` 子代理

---

### ISS-15 `native-scripting.md` 缺多元素构图范例与画布换算

**严重度**：中（直接导致返工）

**现象**：文档只有"画一个矩形"的最小示例；**没有**多元素构图范例、没有画布与截图区的换算（实测默认列宽 48pt、行高 16pt → `A1:U42` = 1008×672pt）、没有布局坐标建议。

**影响面**：子代理走了 **6 轮探测 + 2 次全量重绘**才做出成品；这类任务本应一次成型。

**建议修法**：补一个完整的"信息图 / 流程图"范例（含分组框、箭头、KPI 条），并写清画布尺寸换算与常用布局参数。

**验证方式**：把该范例原样交给一个 AI，能在少量轮次内复现同类成品。

**来源**：`shapes` 子代理

---

### ISS-16 `save_workbook` 是工作簿级操作，并发时会落盘他人在途结果

**严重度**：中（多 Agent 场景下的真实风险）

**现象**：多个任务并行操作同一工作簿时，任一任务调用 `wps_save_workbook` 会把**整个工作簿当前内存状态**写盘，包含其他任务尚未完成/尚未打算保留的中间结果。

**影响面**：并行使用时，"我只保存我的部分"这个预期不成立；本次 4 个并行子代理场景下已实际发生（`shapes` 保存时把兄弟任务的第 9、10 张表一并落盘）。

**建议修法**：在工具说明里写明"保存是工作簿级的，会连同其他未保存改动一起落盘"；长期看可考虑提供"仅保存指定工作表"或显式的变更事务边界。

**验证方式**：说明文档有该警告；或有针对性的并发行为测试。

**来源**：`shapes` 子代理旁注

---

#### ISS-16 实测场景与规避建议（2026-09-22 补充，证据取自 `mcp-sweep/` 并行子代理记录）

**★ 说明文案本身写在哪**：工具描述（`wps_save_workbook` / `excel_save_workbook`）的告警文字**不在本条目范围**，由工具说明批次负责（见 [p5/mcp-sweep/09-tool-desc-fixes.md](p5/mcp-sweep/09-tool-desc-fixes.md)）。本小节只补**实测场景 + 规避建议**，供写说明者直接取用。

**实测场景 1（原始发生）**：第一轮盘点派了 **4 个并行子代理**（`inventory` / `tables` / `charts` / `shapes`），**同时操作同一个工作簿**，各自建自己的表（`sweep图形_*` / `sweep图表_*` / `sweep表格_*`）。
`shapes` 子代理收尾调 `wps_save_workbook` 时，**把它两个兄弟任务当时还只在内存里的第 9、10 张表一并写进了磁盘文件**——它自己只打算保存第 8 张表。
（场景与并行方式见 [p5/mcp-sweep/README.md](p5/mcp-sweep/README.md) §方法、§三个真实成品。）

**实测场景 2（独立复核，更能说明问题）**：稍后 `e2e` 子代理对同一个工作簿做端到端链路。它只新建了自己的 `e2e_钙钛矿数据集` 一张表，然后调 `wps_save_workbook`：

```
wps_save_workbook → 已成功保存到磁盘，磁盘 mtime 15:55
落盘后独立解 OOXML 包核对：xl/workbook.xml 表名清单共 18 张，
含 e2e_钙钛矿数据集，以及第一轮并行盘点留下的全部 sweep* / fault_* / __MCP验收测试__ 表
```

**关键点**：`e2e` 子代理在报告里只能写"**未改动**任何 `sweep*` 表"（它确实没改内容），但它**无法选择不写它们**——一次 `Save()` 就把 17 张不属于它的表连同自己的表一起落盘。（见 [p5/mcp-sweep/05-e2e.md](p5/mcp-sweep/05-e2e.md) §产物表、§步骤 2。）

**根因（源码，一句话）**：保存走的是宿主 `Workbook.Save()`，粒度就是整个工作簿，实现里没有任何"只保存某张表"的分支：

```js
// wps-addon/src/excel.js:989-998  saveWorkbook()
const wb = getWorkbook(app, workbookName);
wb.Save();                       // ← 整个工作簿，包含所有未保存的内存改动
return { success: true, workbookName: wb.Name, fullName: wb.FullName,
         message: `工作簿 [${wb.Name}] 已成功保存到磁盘` };
```

同源记录：[tool-inventory.md](tool-inventory.md) 对 `excel_save_workbook` / `wps_save_workbook` 的说明均为"**整簿多表一次落盘**；仅调用 `wb.Save()` 后即返回成功，未校验磁盘结果"。

**为什么调用方一定会踩**：返回文案是 `工作簿 [X] 已成功保存到磁盘`，**不含"本次保存影响了哪些表"**。调用方读到的是"我的保存成功了"，完全看不出它刚刚替别人提交了未完成的改动。

**规避建议（按可行性排序，给多 Agent / 并行编排场景）**：

1. **一个工作簿只给一个写任务**（最省事、最可靠）。并行任务各用一个工作簿，或各用**独立文件**；本项目并行盘点的正确做法本该是"每个子代理一个 `.xlsx` 副本"，而不是共用一份。
2. **把"保存"收归协调方**。子代理**不要自己调 `save_workbook`**，只改内存并**报告自己改了哪些表**；由协调方在所有子任务结束后**统一保存一次**。这样"落盘"这个动作只发生在没有在途任务的时间点上。
3. **并行期间一律不保存**。若必须共用一份工作簿，把 `save_workbook` 视作**全局屏障操作**：先确认所有兄弟任务都已完成或已明确放弃，再保存。P5 的落盘核验（P5.6）就是这个模式。
4. **保存前先读回自己的表做完整性自查**，并核对**磁盘 mtime 与自己的调用时间对得上**——但要注意：这只证明"保存发生了"，**不能**排除"顺手写了别人的表"；能排除的是"保存根本没生效"（ISS-59 场景）。
5. **需要真正隔离时用脚本路径**：若仅需导出单表，改用 `wps_execute_script` 单独 `SaveAs` 到新文件，而不是对共享工作簿调 `Save()`。
6. **给编排指令加一条硬约束**：派并行子代理时明确写"**禁止调用 `wps_save_workbook`**"——本次 4 个子代理的指令里没有这一条，因此它是**主动踩出来的**。（对照：`README.md` §过程问题已记录"后续派子代理时应在指令里指定中间产物目录"，同一类指令缺陷。）

**长期修法（不在本次范围，供排期参考）**：提供"仅保存指定工作表/新建副本另存"的能力，或引入显式的变更事务边界（开始/提交/回滚），使"保存"不再是事实上的全局提交。

---

### ISS-17 `wps_add_chart` 的 `chartType` 静默降级

**严重度**：高（返回 success，结果不是要的类型，调用方无法察觉）

**现象**：`chartType` enum 列了 10 个值，**实际只有 6 个真实映射**。实测：
- `scatter` → 建出**簇状柱状图**（读回 `ChartType=51`，2 个系列）
- `area` → 同样降级成 51
- `column` → `column_clustered`、`bar` → `bar_clustered`

调用返回 `success`，`title`/`yAxis` 设置都生效，**没有任何降级提示**。真散点只能 `wps_execute_script` + `AddChart2(201, -4169, ...)`。

**建议修法**：要么补上真实映射，要么把 enum 收窄到实际支持的类型，要么在降级时**明确返回警告**（例如 `data.warnings: ['chartType scatter 降级为 column_clustered']`）。

**验证方式**：逐个 chartType 建图并读回真实 `ChartType`，断言与请求一致或有显式警告。

**来源**：`charts` 子代理

---

### ISS-18 `wps_add_chart` 定位参数被静默忽略，且说明严重夸大

**严重度**：高（说明与实现不符 + 静默叠图）

**现象**：`startCell` / `endCell` / `cellRange` / 顶层 `left`/`top`/`width`/`height` **全部被忽略**；实测 4 张图全部落在默认 `left=360, top=40`，互相压在一起。只有 `position.leftCell` + `position.width/height` 生效。

而工具说明把 `cellRange` 描述为"**刚性单元格吸附…100% 完美的行级对齐…杜绝像素漂移**"——这是**与实现完全不符的承诺**。

另外 `position` 对象的 `required:["leftCell"]` 有效，但 `position` 本身不在顶层 `required` 里，缺省时静默落到 360/40，多图场景**没有任何警告**。

**建议修法**：按宿主是否真的读取这些字段来重写说明；删掉"100% 完美/杜绝像素漂移"这类无依据表述；`position` 缺失时给默认值并返回警告。

**验证方式**：文档描述的定位参数逐个实测生效；多图不传 `position` 时能收到警告。

**来源**：`charts` 子代理

---

### ISS-19 `wps_delete_chart` 的 `leftCell` 是像素邻近批量匹配，会误删多张图

**严重度**：**高**（破坏性操作，参数名误导）

**现象**：传 `leftCell:"H2"` 一次**删掉了 14 张图**——因为它按像素邻近（<30px）匹配，叠加在一起的图全部命中。参数名听起来像"精确锚点"，实际是范围杀伤，且与同工具里 `chartName`/`shapeName` 的精确语义混在一起。

**建议修法**：`leftCell` 改为精确匹配或加 `matchMode` 参数并默认 exact；说明里写明匹配规则；删除前返回**将被删除的清单**供确认。

**验证方式**：构造多张邻近图表，断言 `leftCell` 只删目标那一张。

**来源**：`charts` 子代理

---

### ISS-20 `wps_update_chart` 在 WPS 不可用，工具说明未标注

**严重度**：中

**现象**：调用返回 HTTP 422 + `wps_update_chart 在 WPS 宿主上未实现（加载项缺少该 RPC 分支），已拒绝且未执行`。**拒绝得很干净**（未执行、无半成品），与 `contracts/host-methods.ts` 的 `HOST_IMPLEMENTATION_GAPS.wps=['update_chart']` 一致——**但工具 description 里没有一个字说它在 WPS 上不可用**，只在 `bridge_get_capabilities` 的 `unimplementedOnHost` 里能查到。

**影响面**：AI 会反复尝试修改图表，然后退回脚本。

**建议修法**：在工具定义层直接标注宿主可用性（或让描述由 `HOST_IMPLEMENTATION_GAPS` 自动生成）。

**验证方式**：`tools/list` 里 `wps_update_chart` 的描述包含宿主限制说明。

**来源**：`charts` 子代理

---

### ISS-21 `wps_patch_cells` 静默把字符串日期转成序列号

**严重度**：中

**现象**：写入字符串 `"2026-01"`，被静默存成日期序列号 `46023`；调用方要再设 `numberFormat` 才能显示回原样。说明里**没有任何提示**，也没有 `asText` 之类的开关。

**建议修法**：说明里写明该行为；或提供"按文本写入"的显式选项。

**验证方式**：写入 `"2026-01"` 后读回，断言行为与文档一致。

**来源**：`charts` 子代理

---

### ISS-22 `wps_format_cells` 的 `rowHeight` 不生效

**严重度**：低中

**现象**：传 `rowHeight` 无效果，须用脚本直接设 `Rows.Item(n).RowHeight`。

**建议修法**：修实现或在说明里标注"在 WPS 宿主上不生效，请用脚本"。

**来源**：`charts` 子代理

---

### ISS-23 `wps_get_charts` 的定位与详情不自洽

**严重度**：中

**现象**：`leftCell` 与实际 `left/top` 不自洽——多张叠在 360/40 的图**都报 `$D$2`**，作为"定位依据"会误导（子代理据此判断过位置，判断错了）。`detail:true` 又不返回**图表类型枚举名**（只给数字）和**系列颜色实际值**。

**建议修法**：让 `leftCell` 与实际像素位置一致；`detail:true` 补上类型枚举名与颜色。

**来源**：`charts` 子代理

---

### ISS-24 `seriesColors` 对单系列图表是逐点染色

**严重度**：中

**现象**：本应单色的柱图/条形图变成彩虹色。另：`Points()` 集合在 WPS 上计数不可靠（6 个点报 2），只能按系列级重刷。

**建议修法**：说明里区分"系列颜色"与"逐点颜色"；或对单系列图表忽略数组长度>1 的 `seriesColors` 并警告。

**来源**：`charts` 子代理

---

### ISS-25 `wps_capture_sheet_preview` 的参数优先级与返回字段未说明

**严重度**：低中

**现象**：`address`（区域）与 `chartName`（单图）在说明里并列可选，**没说同时传时谁优先**；返回体字段名（实际是 `imageBase64`/`imageMimeType`/`imageSizeBytes`）也没在 description 里写，调用方只能试。

**建议修法**：补参数优先级说明与返回结构。

**来源**：`charts` 子代理

---

### ISS-26 脚本 API 差异补充（`native-scripting` 缺失项）

**严重度**：中（与 ISS-12/15 同源，一并修）

- `ws.Cells(r,c)` **不存在**，报 `ws.Cells is not a function`，只能用 `ws.Range("A1")` 字符串地址；
- `AddChart2` + `SetSourceData` **默认按行取系列**（一行一个系列），必须显式 `PlotBy=2` 或逐个赋 `Series.Values/XValues` 才正确——文档没写。

**建议修法**：并入 ISS-12/ISS-15 的文档修订：补一节 WPS JS API 差异 + 常见陷阱。

**来源**：`charts` 子代理

---

### 过程问题：子代理在工作区根目录留下了中间产物

`charts` 子代理用 Python 驱动桥接，在仓库根目录创建了 `.sweep-chart-probe/`（**4.8 MB**：11 个 `.py` 脚本 + `__pycache__` + 9 张预览 PNG），不属于任何产物目录。

**处置**：保留结论性截图（`preview/dashboard_FINAL.png`）作为验收证据，其余中间脚本清理；后续派子代理时应在指令里**指定中间产物目录**（如 `.scratch/<代号>/`）。

**来源**：`charts` 子代理

---

## 第一轮盘点：能力全景与说明质量（`inventory` 子代理）

> 完整报告：**[mcp-sweep/01-inventory.md](p5/mcp-sweep/01-inventory.md)**（507 行，含逐域清单、原文证据、P-01～P-14 与复现方式）。下面只记可执行条目。

### ISS-27 缺少"我该用哪个"的选型指引

**严重度**：高

**现象**：**27 对 `excel_*` 与 `wps_*` 工具完全同构**，说明文字只差一句样板；选型规则只写在 MCP `instructions` 里，`tools/list` 的调用方看不到。另有一对真重复（`modify` / `manage_rows_and_columns`），说明未讲清区别。

**影响面**：AI 面对 91 个工具无法判断该选哪一个；选错可能落到没有宿主的通道上。

**建议修法**：把"选型规则"写进工具描述本身（或让两者之一在描述里明确指出何时用另一个）；真重复的工具合并或明确分工。

**验证方式**：把 `tools/list` 单独交给一个 AI，能选出正确工具。

---

### ISS-28 `wps_rollback` 的恢复范围被夸大

**严重度**：高

**说明原文（全文 23 字）**：`根据留痕记录 ID 一键撤销修改，原地恢复表格`

**实际范围**：只覆盖 `patch_cells` 写入的**值与公式**；样式、图表、结构变更、Word/PPT 操作、脚本执行**全部不可回滚**。

**影响面**：调用方以为"任何修改都能一键还原"，在真正出事时会发现救不回来。

**建议修法**：说明里写明"仅覆盖值/公式 patch"；更彻底的做法是给其他写操作也加快照。

**验证方式**：文档明确列出覆盖范围；对不支持的操作用例断言返回"不可回滚"而不是成功。

---

### ISS-29 `search_cells` 承诺了宿主没有的能力，且失败返回成功

**严重度**：高

**说明原文**：`在表格中快速搜索包含指定文本或**公式**的单元格坐标`

**实测**：只读 `Value2`，**按公式搜索返回 0 命中且 `success: true`**——比直接报错更危险（调用方会得出"确实不存在"的错误结论）。另外报错文案要求调用方"缩小检索范围"，但该工具**没有**这个参数。

**建议修法**：要么实现按公式检索（读 `Formula`），要么把说明改成"只搜值"；报错文案同步修正。

**验证方式**：写一个含公式的单元格，按其公式文本搜索应命中；或文档明确不支持并有对应测试。

---

### ISS-30 审计族工具说明近乎空白

**严重度**：高

**现象**：审计与回滚一组 4 个工具的说明只有 **15–17 字**，**9 个参数全部没有说明**。调用方不知道能查什么、按什么过滤、返回什么结构。

**建议修法**：补齐说明与参数描述（可参照说明写得好的工具样例，见报告第六节）。

**验证方式**：说明完整度检查（存在无描述参数即报错）。

---

### ISS-31 参数缺 JSON Schema 类型

**严重度**：中

**现象**：**11 个参数没有 `type`**，另有 1 个使用了非标准类型。缺类型会让调用方无法判断该传数组还是对象、数字还是字符串。

**建议修法**：补齐 `type`；加一条契约测试断言"所有参数都有合法 type"。

**验证方式**：契约测试通过。

---

### ISS-32 说明文本被装配层正则误伤

**严重度**：中

**现象**：交付给调用方的 `office_execute_script` 说明**结尾是逗号**、句子不完整——说明组装时的正则替换误伤了文本。

**建议修法**：修装配逻辑；加断言"说明不得以标点结尾且不含未替换占位符"。

**验证方式**：`tools/list` 里该字段是完整句子。

---

### ISS-33 "尚待实机验收"的措辞把范围写窄了

**严重度**：中

**现象**：27/27 个 `excel_*` 的说明都追加同一句 `Windows Microsoft Excel 尚待实机验收。` —— **只提 Windows**，读者会推断"macOS 已验证"。而实际上 Microsoft 通道在 macOS 上也**没有实机验收**（`msExcel` 未连接，候选版本自述未实机验收）。

**建议修法**：改为"Microsoft Office 通道尚未实机验收（macOS/Windows 均是）"。

**验证方式**：与 `compatibility-diff.md` §3.3 的口径一致。

---

### ISS-34 枚举成员与参数说明不齐

**严重度**：中

**现象**：例如 `wps_word_format_document` 的 `preset` 枚举含 `academic`，但参数说明只解释了 `gov_standard` / `business_modern` / `custom` 三项；另有同义别名的优先级未定义。

**建议修法**：枚举值与说明逐一对齐；加断言"枚举成员必须在说明中出现"。

---

### ISS-35 写工具零必填参数

**严重度**：中

**现象**：**17 个写工具没有任何必填参数**（连"目标"都不是必填）。最突出的是 `wps_add_chart` / `excel_add_chart`：说明承诺"创建与数据源动态绑定的原生矢量图表"，但 `dataRange`、`dataRanges`、`chartType` **一个都不是必填**——"与 dataRanges 选其一"只是正文提示，schema 层没有强制。`wps_delete_chart` 同样零必填却带 `clearAll` 开关（与 ISS-19 叠加后危险性更高）。

**建议修法**：把真正必需的参数标进 `required`；破坏性开关（如 `clearAll`）要求显式确认参数。

**验证方式**：契约测试断言关键写工具的必填集非空。

---

### ISS-36 说明营销化，挤占有效信息

**严重度**：低中

**现象**：全清单说明合计 **6549 字符（均值 71 字）**，其中每个 `excel_*` 固定追加 32 字样板（约 864 字，占 13%），另有促销式表述（如"彻底解决旧图表无法清除、留下空白残缺边框的痛点"）。

**影响面**：调用方读到的信息密度低，真正关键的约束（互斥、宿主差异、降级）反而缺失。

**建议修法**：删样板与营销句，把字数花在约束与示例上。

---

### ISS-37 错误提示不给允许值

**严重度**：低

**现象**：参数不合法时的报错只写"未知参数"，**不列出允许值**；参数集在工具间不一致时尤其费轮次（本次盘点实测踩过：`wps_rollback`、`excel_save_workbook`）。

**建议修法**：报错附上允许参数集（与 ISS-03 一并处理）。

---

### ISS-02 补充证据（来自盘点）

盘点独立复现并加固了 ISS-02，补充两点：
1. **自举死锁**：省略 `workbookName` 会回落到上一会话的陈旧锁目标并报"未找到[另一个文件]"；而**唯一能列出已打开文件的工具，恰好就是失败的那两个**（`excel_get_workspace_summary` / `wps_get_workspace_summary`）——调用方没有别的路可走。
2. `wps_get_locked_status` 承诺返回的文件列表，**实测为 `null`**。

### 盘点确认的能力现状

| 项 | 数字 |
|---|---|
| 对外工具 | **91**，分 9 域（诊断 2 / 目标锁 5 / `excel_*` 27 / `wps_*` 27 / Word 12 / PPT 9 / 脚本反射 2 / `office_*` 3 / 审计回滚 4） |
| 只读标注 / 写入 | 21 / 70 |
| 本次只读实测返回 `success` | 11 个 |
| 有真实运行留痕（写入侧） | 仅 `patch_cells`（审计 34 条，含 1 次真实回滚） |
| **未验证的写工具** | **70 个**——未验 ≠ 可用 |

---

## 第一轮盘点：复杂表格实测（`tables` 子代理）

**成品**：`sweep表格_钙钛矿企业跟踪看板`（`A1:N34`）——合并大标题 + 副标题、4 张 KPI 公式卡、**双行表头**（分组带 + 4 处纵向合并）、20 家企业数据、序号/评分/评级/相对均值四列动态公式、合计行、3 张结论卡、冻结前 8 行 + A/B 列、表头筛选、N 列下拉、4 条条件格式、4 条批注、斑马纹与边框。已由协调方拉渲染图核对，确认可用。

### ISS-38 `wps_set_filter_and_sort` 假成功（排序根本没生效）

**严重度**：高（返回 `success`，数据没变，调用方无法察觉）

**现象**：返回 `success` + `sortedRuleCount:1`，**数据完全没排序**。先 `Activate` 目标表重试仍无效。改用 `wps_execute_script` 走 `sh.Sort.SortFields.Add + SetRange + Apply()` 才排成降序。

**根因（待确认）**：适配层走旧式 `Range.Sort(...)`，本机 WPS 接收后**静默 no-op**。

**建议修法**：改用 WPS 实际生效的 `SortFields` 路径；**返回前读回验证排序是否真的生效**，未生效则报错而不是返回 success。

**验证方式**：排序后读回断言顺序变化；未生效必须报错。

---

### ISS-39 `wps_auto_fit_columns` 静默 no-op 且 `address` 参数被忽略

**严重度**：中高

**现象**：不传 `columnRules` 时返回 `success` + **`results: []`**，列宽纹丝不动（全为 8.5），与其说明"不传则自适应全表已用区域"**矛盾**；`address` 参数被实现**完全忽略**。

**建议修法**：要么实现默认自适应全表，要么把 `columnRules` 标为必填并修正说明。

**验证方式**：不传参数调用后读回列宽，断言有变化或明确报错。

---

### ISS-40 筛选范围不受 `range` 约束

**严重度**：中高

**现象**：`Range("A8:N28").AutoFilter()` 被**自动扩展成 `$A$8:$N$30`**，把合计行卷进筛选区。隔空行、显式 `AutoFilter(1)` 均无效——始终扩到最后一个已用行，且**没有参数能约束**。

**影响面**：用户一筛选，合计行就会跟着被过滤掉（或混进结果），是表格交付质量问题。

**建议修法**：让 `range` 真正限定筛选范围；或提供显式的筛选区域参数。

---

### ISS-41 `freeze_panes` 的列索引错位

**严重度**：高（参数与结果差 1，无报错）

**现象**：`freezeColumnIndex` 实测 `set2 → 读回1`、`set3 → 读回2`——**要冻结 A、B 两列必须传 4**。`SplitRow`（行）没有这个问题，只有列错位。

**建议修法**：修正列索引换算；回执文案补上列的信息（当前只提行不提列）。

**验证方式**：传 N 后读回断言恰好冻结 N 列（含边界 0/1/负数）。

---

### ISS-42 `format_cells` 的失败被吞 + `borders` 语义不符

**严重度**：中

- `merge` 失败被 `try/catch` 吞掉，**仍返回 success**；只能靠 `get_range_styles` 的 `merged` 字段回验。
- 说明称 `borders` 传 `'none'` / `false` 可去边框，**实现里 falsy 直接跳过**——永远去不掉边框。

**建议修法**：失败不得吞；边框语义按文档实现或改文档。

---

### ISS-43 `manage_cell_comments` 的 `author` 参数不生效

**严重度**：低中

**现象**：传了 `author`，读回作者仍是系统用户（`jolin`），**传的值被拼进正文**。

**建议修法**：实现真正写入作者，或从参数里移除并在说明中写明。

---

### ISS-44 格式化类操作缺读回工具

**严重度**：中（可验证性缺口）

**现象**：**条件格式、数据验证、冻结窗格都没有对应的读回工具**，只能用 `wps_execute_script` 探测。这与 P5.4 要求的"保留结构化读回"直接冲突——写入类工具没有对应的读取能力，验收只能靠脚本反射或截图。

**建议修法**：补 `get_conditional_formatting` / `get_data_validation` / `get_freeze_panes` 等读回工具（也与 ISS-23 的 `get_charts` 详情缺失同类）。

---

### ISS-45 `clear_range` 不可调用，清空只能写 `null` 矩阵

**严重度**：中

**现象**：`clear_range` 属 `declaredNotCallable`（无对外工具），清空区域只能把 `null` 写进 `values` 矩阵。本次实测因**只清了 A 列而残留过一整行旧合计**，靠渲染截图才发现。

**影响面**：清空操作容易漏列/漏行且无提示，属于静默残留。

**建议修法**：按 DP3 重新评估是否为 `clear_range` 补对外工具；或至少在 `patch_cells` 说明里写明"清空请对整块区域传 null 矩阵"。

---

## 第一轮盘点总结

| 维度 | 结果 |
|---|---|
| 子代理 | 4 个（`inventory` / `tables` / `charts` / `shapes`），全部完成 |
| 成品 | 3 个真实工作簿成果（信息图 134 元素 / 5 图表看板 / 20 行跟踪看板），**均经协调方拉图核对** |
| 新增问题 | **ISS-11 ～ ISS-45**（含 ISS-02 补充），合计 35 条新增条目，其中 **高严重度 12 条** |
| 最典型模式 | **"返回 success 但什么也没发生"**：排序、列宽自适应、图表类型降级、公式搜索、对齐枚举、合并失败 —— 这一类占了高严重度的多数 |
| 说明质量主线 | 说明描述的定位/范围参数被实现忽略（`cellRange`/`startCell`/`address`/`range`），且**写得越绝对越危险**（"100% 完美行级锁定""彻底解决…痛点"） |

> 未验证的部分照实说明：`inventory` 只读实测 11 个工具，**70 个写工具的真实行为仍未验证**；3 个成品都**没做关闭重开后的持久化复验**（P5.6 只对测试表做过一次）。

## 第二轮：故障与并发用例（`fault` 子代理）

原始证据：`.scratch/fault/REPORT.md`（含每条用例的完整返回原文）。目标工作簿新建了 6 张 `fault_*` 测试表，未改他人表。

### ✅ 通过项（这些是站得住的）

| 用例 | 结果 |
|---|---|
| 回滚负向：事后改动同区域再回滚 | **正确拒绝** —— `目标区域已有后续修改或快照缺失，已拒绝覆盖。请先核对当前内容。` |
| 回滚负向：行结构变更后回滚 | 正确拒绝 |
| 重复回滚 / 不存在的 id / 缺参 / 类型错 | 均有明确且不同的提示 |
| 乱序撤销：先撤销新记录、再撤销旧记录 | **逆序撤销正确** |
| 区域外改动不影响回滚 | 正确（只回滚目标区域） |
| 双客户端并发 8 次写入 | 8/8 成功，宿主串行 7.93s，**无串写** |
| 双客户端锁隔离 | 锁按会话隔离，A 锁 A 文档、B 锁 B 文档互不影响 |
| 同区域并发写（两会话各 5 次） | 10/10 成功，最后写入者获胜，审计 10 条按序完整，**无丢写** |
| 审计按 sheetName 过滤 | 只返回该表记录，无串扰 |
| 后台重启 | 加载项**自动重连**，工作簿 16 张表全在，审计历史 68 条保留 |

### ISS-46 回滚把"别人重建的同内容数据"当成自己的修改，直接清空

**严重度**：**高（真实数据破坏）**

**复现**：
1. `patch_cells` 写入 `fault_case2_tmp!A1:B2`，拿到 `auditId`；
2. 删除该工作表；
3. 重建同名空表并写回**完全相同的内容**（用脚本）；
4. 用**旧 `auditId`** 回滚 → 返回 `{"success":true,"message":"成功恢复区域 $A$1:$B$2 的数据"}`；
5. 读回 → `[[null,null],[null,null]]` —— **重建的数据被清空**。

**根因**：回滚的"是否有后续修改"判定用的是**内容相等**，而不是"是否同一次修改"。内容碰巧相同就被放行。

**建议修法**：判定加上工作表身份（如创建时间/内部 ID）或版本戳，而不是只比内容；或要求回滚时显式确认。

**验证方式**：本用例应被**拒绝**；补一条回归测试。

---

### ISS-47 回滚执行后样式"没有回来"，但文案让人以为整表还原了

> ⚠️ **本节结论已在下方"复核与修复记录"中修正**：现象成立，但"部分成功被报成完全成功"的因果判断**不成立**——样式从来不在该记录的范围内。降级为中，按文案清晰度修。

**严重度**：**高（部分成功被报成完全成功）**

**复现**：patch 得 `auditId` → 同区域**加粗 + 红底** → 用该 `auditId` 回滚 →
`{"success":true,"message":"成功恢复区域 $A$1:$B$3 的数据","restoredRows":3,"restoredCols":2}`
→ 读回：**值被清空，但 6 格仍是 `bold=true,fontColor=#FFFFFF,backgroundColor=#FF0000`**。

**影响面**：调用方以为回到原状，实际样式仍是改后的状态，且文案没有任何提示。

**建议修法**：要么把样式纳入快照/恢复，要么明确返回 `restored: {values: true, styles: false}` 并改文案。

---

### ISS-48 格式化类操作完全不进审计，也无法回滚

**严重度**：中高

**现象**：`format_cells` / `add_conditional_formatting` / `freeze_panes` / `modify_rows_columns` / `add_chart` **都不返回 `auditId`**；`wps_get_audit_history` 里只有 `actionType=update_values`（即 `patch_cells`）。图表即使建了，回滚数据记录也**不会删除图表**。

**影响面**：与 ISS-28（`wps_rollback` 说明夸大范围）互为因果——说明没写，实现也没覆盖。

**建议修法**：给这些写操作补审计记录与快照，或统一在说明里划清"可回滚 / 不可回滚"清单。

---

### ISS-49 两个不同的 MCP 客户端在审计里无法区分

**严重度**：中高（审计归属失去意义）

**现象**：两个独立 MCP 会话（sessionId `2752ce20…` / `dee788a5…`）写入后，审计里的 `clientName` **恒为 `MCP Agent`**（`src/bridge/mcp-server.ts:29`）。只有走 HTTP `/api/v1/tool/call` 才认 `clientName`（实测记为 `fault-HTTP-Client`）。

**影响面**：多客户端场景下，审计记录无法回答"这条是谁改的"——而审计归属正是 P2.2 专门修过的问题。

**建议修法**：MCP 会话应从客户端 `initialize` 的 `clientInfo` 取名称，而不是写死。

**验证方式**：两个不同 MCP 客户端写入后，审计记录里的 `clientName` 应可区分。

---

### ISS-50 后台未运行时的错误不可操作

**严重度**：中

| 调用方 | 返回 | 问题 |
|---|---|---|
| 项目 skill 的 `bridge_client.py` | `<urlopen error [Errno 61] Connection refused>` | 不点明服务、端口、恢复方式 |
| skill 的 `wps_client.py` | `无法连接到 WPS Bridge 服务 (http://127.0.0.1:19890)。请确认本地守护进程正在运行。` | ✅ 这条写得好，可作模板 |
| `--doctor` | `{"installationRecord":true,"credentialsPresent":true,"error":"<urlopen error ...>"}` | 正常字段与原始错误混在一起 |
| `--status` | `{"ready":false,"occupied":false,"message":"fetch failed"}` | **`fetch failed` 完全不可操作** |

**建议修法**：统一按 `wps_client.py` 的口径（服务名 + 地址 + 恢复动作）。

---

### ISS-51 HTTP `/mcp` 会话无空闲过期，累积到 64 直接 429

**严重度**：中（可用性）

**现象**：`/mcp` 会话硬上限 **64**（`src/bridge/ws-server.ts:188`），**没有空闲过期**。一次性会话累积后 `initialize` 返回 `HTTP 429 {"error":"会话数量达到上限"}`（实测两次）；`DELETE /mcp` 可手动释放，恢复后新会话正常。

**建议修法**：加空闲超时回收；429 文案里给出释放方式。

---

### ISS-52 `SKILL.md` 关于脚本里 `doc` 的说明与实测不符

**严重度**：中

**现象**：`SKILL.md` 第 11 行称脚本里 `doc`"已自动绑定"；实测 Excel 场景下 `doc` 为 `null`，脚本里访问 `doc.Worksheets` 报
`原生脚本执行异常: Cannot read properties of null (reading 'Worksheets')`，必须改用 `wb`。

**建议修法**：按组件说明可用变量（Excel → `wb`，Word/PPT 另计），并入 ISS-12 的"WPS JS API 差异"一节。

---

## 复核与修复记录（2026-09-22，第二批）

> 复核原则：**动手前先读实现**。下面前两条是**修正我自己之前的判断**——子代理报的现象成立，但因果推断不成立。

### ⚠️ ISS-47 重新定性：原判"部分成功被报成完全成功"**不成立**

**子代理报的现象**（成立）：`patch_cells` 拿到 `auditId` → 同区域加粗+红底 → 用该 `auditId` 回滚 → 读回「值已恢复、样式仍是红底加粗」。

**原判**（我写的）：回滚报成功但样式没恢复 = 部分成功被报成完全成功。**这条是错的。**

**复核依据**：读完 `rollbackCells`（`wps-addon/src/excel.js`）与 `src/bridge/gateway/audit.ts` 后确认——**该审计记录从一开始就只快照了值与公式**，样式从来不在它的范围里。回滚确实**完整地做完了它承诺的事**（值+公式恢复），原文案"成功恢复区域的数据"里"数据"指的就是值/公式，**是准确的**。

**真正的问题在三处，已分别归属**：
- 工具**说明**写"原地恢复表格"，读起来像整表还原 → 归 **ISS-28**；
- `format_cells` 等格式类写操作**根本不进审计**，所以样式改动无从回滚 → 归 **ISS-48**；
- 返回**文案与结构**没能让调用方（连测试者都被误导）看清覆盖边界 → **本次修复**。

**因此 ISS-47 降级为中，并按"文案清晰度"修**，不再按"数据缺陷"修。

### ⚠️ ISS-41 结论存疑：子代理的推断与代码不符

**子代理结论**：「要冻结 A、B 两列必须传 `freezeColumnIndex = 4`」。

**代码事实**（`wps-addon/src/excel.js` `freezePanes`）：
```js
if (freezeRowIndex && freezeRowIndex > 1) win.SplitRow = freezeRowIndex - 1;
if (freezeColumnIndex && freezeColumnIndex > 1) win.SplitColumn = freezeColumnIndex - 1;
```
行列**用的是同一套 `-1` 约定**，完全对称。要冻结 A、B 两列应传 `freezeColumnIndex = 3`（→ `SplitColumn = 2`），不是 4。

**子代理的观察"set3→读回2"恰好就是这个正确行为**，它把"写入值比参数小 1"误当成了错位。

**判定**：**不修**。若按子代理结论去掉 `-1`，反而会把正确的行为改坏。已标记为「结论存疑，待实测」——需要一个能读出"实际冻结了几列"的独立验证（`SplitColumn=2` 意味着冻结前两列，这是 Excel/WPS 的既有语义）。**在拿到实测证据前不动代码。**

### ✅ ISS-42 修复（merge 失败不再被吞）

**改动**：`formatCells` 的合并/取消合并从 `try { range.Merge() } catch { log(...) }` 改为**执行后读回校验**——捕获异常后抛出；并用 `range.MergeCells` 读回实际状态，与请求不符时抛出带读回值的错误。符合"写工具未生效必须报错"的决策。

未覆盖部分：`borders: 'none'` 文档说可去边框而实现 falsy 跳过（同条目的后半），留待工具说明/边框批次。

### ✅ ISS-39 修复（`auto_fit_columns` 不再静默 no-op）

**改动**：不传或传空 `columnRules` 时，按 `sheet.UsedRange` 的列范围**逐列自适应**并返回实际结果（含 `mode: "usedRange"` 与条数说明）；已用区域为空时明确回"未调整任何列"。与说明承诺的"不传则自适应全表已用区域"对齐。

### ✅ ISS-47/28（文案层）修复

`src/bridge/gateway/audit.ts` 的返回改为：
- `message`: `成功恢复区域 ${address} 的值与公式`
- 新增 `restoredScope: { values: true, formulas: true, styles: false, charts: false, structure: false }`

让调用方一眼看清覆盖边界，不再靠猜。

**验证**：`npm run build:wps-addon` + `node --check` 通过；`typecheck 0`；全量测试 **76/76**；生成物与源码一致。三条均为**已修待验**——需重新部署加载项后在真实宿主复跑。

---

## 第二轮：PPT / Word 子代理（被 DSH 重启中断，由协调方接手复核）

> **过程说明**：DSH 于 15:39 重启以开通屏幕查看权限，`ppt` 与 `word` 两个子代理进程被销毁（`send_message` 冷启动返回 "active teammate not found"），**无法续跑**。协调方接手：从它们留下的 `.scratch/ppt/`、`.scratch/word/`（含 244 行 `run.log`）与落盘成品中复核出下述结论。
> **成品本身已在磁盘上**（见文末"成品完成度"），缺的是它们最后的汇报。

### ISS-53 `generate_deck` 生成的标题页背景**超出页面 33%**

**严重度**：中高（几何缺陷，实测可复现）

**实测**（`测试演示文稿.pptx`，页面 `960×540` pt）：

| 页 | 形状数 | 最右/最下 |
|---|---|---|
| 1（标题页） | 3 | **1280 × 720** ← 超框 |
| 2/3/4/9（卡片页） | 22/17/22/17 | 893 × 480 ✅ |
| 5/8 | 2 | 893 × 100 ⚠️ |
| 6/7（图表页） | 3 | 893 × 500 ✅ |
| 10（结尾页） | 2 | 960 × 540 ✅ |

第 1 页的背景矩形是 `1280×720`，而页面只有 `960×540`。逐形状读数：
```
1|1|0,0|1280x720|          ← 背景，超框
1|2|80,187|800x68|Office Agent Bridge
1|3|80,307|800x39|让 AI 直接操作正在运行的办公文档｜2025 Q3 产品与运营复盘
```

**候选根因（假说，尚未证实）**：`wps-addon/src/ppt.js` 里 `generate_deck` **混用了两套坐标系**——标题页背景用**真实页面尺寸**创建（`AddShape(1, 0, 0, pageWidth, pageHeight)`），而多数元素用 **720×405 设计基准**的常量；随后第 357 行 `fitGeneratedPptShapes(slide, 1, page, warnings)` 又把**从索引 1 开始的全部形状**按 `pageWidth/720 = 1.333` 统一放大 → 按真实尺寸创建的元素被**二次放大**。

**未证实的原因**：标题文本框读数（`80,187` / `800` 宽）只被放大过一次，与"背景被放大两次"不一致，说明还有别的因素（子代理的换算脚本也动过这份文件）。**需要一次干净实验**：新建空演示文稿 → 只调 `generate_deck` → 逐形状测量，才能定死根因。**在拿到该实验前不修代码。**

### ISS-54 `content` 布局在 spec 缺 `bulletPoints` 时**静默只出标题**

**严重度**：中高

**实测**：第 5 页（"调用链路：一次请求如何落到文档"）与第 8 页（"关键指标明细"）只有两个形状：
```
5|1|67,40|827x48|调用链路：一次请求如何落到文档
5|2|67,100|93x0|            ← 高度 0、无文字
```
正文完全缺失，且**没有任何警告**。

**成因**：`generate_deck` 的 `content` 分支只在 `spec.bulletPoints` 非空时才建正文文本框；`deck-outline.json` 里第 5、8 页只有 `layout` 与 `title`，没有 `bulletPoints` → 什么都不生成。

**建议**：`content` 布局在缺少正文时**至少要给警告**（工具已经会返回 `layoutWarnings`），或让 `content` 直接拒绝空正文。

### ISS-55 `generate_deck` 的坐标基准与说明不符

**严重度**：中

**现象**：PPT 子代理留下一份 `js/01-normalize.js`，注释写着：

> 背景：generate_deck 声称按真实页面尺寸生成，实测所有形状仍是 720x405 设计坐标，只占 960x540 页面的左上 75%。这里等价重做 addon 内部 fitGeneratedPptShapes 的换算。

即**调用方被迫自己补一次换算**才能得到正确版面。

**建议**：要么让说明写明"坐标为 720×405 设计基准"，要么让工具自己保证输出即正确版面（与 ISS-53 一并处理）。

### ISS-56 `wps_execute_script` 返回值**嵌套超过两层就丢属性**

**严重度**：中（脚本能力的关键可用性问题）

**实测**（本次复核时亲历）：
- 返回 `[{ page:{w,h}, slides:[{s,shapes,maxRight,maxBottom}] }]`（两层）→ **正常**
- 返回 `[ { slide, shapes:[ { i, type, L, T, W, H, txt } ] } ]`（三层）→ **内层对象的全部属性变成 `undefined`**
- 改成返回**扁平字符串数组** → 正常

**影响面**：脚本里做结构化探查（最自然的写法）会拿到一堆 `undefined`，且不报错——又是一个"静默失真"。

### ISS-57 页码格式无工具支持

**严重度**：中高（Word 交付物必需能力缺口）

**实测**：`wps_word_page_layout_and_watermark` 传 `pageNumberFormat: "第 X 页 / 共 Y 页"` → **HTTP 422 `arguments.pageNumberFormat: 未知参数`**（子代理试了两次都这样）。只能退回 `wps_execute_script`，用 `Find` 定位占位符 + `doc.Fields.Add(..., "PAGE"/"NUMPAGES")` 手工插域。

**建议**：补页码格式参数，或在说明里写明"页码需自行用脚本插入域"。

### ISS-58 专用水印只落在正文层第 1 页

**严重度**：中

**实测**：`wps_word_page_layout_and_watermark` 的 `watermarkText` 生成的是**正文层 WordArt**（`watermarkShape.name = "WordArt 2"`，`Anchor.Information(3) = 1`，即锚在第 1 页）。子代理为了让水印**每页都有**，删掉正文层水印、改在**页眉层**重建（`sec.Headers.Item(1)`），并同时处理 `DifferentFirstPageHeaderFooter`。

**建议**：说明里写清水印的落点与跨页行为，或直接改为页眉层实现。

### 成品完成度（协调方实测）

| 成品 | 状态 |
|---|---|
| `测试演示文稿.pptx` | **10/10 页已建**（计划 10 页）；第 1、2、3、4、6、7、9、10 页内容完整；**第 5、8 页只有标题**（ISS-54）；第 1 页背景超框（ISS-53） |
| `测试文字文稿.docx` | **完整**：封面（文档编号/密级/版本号/编制日期/编制单位）、114 段、2 张表、1027 字、多级大纲、页眉、页眉层水印、页脚 PAGE/NUMPAGES 域（"第 1 页 / 共 5 页"）、1 条批注、2 条修订痕迹、占位符全量替换（`{{VERSION}}`/`{{DATE}}` 0 残留）、已原地保存并导出 PDF |

## 复核与修复记录（2026-09-22，第三批）

### ✅ ISS-38 修复（排序不再假成功）

**改动**（`wps-addon/src/excel.js` `setFilterAndSort`）：旧式 `targetRange.Sort(key1, order1, ...)` 在本机 WPS 上静默 no-op。改为：
1. 先走 **`SortFields` 路径**（`SortFields.Clear/Add` + `SetRange` + `Header=1` + `Apply`）——上一轮子代理实测该路径有效；
2. **读回校验**：用新加的纯函数 `isSortedByRules(matrix, rules)` 判断排序后顺序是否真的满足要求（首行按表头跳过，`colIndex` 按区域内相对列号）；
3. 未通过则退回旧式 `Range.Sort` 再校验一次；
4. **两条路径都失败就抛错**，错误信息里带上已尝试的路径与"`colIndex` 是区域内相对列号"的提示——不再返回 `success`。

返回体新增 `sortApplied: { changed, attempts }`，调用方能看出是否真的改变了顺序。

### ✅ ISS-40 修复（筛选范围改为回读并告警）

**改动**：启用自动筛选后**回读** `sheet.AutoFilter.Range.Address()`，作为 `appliedFilterRange` 返回；与传入 `range` 不一致时追加 `warnings`（说明宿主会把筛选扩展到相邻数据块、如何用空行隔离）。不再让调用方以为筛选范围就是自己传的那个。

### ✅ ISS-54 修复（`content` 布局缺正文不再静默）

**改动**（`wps-addon/src/ppt.js`）：`generate_deck` 的 `content` 分支在没有 `bulletPoints` 时，往 `warnings` 里推一条带 `slideIndex` 的说明——本页只有标题、没有正文，请补 `bulletPoints` 或改用 `cards`/`chart` 布局。工具本来就返回 `layoutWarnings`，现在这一页也会出现在里面。

### 本批验证

`npm run build:wps-addon` + `node --check` 通过；生成物与源码一致；`typecheck 0`；全量测试 **76/76**。三条均为**已修待验**。

### ⚠️ 部署状态提醒（重要）

以上所有修复（ISS-01/28/39/42/47/38/40/54）**都只在仓库里**。运行中的 WPS 加载项仍是旧版，因此：
- 第三轮子代理测到的仍是**旧行为**，它们若再次报出 ISS-38/39/40/01，属**预期**，不是修复无效；
- 要真正验证，需要在**所有子代理收工后**重新部署加载项，再跑一遍对应复现步骤。

---

## 复核与修复记录（2026-09-22，第四批）

> 本批起因：`sheet-adv` 子代理在报告里**直接质疑我刚写的排序校验函数有 off-by-one**。我按"先验证再改"的原则处理，并顺带修正了排序路径本身。

### ✅ 质疑被推翻：`isSortedByRules` 没有漏最后一行（附测试）

**质疑原文**：「`isSortedByRules`（`excel.js:1627`）从 `i=2` 起、且漏最后一行，off-by-one 让校验恒真，本该报错的保护失效。」

**复核结论：不成立。**
- `i` 从 2 到 `matrix.length - 1`，比较的是 `(i-1, i)`，即从 `(1,2)` 一直到 `(n-2, n-1)` —— **最后一对相邻行是被比较的**；
- 从 `i=2` 起是正确的：第 0 行是表头（排序时传 `header=1`），不该参与比较。

**用测试钉死，而不是靠嘴**：把两个纯函数抽成 `wps-addon/src/sheet-sort.js`（沿用 `ppt-layout.js` 的 `@build-strip` 导出模式，构建时整段移除），新增 `tests/sheet-sort.test.ts` **10 项**，其中专设一条回归项：

```ts
test('isSortedByRules 必须比较**最后一对**相邻行（"漏最后一行"回归项）', () => {
  // 表头 + a, c, b —— 只有最后一对逆序
  assert.equal(isSortedByRules([['h'], ['a'], ['c'], ['b']], asc()), false);
});
```
实测 **10/10 通过**（含降序镜像、多列规则、非法列号、缺数据不抛错）。全量回归 **86/86**。

**但这条质疑带来了两个真实收益**（照做）：

1. **`Range.Sort` 在本机 WPS 上是方法不是对象** —— 子代理实测 `targetRange.Sort.SortFields` 抛 `Cannot read properties of undefined (reading 'Clear')`，而 `sheet.Sort` 才是可用的 Sort 对象。我原来的实现**只试了 `Range.Sort.SortFields`**，在这台机器上必然走进异常分支。**已改为三条路径依次尝试**（`sheet.Sort.SortFields` → `range.Sort.SortFields` → 旧式 `range.Sort(...)`），每条都读回校验，全不生效才报错；即使某条抛错也会读回（可能已部分生效）。
2. **校验函数变得可测** —— 原来埋在 IIFE 里谁也测不到，"校验恒真"这类缺陷只能靠人读代码发现。现在有 10 项测试守着。

### 本轮盘点新增问题

第三轮 4 个子代理（`ppt-probe` / `word-adv` / `sheet-adv` / `pipeline`）的完整报告在 [mcp-sweep/](p5/mcp-sweep/)：
- [02-ppt.md](p5/mcp-sweep/02-ppt.md) — 含**版本问题**的关键发现（磁盘新、进程旧）
- [03-word.md](p5/mcp-sweep/03-word.md) — Word 高级能力 383 行
- [04-sheet-advanced.md](p5/mcp-sweep/04-sheet-advanced.md) — 表格高级能力
- [05-e2e.md](p5/mcp-sweep/05-e2e.md) — 跨组件端到端

新增条目 **ISS-60 ~ ISS-74**（见汇总表）。其中两条最重（**原列的 ISS-67 已撤稿，见文末"撤稿记录"**）：

| 编号 | 为什么重 |
|---|---|
| ~~**ISS-67**~~ | **【撤销·非缺陷，2026-09-22 复核】** 原文：`wps_word_write_content` **吞掉所有小写字母 `a`**：发 `a ab abc banana A Aa 啊阿` 读回 ` b bc bnn A A 啊阿`，被"隔离到 `Paragraphs.Add(targetRange)` 路径"。**复核结论：原现象不成立**——那条"精确等值对照"的探针自己写了 `.replace(/[\r\a]/g, "")`，JS 里 `\a` 恒等转义等于字母 `a`，探针先删掉读回文本里所有 `a` 再比对；原始输出 `strippedMatch:[148]`（共 149 段）恰好证明内容**正确写入文末**。所谓的"隔离表"同一个坑，且其原始输出中 `Paragraphs.Add()+Range.Text` 实为**空串**，被误写成"完整保留"。**教训：验证方法本身出错也会指向产品** |
| **ISS-69** | 导出 PDF 返回 `success` + `savedPath`，但**磁盘上根本没有这个文件**——这条同时被上一轮 Word 子代理的 `run.log` 佐证（当时 `.scratch/word/export/` 里只有截图、没有 PDF） |
| **ISS-71** | `wps_word_capture_preview` **已实现但未注册**，调用报"未知工具"——正是 P2.5 台账里 6 条死分支的实例，且它恰好卡住 Word 侧视觉验收 |

### 版本时间线（2026-09-22 定死，读全部宿主结论的前提）

`ppt-probe` 的干净实验**对比出了重载前后的行为差异**，把版本之谜解开了：

| 阶段 | 运行中的构建 | 判据 |
|---|---|---|
| 会话开始 ~ 中途 | **09-21 13:05 构建** | `generate_deck` 缩放系数 = **1.0**（形状停在 720×405 设计原值，只占 960×540 页面的 75%）；返回体**缺** `pageWidth`/`layoutWarnings`，与 `addon-core.js.backup-1789967115643`（sha `3b34c6a5…`）逐字一致 |
| 中途之后 ~ 现在 | **09-22 11:51 部署的构建** | 加载项被重载后，**同参数**生成：背景正好 960×540、标题框 80/186.67/800、字号 36→48（sx=sy=4/3）；返回体**含** `pageWidth`/`layoutWarnings` |

**成因**：WPS 进程在 11:51 那次部署之前就启动过；部署只写磁盘**不重载进程**，所以"文件是新的、进程里是旧的"。中途某次加载项重载（子代理操作或宿主行为）让新构建生效。

**这对证据链的含义**：
1. **重载前**的宿主结论 → 描述的是 **09-21 构建**（比我们的改造还早）；
2. **重载后**的宿主结论 → 描述的是 **11:51 构建**（含 P3 重构，但**不含我本轮的修复**）；
3. **我本轮的全部修复仍未上线** —— 磁盘部署文件 `c30e548a…`，仓库构建物已是 `42e884b9…`；
4. 重载可能**打断了当时在跑的其他子代理**，个别异常行为需谨慎归因；
5. 凡经**源码复核**的结论（ISS-01/38/39/40/42/54 等）不受影响——当前源码里确认存在同样的缺陷。

**因此**：所有"已修待验"必须在**重新部署 + 重启 WPS** 之后再验，这是接下来唯一有效的基准。

## 第三轮盘点总结（4 个子代理全部完成）

| 代理 | 报告 | 核心产出 |
|---|---|---|
| `ppt-probe` | [02-ppt.md](p5/mcp-sweep/02-ppt.md) | **版本问题定案**；问题 B（content 无正文）稳定复现且根因在源码；问题 A 查明是旧构建行为；PPT 表格/形状/占位符/版式/批量 5 页 0.24s 均做成 |
| `word-adv` | [03-word.md](p5/mcp-sweep/03-word.md) | 样式/分节/横向页/书签/交叉引用/内容控件/文档属性做成；**4 处 success 但没生效**（含吞字母 `a`） |
| `sheet-adv` | [04-sheet-advanced.md](p5/mcp-sweep/04-sheet-advanced.md) | 透视表/条件格式五种/大数据量（5000 格 0.34s）/合并做成；**行隐藏假成功**等 6 条新问题 |
| `pipeline` | [05-e2e.md](p5/mcp-sweep/05-e2e.md) | **端到端链路走通**，三份成品落盘并经 OOXML 独立解包校验；**跨组件一致性 18/18 全对**；失败恢复无半成品、指纹不变 |

**新增条目 ISS-60 ~ ISS-88**（29 条）。本轮最重的四条（**ISS-67 已于 2026-09-22 复核撤稿，见文末"撤稿记录"**）：

| 编号 | 内容 |
|---|---|
| **ISS-77** | **目标锁语义不一致**：宿主 `shared.js` 的 `lockedTargets` 是加载项**进程级全局**，网关 `TargetLockStore` 是 `sessionId:host` 级 → 不传目标时行为不可预测（同一无参调用一次报错命中残留锁、一次静默指向别的代理正在用的文稿）。**这条解释了 ISS-02** |
| ~~**ISS-67**~~ | **【撤销·非缺陷】** 原文"`wps_word_write_content` 吞掉所有小写字母 `a`（现象确凿、根因待定）"**不成立**：探针自己用 `.replace(/[\r\a]/g, "")` 删掉了读回文本里的 `a`（JS 里 `\a` = 字母 `a`）。详情见文末"撤稿记录" |
| **ISS-80** | `insert_native_chart` 100% 失败且文案误导（宿主 `AddChart/AddChart2` 返回 `null` 不建形状，却报"数据配置未完成"） |
| **ISS-87** | `set_background` **串改全部页**（设第 1 页后第 2 页也变红） |

**也有很扎实的好消息**：跨组件一致性 18/18、失败注入 4 类全部响亮报错且不落半成品、三份成品文件指纹 before/after 一致、5000 格写入 0.34s 逐格正确。

## ✅ 第一批修复的真实宿主复验（2026-09-22 16:0x，**7/7 全部通过**）

**复验环境**（三件事都做了才算数）：
1. **加载项重新部署**：`npm run setup -- --addon` → 6 个组件目录全部更新，`addon-core.js` 哈希 `22636845359023ccffeb`，时间 16:03:02；
2. **WPS 重启**：加载项重载，三类文档都打开；
3. **桥接服务重建并重启**：`npm run build:main` → `dist/bridge/cli.cjs` 由 14:59 更新为含 `restoredScope` → `--stop` / `--start` → 新 pid 73823。

> ⚠️ **ISS-59 又复现了一次（第二例）**：加载项修好了，但 **`dist/bridge/cli.cjs` 还是 14:59 的旧构建**，导致桥接侧修复（ISS-47）第一次复验**失败**。重建 + 重启后台后才通过。
> **这坐实了 ISS-59 不是偶发**：本项目有**两个独立的代码落点**——WPS 加载项（需部署 + 宿主重载）与桥接后台（需 `build:main` + 后台重启），任何一处漏掉都会得到"改了没生效"。

| 编号 | 复验项 | 结果 | 关键读数 |
|---|---|---|---|
| ✔ ISS-01 | `values` + 空公式矩阵不得清空同批写入的值 | ✔ | 读回 `[['新A','新B']]` |
| ✔ ISS-38 | 排序必须真的生效 | ✔ | 读回 `[1,2,3,4]` 升序；`sortApplied.attempts = ["sheet.Sort.SortFields:ok:已生效"]` |
| ✔ ISS-39 | 不传 `columnRules` 时真的自适应 | ✔ | `results` 5 列、`mode: "usedRange"` |
| ✔ ISS-40 | 筛选范围回读 | ✔ | `appliedFilterRange: "$D$1:$E$5"`（与传入一致时 `warnings: []`） |
| ✔ ISS-42 | 合并 / 取消合并读回校验 | ✔ | `merged: true` → 取消后 `merged: false` |
| ✔ ISS-47 | 回滚文案与覆盖边界 | ✔ | `message: "成功恢复区域 $I$1:$I$2 的值与公式"`；`restoredScope: {values:true, formulas:true, styles:false, charts:false, structure:false}` |
| ✔ ISS-54 | `content` 布局缺正文必须告警 | ✔ | `layoutWarnings: [{slideIndex:2, reason:"content 布局未提供 bulletPoints：本页只有标题、没有正文。…"}]` |

**附带确认**：`sheet.Sort.SortFields` 确实是本机 WPS 上唯一有效的排序路径——`attempts` 里第一条就 `已生效`，**没有走到后两条兜底**。子代理的实测结论成立，第四条修复记录里改成"多路径依次尝试"是对的。

测试脚本：`/tmp/verify-excel.mjs`、`/tmp/verify-ppt47.mjs`（临时脚本，未入库）。复验建在 `verify_修复复验` 工作表；另新建了 `verify-iss54.pptx`。两者均可随时删除。

## 修复记录（第五批：向"假成功"整类开刀）

> 决策 #3 定的原则：**写工具执行后必须读回校验，未生效就报错**。本批把这条落到四个最典型的点上。

### ✅ ISS-60 行隐藏（逐行写 + 读回校验）

**根因**：行用 `sheet.Range("5:6").Hidden = true`，本机 WPS 上**静默 no-op**；而列用 `sheet.Columns.Item(n).Hidden` 是正常的——同一个函数里两条路径行为不一致。

**改动**：改为**逐行** `sheet.Rows.Item(r).Hidden = want`，写完**逐行读回**；任何一行与期望不符就抛错，错误里给出读回数组并提示检查工作表保护。

### ✅ ISS-68 Word 查找替换（补上缺失的纯查找分支）

**根因有三处**：
1. **纯查找分支根本不存在** —— `replaceText` 与 `replaceFormatting` 都没传时两个 `if` 都不进，`matchCount` 恒为 0，却返回"已找到并应用格式化"；
2. `catch (e) {}` **吞掉全部异常**；
3. 替换分支**忽略 `Execute` 的返回值**并无条件 `matchCount++`，"一处没命中"也报 `replaced_all`。

**改动**：补上纯查找/查找并格式化的遍历计数（5000 次上限防死循环）；替换分支**检查 `Execute` 返回值**，据此返回 `action: "replaced_all"` 或 `"no_match"`，文案区分"已替换 N 段"与"未找到、未做任何替换"；异常收进 `errors` 随响应返回，不再吞掉。

### ✅ ISS-69 PDF 导出（桥接侧落盘校验）

**根因**：加载项里 `doc.ExportAsFixedFormat(filePath, 17)` 之后**直接返回 success**，而宿主实测**不落盘**。加载项没有文件系统访问，校验只能放到桥接侧。

**改动**（`src/bridge/gateway/word.ts` 的 `saveDocument`）：拿到 `savedPath` 后由桥接（Node，有 fs）`existsSync` + 大小检查；不存在或 0 字节**直接报错**；成功则返回 `verifiedOnDisk: true` 与 `fileSizeBytes`。

### ✅ ISS-19 `delete_chart` 误删（破坏性操作先设卡）

**根因**：`leftCell` 换算成像素坐标后按 **±30px 邻近**匹配——多张图叠在同一位置（正是 ISS-18 的后果）时**全部命中**，实测一次删掉 14 张；且匹配优先级里 `leftCell` 排在 `chartTitle` 之前。

**改动**：先**收集候选**再决定删不删；`leftCell` 命中**多于一张**时**拒绝执行**并列出候选清单与坐标、提示改用 `shapeName`/`chartIndex` 或显式 `clearAll: true`；**一张都没匹配**时也报错（原来返回 `deletedCount: 0` 的假成功）；返回体新增 `matchedBy`。

### 本批验证

`build:wps-addon` + `node --check` 通过、生成物与源码一致、`build:main` 通过、`typecheck 0`、全量测试 **86/86**。

四条均为**已修待验**——需要**再部署一次加载项 + 再重启一次后台**（两个落点都要，见 ISS-59）。

### 复验清单（下次部署后依次跑）

| 编号 | 复验步骤 | 期望 |
|---|---|---|
| ISS-60 | 对某几行 `hide` → 读回 `Rows.Item(n).Hidden` | 全部 `true`；失败时报错而不是 success |
| ISS-68 | 传 `searchQuery` 不传替换 → 看 `matchCount` | 真实命中数；`action` 明确；异常不被吞 |
| ISS-68 | 传不存在的词 + 替换 → 看 `action` | `no_match`，文案"未找到、未做任何替换" |
| ISS-69 | 导出 PDF 到**不可写**目录 | 报错"保存未落盘"，不再返回假成功 |
| ISS-19 | 两张图叠放，用 `leftCell` 删其中一张 | **拒绝执行**并列出候选；改用 `shapeName` 可精确删 |
| ISS-19 | `clearAll: true` | 全部删除（显式确认路径仍可用） |

## 🔴 WPS 宿主崩溃（2026-09-22，使用者报告；已定位到崩溃报告）

使用者反馈"WPS 中途崩溃了几次"。查 `~/Library/Logs/DiagnosticReports/`，**今天 WPS 崩了 3 次**：

| 时间 | 信号 | 崩溃线程 | 调用栈特征 |
|---|---|---|---|
| 15:56:10 | `EXC_CRASH / SIGABRT`（Abort trap: 6） | 线程 51（未命名） | `-[NSView addSubview:]` → `objc_exception_rethrow` → `std::__terminate` → `abort`；`libqcocoa` + `QtWidgetsKso`（**UI 层**） |
| 16:10:24 | `EXC_BAD_ACCESS / SIGBUS`（KERN_PROTECTION_FAILURE） | 0 `CrBrowserMain` | `kso → etcore → etapi → **jsetapi → ksojscore** → jsapiservice → kshell → etmain` |
| 16:10:58 | 同上，**栈与偏移完全一致** | 0 `CrBrowserMain` | 同上 |

### ISS-89 深度属性反射会让 WPS 进程崩溃

**严重度**：**高**（宿主级崩溃，非普通报错）

**证据**：
- 16:10 两次崩溃**调用栈逐帧一致**（`kso+4133660`、`etcore+11346284`、`jsetapi+88516`…），说明是**可稳定复现**的操作触发的，不是随机不稳定；
- 栈里有 **`jsetapi` + `ksojscore`**（JS 引擎桥）→ `etcore`/`etapi`（**表格**核心），即"JS 属性访问扎进原生 ET 层"；
- 当时正在跑的是 `wps-api-map` 子代理的**逐表达式反射**：`reflect.js` 对每个对象 `for (const k in v)` 并**逐个 `v[k]` 求值**，表达式列表里包含 `Comment` / `CommentThreaded` / `CommentsThreaded` / `Sort.SortFields` / `AutoFilter.Filters` / `PivotCaches()` / `ListObjects.Item(1)` / `FormatConditions.Item(1)` 这类**未防御的原生 getter**。

**影响面**：`wps_execute_script` 允许执行任意 JS，**一次"无害的探测"就能把宿主搞崩**，且 AI 完全无法预判哪些成员是危险的。

**建议修法**（与 ISS-13"反射能力不足"是一对张力，要一起定）：
1. 反射类工具加**超时与逐项隔离**（每次只探一个表达式，崩溃即止损）；
2. 提供**经过验证的安全白名单**（哪些成员可安全读取、哪些会崩），而不是让调用方盲探；
3. `native-scripting.md` 里补"危险成员清单"（至少收录本次确认的）；
4. 长期：把常用反射结果**缓存成静态能力表**，减少运行时盲探。

### ISS-90 崩溃后宿主组件掉线，且没有恢复与提示

**严重度**：中高

**现象**：16:10 两次崩溃后，**Excel 组件从连接状态变为断开**（`excel: false`，Word/PPT 仍 `true`），但**桥接侧没有任何"宿主崩溃过"的信号**——调用方只会看到后续调用失败。

**建议修法**：桥接检测到组件从 connected 变为 disconnected 且非正常退出时，显式上报"疑似宿主崩溃"并给恢复指引。

### 备注：我们自己的 App 崩溃已修复，未再复现

`~/Library/Logs/DiagnosticReports/Office Agent Bridge-*.ips` 有 4 份，时间 **11:33 / 11:34 / 11:38 / 11:39**，签名都是 `EXC_BREAKPOINT / SIGTRAP` on `ThreadPoolForegroundWorker`（栈顶 `v8::String::NewFromOneByte` / `v8::RegExp::New`）——**正是此前定位并修复的浮窗崩溃**（移除 `transparent`/`titleBarStyle` 与 `app.getFileIcon()` 这个唯一的线程池原生异步调用）。
**修复后至今（15:39 起持续运行）没有新的 App 崩溃报告**，该修复判定为有效。

## 能力对照盘点（3 个子代理，报告 06/07/08）

> 目标：回答"我们还能补强哪些能力，让 AI 更精细地操控文档"。三份报告都在 [mcp-sweep/](p5/mcp-sweep/)。
> ⚠️ MS 侧**零实测**（Excel 加载项未连接）——凡涉及 Microsoft 的结论都是**源码可证**或**文档推断**，已在各报告里标注依据等级。

### 三条最有价值的结论

**① 覆盖率：没有一项能力三边全通。** 对 `EXCEL_METHODS`（30 项）的覆盖：**WPS 29/30**（缺 `update_chart`）、**Office.js 29/30**（缺 `rollback_cells`）、**COM 28/30**（缺 `update_chart`、`save_workbook`）。

**② WPS 宿主对象模型比 Office.js 更全。** 官方文档显示 WPS 表格对象模型是 VBA/COM 的**近乎全量克隆**：有 `AddSmartArt`、艺术字 `AddTextEffect`、`BuildFreeform`、`ThreeD/Shadow/Glow/Reflection/SoftEdge`、格式刷、`Application.Run`、`ExportAsFixedFormat`、45+ 事件、`Range.Characters`、`Names`、`FileSystem/PluginStorage/ApiEvent`。
→ **短板不在宿主，在我方实现**：表格侧形状/图片能力为零、冻结窗格依赖 `ActiveWindow`（跨簿会打错目标）、`update_chart` 缺失、图表定位是磅值像素而非单元格锚定。

**③ 补强方向（按报告给出的优先级）**：
1. **修跨宿主字段契约**（本轮最大发现，M1–M8 源码可证）——schema/gateway/normalizer 三层按 WPS 语义写，Office.js 读另一套字段名；
2. **补 MS 侧 Excel 矢量绘图域**（Office.js 宿主全套都在：几何形状/连接符/SVG/文本框/分组/导图，我们只实现了 `addImage`）；
3. **修 `capture_sheet_preview` 伪渲染**（见 ISS-96）。

### ISS-96 Office.js 侧预览是**伪造的渲染图**

**严重度**：高（**证据可信度**问题）

**现象**：MS 侧无原生图表时，`capture_sheet_preview` 用 **Canvas 按 `cellW=110` / `rowH=28` 硬编码合成**一张"看起来像表格"的图，而不是从真实文件渲染。

**影响面**：这与本次验收反复强调的"**渲染自查**"直接冲突——AI 以为自己看到了真实版面，实际看到的是一张**按固定参数画出来的示意图**，据此得出的版式结论可能与真实文件完全不符。**这比报错更危险**：报错会让人停下来，假图不会。

**建议修法**：COM 侧 `excel.ps1` 已有真实 `CopyPicture` 渲染可复用；MS 侧应改为真实渲染，或**明确标注"这是示意图，非真实渲染"**，绝不能让调用方误以为在看文件。

### ISS-97 COM 回退白名单过度声明

**严重度**：中高

`adapter` 把整张 `EXCEL_METHODS`（30 项）当作"可回退到 Windows COM"的白名单，而 COM 实际只覆盖 **28/30**。后果：某些方法在 Office.js 失败后会被判为"可回退"，实际 COM 侧也没有实现。

### ISS-93 补充：M1–M8 跨宿主字段错位清单（源码可证）

| 症状 | 后果 |
|---|---|
| `excel_set_data_validation`（microsoft） | 只认 `params.rule` → **先清空校验再什么都不设** |
| `excel_manage_sheet` 的 move / tab_color / protect / unprotect | **4 个 action 静默 no-op** |
| `excel_find_and_replace` | `text=""` → 每个非空单元格都命中 + `replaceAll("")` → **可能破坏内容**（已修，见 ISS-98） |
| `excel_get_charts` | 忽略全部选择器参数 |
| `excel_add_chart` | 静默忽略 5 个参数 |
| `excel_manage_rows_and_columns` | `targetType` 被忽略 → 想插列却插行 |

### ✅ ISS-91 / ISS-92 / ISS-98 修复（本批，Office.js 侧）

**ISS-91**（`office-addon/src/rpc.js`）：`list_conditional_formats` / `update_conditional_format` 原来和"新增条件格式"共用写处理函数——**调"读"会新增一条规则**。现已从写路径摘除并**显式报错**（Office.js 侧确无对应实现），错误里说明可改用 WPS 侧能力。

**ISS-92**（`rpc.js` + `excel/comment.js`）：`handleManageComments` 的 `action` 默认是 `"add"`，而 `list_comments` 走同一条路 → **调"列批注"会在 A1 插一条空批注**；`update_comment` 则掉进未知 action 的 `return { success: true }` **静默成功**。现改为：`list_comments` 强制 `action: "list"`；`update_comment` 显式报错；未知 action 一律抛错（不再静默成功）；`delete` 缺 id 也报错。

**ISS-98**（`excel/range.js`）：`find_and_replace` 的搜索文本兼容 `searchQuery`（网关字段名），并**在文本为空时直接抛错**——阻断 `replaceAll("", …)` 这条可能破坏内容的路径。

验证：`build:office-addon` + `node --check` 通过、生成物与源码一致、`typecheck 0`、全量测试 **86/86**。三条均**已修待验**。

## ✅ 第二批修复的真实宿主复验（2026-09-22 16:2x，**14/14 全通过**）

**复验前置**（"两个落点"都做了）：
1. `npm run setup -- --addon` → 4 个组件目录全部更新为 `d4c01c7b`，**与仓库构建物哈希一致**（16:23:30）；
2. `npm run build:main` + 后台 `--stop`/`--start` → 新 pid **77013**，产物含 `restoredScope` 与 `verifiedOnDisk`；
3. WPS 重启，表格/文字/演示三组件均 `connected=true`。

| 编号 | 复验项 | 结果 | 关键读数 |
|---|---|---|---|
| ISS-01 | `values` + 空公式不得清空 | ✔ | 读回 `[['新A','新B']]` |
| ISS-38 | 排序真实生效 | ✔ | 升序 `[1,2,3,4]`；`attempts = ["sheet.Sort.SortFields:ok:已生效"]` |
| ISS-39 | 不传 `columnRules` 也自适应 | ✔ | `mode: "usedRange"`，返回实际列宽 |
| ISS-40 | 筛选范围回读 | ✔ | `appliedFilterRange: "$D$1:$E$5"` |
| ISS-42 | 合并 / 取消合并读回 | ✔ | `merged: true` → `false` |
| ISS-47 | 回滚文案与边界 | ✔ | `"…的值与公式"` + `restoredScope` |
| ISS-54 | `content` 缺正文告警 | ✔ | `layoutWarnings: [{slideIndex:2, reason:"…未提供 bulletPoints…"}]` |
| **ISS-60** | 行隐藏（**独立读回**，不只看工具返回值） | ✔ | `hide` 后 `ws.Rows.Item(10..12).Hidden` = `10=true 11=true 12=true`；`unhide` 后全 `false` |
| **ISS-68** | 纯查找返回真实命中数 | ✔ | 查"概述" → `matchCount=3`（修复前恒 0） |
| **ISS-68** | 未命中的替换不谎报 | ✔ | `action: "no_match"`，文案 `"未找到 …，未做任何替换"` |
| **ISS-68** | 命中的替换 | ✔ | `action: "replaced_all"`，`matchCount=1` |
| **ISS-69** | 真实落盘 | ✔ | `verifiedOnDisk: true`，磁盘文件 **617,992 字节** |
| **ISS-69** | 导出到不可写目录 | ✔ | 明确报错 `保存未落盘：/System/verify-iss69.pdf 不存在。宿主返回成功不代表文件已写出…` |
| **ISS-19** | 两张叠图时 `leftCell` 拒绝批量删除 | ✔ | `leftCell 按像素邻近（±30px）匹配，本次命中 2 张图表，**已拒绝批量删除**以免误伤。命中清单：Chart 3(左360,上40)；Chart 4(左360,上40)…`；**拒绝后图表数不变**；显式 `clearAll` 仍可全删 |

> **两处复验方法教训**（记录备查）：
> 1. **ISS-60 第一轮"失败"是测试脚本的错**：上一个脚本结尾已经 `unhide`，第二个脚本读到的是取消后的状态。→ **验"写操作"必须在自己脚本内成对完成 hide→读→unhide，不能跨脚本。**
> 2. **ISS-19 第一轮没触发拒绝路径**：因为前一步已删掉一张图，只剩 1 张，`leftCell` 命中唯一目标属正常删除。→ **验"防护逻辑"必须先构造出触发条件（叠图 ≥ 2），否则会得到假通过。**

### Office.js 侧（ISS-91 / 92 / 98）：**只做了静态验证**

改的是 `office-addon/src/{rpc.js, excel/comment.js, excel/range.js}`，构建 + 语法 + 生成物一致性都过了，但**本机 Microsoft Excel 加载项未连接**，无法运行时验证。状态标为"已修·静态验证"，**不得当成已验证**。

## 验收证据与文档缺口处理批次（2026-09-22，ISS-06 / ISS-10 / ISS-16 / ISS-65）

> 执行记录：[p5/mcp-sweep/14-evidence-fixes.md](p5/mcp-sweep/14-evidence-fixes.md)。ISS-06 的三项比对已重跑留档（见 §ISS-06 修复记录），ISS-10 已产出 [p5/open-questions.md](p5/open-questions.md)，ISS-16 已补实测场景（见 §ISS-16），ISS-65 见下。

### ISS-65 `find_and_replace` 的 `results[].row/col` 是**区域相对偏移**而非工作表绝对行列号

**严重度**：低（不报错、结果可见，但调用方按返回值定位会**错位**）

**现象**：`wps_find_and_replace` / `excel_find_and_replace` 在 `results[]` 里返回的 `row` / `col`，是**相对于本次检索区域**的 1 基偏移，**不是工作表绝对行列号**。工具说明对此**一字未提**，参数名也没有 `relative` 之类的提示。

**实测复现（真实 WPS 宿主，`mcp-sweep/04-sheet-advanced.md` §S7）**：

| 检索区域 | 命中位置（绝对） | 返回的 `row`/`col` | 若按返回值去找会落到 |
|---|---|---|---|
| `G4:G15` | 第一条就在 **G4** | `row:1, col:1` | **A1**（错位 3 行 6 列） |
| `I4:I6` | 命中 **I5** | `row:2, col:1` | A2 |
| `B4:B15` | 命中 **B5** | `row:5, col:1` | A5 |

三例的方向一致：返回的 `row` 恰好等于 **命中行 − 区域首行 + 1**，`col` 同理（同一次盘点还单独复核过 `G4` 报 `row:1`）。**同一份返回体里 `address` 字段却是正确的绝对地址（如 `$G$4`）**，两个字段语义不一致，更容易误导。

**根因（源码，`wps-addon/src/excel.js:911,928-929`）**：

```js
const cell = range.Cells.Item(r + 1, c + 1);   // ← 对宿主 API 而言，区域相对寻址是正确用法
const addr = cell.Address;                      // ← address 是绝对地址，正确
matches.push({
  address: addr,
  row: r + 1,      // ← 但 row/col 直接把区域内的下标当成了「行列号」返回
  col: c + 1
});
```

即：**`r`/`c` 是遍历检索区域 `values` 矩阵时的下标**（`excel.js:872` 起的 `findAndReplace`），`r + 1` 只是把它转成 1 基，**从没加上区域首行/首列的偏移**。`Cells.Item(r+1, c+1)` 之所以正确，是因为 `range.Cells` 本身就相对 `range` 计数——两处用了同一个下标，但一个语义是"区域内偏移"，另一个被当成"绝对行列号"返回了。

**影响面**：调用方（尤其是 AI）拿到 `results[]` 后**很自然会把它当坐标**去 `read_range` / `format_cells`，于是操作到完全不相干的单元格。它比"报错"更危险的地方在于：**返回值看起来完全合理**（`row:1, col:1` 形式合法，数值也在表内），没有任何异常可循。ISS-65 定级"低"是因为**不破坏数据**，但在"AI 依赖读回结果做下一步"的链路上，它的误导性是实打实的。

**规避（当前版本，调用方侧）**：

1. **只信 `results[].address`**（绝对地址，如 `$G$4`，可直接喂给 `read_range` / `format_cells`）；**不要用 `row`/`col`**。
2. 若一定要还原绝对行列号，自行换算：`绝对行 = 区域首行 + row − 1`，`绝对列 = 区域首列 + col − 1`（区域首行/首列从你传入的 `searchRange` 里取）。
3. 对"必须先定位再改写"的流程，改成**先把 `searchRange` 收窄到目标行列**，让"区域内偏移"退化为"区域内第几格"，从而绕开换算。

**建议修法（不在此批次范围，供排期）**：`row`/`col` 改为返回**工作表绝对**行列号（或同时返回 `relativeRow`/`relativeCol` + `rangeStart`）；无论选哪种，工具说明里必须写明语义。**说明文案由工具说明批次负责**（见 [p5/mcp-sweep/09-tool-desc-fixes.md](p5/mcp-sweep/09-tool-desc-fixes.md) 第 75 行，已排入该批次）。

**验证方式**：构造三例上表场景，断言 `results[].row/col` 与 `address` 指向同一单元格；或说明中写明"相对 `searchRange` 的偏移"并给出换算示例。

**证据**：[p5/mcp-sweep/04-sheet-advanced.md](p5/mcp-sweep/04-sheet-advanced.md) §S7（第 134 行复现、第 206 行登记）；源码复核 `wps-addon/src/excel.js:911,928-929`。

---

## 复核记录：一条被推翻的误报（2026-09-22）

Microsoft 侧子代理在交付报告里提出："`excel_create_sheet` 的 schema 只收 `sheetName`，而网关处理器读 `args?.name`，**该工具当前无法真正建指定名字的表**"。

**复核结论：不成立。** 依据两条，均可复现：
1. **源码**：`src/bridge/gateway/excel.ts` 的 `createSheet` 读的是 `args?.sheetName`（并且 `if (!args?.sheetName) throw` 就是它的必填校验），全文没有 `args?.name` 这一读法——`args?.name` 只出现在图表/形状类工具里（作为别名）。
2. **真机**：对 Microsoft Excel 实调一次 `excel_create_sheet{host:"microsoft", sheetName:"probe_create_sheet"}` → 返回 `{"success":true,"name":"probe_create_sheet","position":4}`，**表确实建出来了**（随后已用 `excel_delete_sheet` 删除，工作簿恢复为原来的 4 张表）。

**教训**：子代理的结论要回代码或真机复核再采纳——本轮已累计推翻 **4** 条（ISS-41 `freeze_panes` 错位、ISS-47 "样式未恢复=缺陷"、本条、以及文末的 **ISS-67 撤稿记录**）。**误报若被直接采纳，会去"修"一个本来正常的功能。**

---

## 撤稿记录：ISS-67（2026-09-22，撤销·非缺陷）

**原结论（现已被推翻）**：`wps_word_write_content` **吞掉所有小写字母 `a`**——"发 `a ab abc banana A Aa 啊阿`，读回 ` b bc bnn A A 啊阿`"，被"精确等值匹配"确认，并被"隔离到 `Paragraphs.Add(targetRange)` + `Range.Text` 复合路径的宿主行为"。当时严重度按 **高** 记（静默改坏用户内容）。

**复核结论：原现象不成立，宿主从未吞字符，本仓库代码也没有任何删 `a` 的逻辑。** 是**验证探针自己**制造的假象：

```js
// 子代理当时的 verify_tool.js（逐字引用）
const t = String(D.Paragraphs.Item(i).Range.Text).replace(/[\r\a]/g, "");   // ← 这里删掉了所有 a
if (t === want) found.want.push(i);
if (t === stripped) found.stripped.push(i);      // stripped = want.replace(/a/g, "")
```

**已证实原因**：`/[\r\a]/` 在 **JavaScript** 正则里 `\a` 是**恒等转义**，等于字母 `a`（`\a` = BEL `0x07` 是 C/PCRE/.NET 的语义）。所以探针**先把读回文本里每个小写 `a` 删掉**再与原始串比较，必然得出"少了 `a`"的结论。这也解释了大写 `A` 与中文完好——删的就是字面 `a`。

**原始输出（逐字引用，来自子代理会话 `629ea463-…` 的 tool/result）**：

```
SENT to tool: '确认字符: a ab abc banana A Aa 啊阿'
tool success: True
{ "sentByPython": "确认字符: a ab abc banana A Aa 啊阿",
  "exactMatch": [], "strippedMatch": [148],
  "strippedForm": "确认字符:  b bc bnn A A 啊阿", "total": 149 }
```

`strippedMatch:[148]`（共 **149** 段）恰好证明内容**正确写入文末**（`location:"end"` 语义正确），`exactMatch:[]` 与"少 a"完全由探针那行造成。

**同时撤回的"第二轮隔离表"**：四个写法**每一个**都以同一行 `.replace(/[\r\a]/g, "")` 收尾，其**原始输出**为：

```
SENT: 'Xa a ab abc banana A Aa 啊阿'
Paragraphs.Add(targetRange)+Range.Text     -> X  b bc bnn A A 啊阿   （读回串 = S 去掉全部 a，即写入正确）
Paragraphs.Add()+Range.Text                -> （空串）
Paragraphs.Add(targetRange)+InsertAfter    -> （空串）
ContentEnd Range.InsertAfter               -> X  b bc bnn A A 啊阿   （同上，写入正确）
```

原表把第 2 行写成"✅ `S` 完整保留"，实际是**空串**；把第 1/4 行写成"❌ 吞 a"，实际是探针自己删的。**该表不能作为"哪些宿主 API 会吞字符"的依据。**

**已排除的可能性（方法可复核）**：

| 假设 | 排除方法 | 结果 |
|---|---|---|
| 本仓库源码里有删 `a` 的代码 | 把 `src/**`、`wps-addon/src/**`、`office-addon/src/**`、`scripts/**` 里 **1161 条正则字面量 + `new RegExp`** 逐条当"逐字符清洗器"实测 | **0 条**能把样本串的每个 `a` 删掉 |
| 历史版本曾用 `\a` | `git log --all -S` 全历史搜 `[\r\n\a]` 与 `\a` | **从未出现**（现存唯一出现处是本次新增的注释） |
| 构建时改坏转义 | 逐行比对 `wps-addon/addon-core.js` 与 `src/**` 的 `[\r\n\x07]` | 完全一致 |
| MCP/HTTP 传输吞字符 | 子代理 transport 回显 + 磁盘 15 份已部署 `addon-core.js` 与仓库产物同一 sha256 | 排除 |
| 宿主 `Range.Text`/`InsertAfter` 吞 `a` | 上表原始输出第 1/4 行：读回串只差被探针删掉的 `a` → 写入本身完好 | **不成立**（原核心猜测） |

**处置**：
1. 严重的**兜底**保留并加强：`wordWriteContent` 的读回校验从"只比长度"升级为**逐字比对**（等长改写同样报错），报错带出写入/读回原文，每行返回 `readBackMatches`。**没有改写入路径**——前提不成立，改动它属无据风险。
2. 全量结论与代码证据见 [agent-tests/word-fixes.md](agent-tests/word-fixes.md) §1。

**教训（本轮已出现 4 次同类模式）**：测试/探针**自身**出错，结论却指向产品。前三次：BGR 常量算错、`Item(name)` 假阳性、`$M$1` 断言错误；本条是第 4 次。**看到"宿主/产品行为诡异"时，先怀疑探针**：把探针的清洗、比较、断言逻辑逐行回读一遍，再决定要不要动产品代码。

## 待补充

剩余待修 **77 条**（高严重度 15 条）。Microsoft Excel 通道仍未连接；`06-wps-api-map.md` 的 Word / PPT 两组件测绘待补（需改用**逐项隔离 + 限流**的安全扫法，见 ISS-89）。
