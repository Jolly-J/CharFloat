# 09 · MCP 工具说明与契约修正（P5 批次）

> 任务：把 `issues.md` 中已证实的「工具说明 / 参数 schema」问题在**工具定义层**修掉。
> 写区：`src/bridge/tools/definitions/**`（`shared.ts` / `excel.ts` / `word.ts` / `ppt.ts` / `microsoft.ts` / `lock.ts` / `script.ts` / `audit.ts` / `index.ts`）。
> 状态口径：`待改` / `改中` / `已改` / `改不动（写区外）` / `不修（附理由）`。
> 契约快照差异见本文件 §快照差异。本文件随改动**边改边更新**（避免中断丢成果）。

## 0. 基线（改动前实测，`npx tsx .scratch/mcp-sweep/diag.ts`）

| 指标 | 值 |
|---|---|
| 工具数 | 91（诊断 2 / 目标锁 5 / `excel_*` 27 / `wps_*` 57 含 Word/PPT / `office_*` 5 / 审计 4 …） |
| 说明总字符 | 6549（均值 72） |
| 缺 `type` 的参数 | 9 处（见 ISS-31） |
| 联合类型写法 | `borders`、`shapeId` 等用 `oneOf`/`type:[…]` |
| 零必填的写工具 | 17 个（见 ISS-35） |
| 参数无说明 | `wps_get_audit_history` 10 个 + `wps_get_audit_record.auditId` + `office_unlock_target.component` + `wps_word_format_document.margins` 等 |

## 1. 逐条改动计划（编号 → 文件 → 改什么）

| 编号 | 文件 | 改什么 | 状态 |
|---|---|---|---|
| ISS-27 | `definitions/excel.ts` | 27 个 `wps_*` 表格工具说明加选型指引：`wps_*` 为 WPS 专用兼容名，同名 `excel_*` 是跨宿主统一入口（必传 `host`）；`wps_get_workspace_summary` 里写完整规则，其余工具用短标记；并写清 `wps_modify_rows_columns` 与 `wps_manage_rows_and_columns` 的区别（后者多 `set_size`） | 待改 |
| ISS-28 | `definitions/audit.ts` | `wps_rollback` 说明写明覆盖边界：只回滚 `patch_cells` 的值/公式；样式/图表/结构/Word/PPT/脚本**不可回滚**；回滚前校验、有后续改动会拒绝 | 待改 |
| ISS-30 | `definitions/audit.ts` + **`src/bridge/tools/audit.ts`（写区外）** | 审计族 4 个工具：`wps_rollback` 可在写区内补全；另外 3 个（`wps_get_audit_history` / `wps_get_audit_record` / `wps_clear_audit_history`，含 10 个无说明参数）定义在 `src/bridge/tools/audit.ts`，**不在写区** | 部分改不动 |
| ISS-20 | `definitions/excel.ts` | `wps_update_chart` 说明标注宿主限制：WPS 未实现（与 `contracts/host-methods.ts` 的 `HOST_IMPLEMENTATION_GAPS.wps=['update_chart']` 对齐），给出替代路径 | 待改 |
| ISS-33 | **`src/bridge/tools/excel.ts`（写区外）** | `excel_*` 后缀「Windows Microsoft Excel 尚待实机验收。」由该文件第 26 行派生，写区内无法改到 | 改不动（写区外） |
| ISS-31 | `definitions/excel.ts` / `word.ts` / `ppt.ts` | 补 `type`：`manage_rows_and_columns.index`（number∪string）、`find_and_replace.searchQuery`（string∪number）、`word_write_content.content`（string∪array）、PPT `shapeId`/`shapeId1`/`shapeId2`（string∪integer）；`borders` 的联合类型写法统一。**逐个验证接受集不变**（不收紧任何现有合法调用） | 待改 |
| ISS-32 | `definitions/microsoft.ts` | 装配层正则（`tools/index.ts:22-23`）删掉「实现 100% 任意操作无死角！」后留下悬空逗号。**不动正则**，改写 `office_execute_script` 定义文本，使清洗后仍是完整句；并全量扫描 91 条清洗结果确认无其它残句 | 待改 |
| ISS-34 | `definitions/word.ts` / `ppt.ts` / `excel.ts` | 枚举与说明逐项对齐：`preset` 补 `academic`；`read scope` 补 `paragraphs`/`tables`；`write type` 补 `code_block`；`stylePreset` 补 `clean_minimal`/`none`；PPT `chartType` 补 `bar_stacked`；`action` 补 `delete_shape`；组件枚举写上取值 | 待改 |
| ISS-35 | `definitions/*.ts` | 见 §2 决策：17 个工具的必需输入**全部是「多选一」集合**，无法在不破坏合法用法的前提下加单项 `required`；改为在说明首行写明「必传其一：…」，破坏性开关写明显式确认要求 | 待改（说明层） |
| ISS-36 | `definitions/*.ts` | 删营销句与样板（「彻底解决…痛点」「100% 完美」「杜绝像素漂移」「极大提升…」「图灵级」「超级引擎」等），字数改花在约束上；`excel_*` 的 864 字样板在写区外，无法删 | 待改（写区内） |
| ISS-37 | `definitions/excel.ts` / `audit.ts` | 报错文案在 `catalog.validateArgs`（写区外），改为在说明里补参数集：`wps_rollback` 只收 `auditId`；`wps_save_workbook` 不接受 `sheetName`；`wps_create_sheet`/`wps_delete_sheet` 的参数集 | 待改 |
| ISS-25 | `definitions/excel.ts` | `wps_capture_sheet_preview` 补参数优先级（`address` 优先于 `range`；WPS 宿主忽略 `chartName`/`name`）与返回结构（`imageBase64`/`imageMimeType`/`imageSizeBytes`） | 待改 |
| ISS-82 | `definitions/ppt.ts` | `layoutIndex` 写明是 **ppLayout 枚举**（实测 1=标题幻灯片、2=标题和文本、7=标题和图示/组织结构图、12=空白），**不是** `CustomLayouts` 序号；越界不报错；默认 12 | 待改 |
| ISS-83 | `definitions/ppt.ts` | `wps_ppt_manage_table.action` 从 enum 移除宿主未实现的 `set_table_data`，说明写明改数据用 `create_table(data)` 或 `set_cell_text` | 待改 |
| ISS-84 | `definitions/ppt.ts` | `data` 参数说明写明元素类型：**只接受字符串**，数字会被 schema 拒绝（`arguments.data[r][c]: 类型不正确`） | 待改 |
| ISS-88 | `definitions/ppt.ts` | `swap_shapes` 写明只交换 Top（垂直换位，Left 不变）；`align_shapes` 写明是「对齐到 `shapeIds` 中第一个可解析形状」；`shapeId*` 写明解析顺序：数字先按 `Shape.Id`、再按页内 1-based 索引，字符串按名称 | 待改 |

## 2. 需要使用者/协调方确认的三点

1. **ISS-30 / ISS-33 落在写区外**（`src/bridge/tools/audit.ts`、`src/bridge/tools/excel.ts`）。已向协调方报告并请求授权或转交。
2. **ISS-35 不加 `required`**：17 个工具的必需输入都是多选一（如 `add_chart` 的 `dataRange`∪`dataRanges`、`freeze_panes` 的三选一、`delete_chart` 的目标∪`clearAll`），单项 `required` 会破坏合法用法；用 `anyOf` 表达会改变 schema 兼容面。本批按「说明层写明必传其一」处理，schema 层留给协调方决策。
3. **附带的说明级纠正**（不在 18 条清单内，但属同一段文字里的失实承诺，已一并改）：`wps_search_cells` 只搜值不搜公式、`wps_add_chart` 的定位参数实际生效范围、`wps_delete_chart` 的邻居匹配规则、`wps_patch_cells` 的值/公式互斥。

## 3. 快照差异

（改动完成后填写）
