# MCP 工具缺口清单（宿主能做、但 MCP 没暴露给 AI）

> 状态：**已完成**（骨架 → 工具面枚举 → 宿主对照 → 缺口表 → 参数忽略清单 → 方法与局限）
> 范围：`src/bridge/tools/definitions/**`（工具面）× `wps-addon/src/**` + `office-addon/src/**`（宿主实现）
> 方法：以静态读码为主（引用一律给 `文件:行号`），未做真机验收；行号对应 2026-09-22 16:12 快照。
> 约束：本报告为只读产出，未改动任何仓库源码与既有文档；中间产物在 `.scratch/gap/`。
> 结论速览：**缺口 48 条**（主类 A 6 / B1 4 / B2 24 / C 7 / D 4，另 3 条跨类）；B1 死分支 3 个可零成本补；C 类 7 项只写不读。

## 0. 结论摘要

- **工具面**：91 个工具 = 诊断 2 + 统一 `excel_*` 入口 27 + WPS 表格 27 + Word 12 + PPT 9 +
  目标锁 5 + Microsoft 原生通道 3 + 脚本反射 2 + 审计回滚 4（§1）。
  工具面本身**不缺数量**，缺的是"写完能读回"与"能定位到具体对象"。
- **缺口 48 条**（§3.1 排序总表），主类计数：A 6 / B1 4 / B2 24 / C 7 / D 4；
  另有 3 条跨类（#11 A+D、#22 B2+C、#28 A+B2）——按"涉及该类"统计则为 A 8 / B 29 / C 8 / D 5。
- **最硬的一类 B1 缺口**：`wps_clear_range`、`wps_get_style_token`、`wps_word_capture_preview`
  三个分支**宿主实现与网关处理器都已写好**，只是没有 schema，AI 永远调不到
  （台账见 `tests/contract-consistency.test.ts:53-56`）。
- **影响 AI 自我验证的 C 类 7 项**：条件格式、冻结窗格、数据有效性、筛选状态、
  工作表保护/标签色、Word 页眉页脚/水印、透视表 —— 全部只返回 `success:true` + 入参回显（§3.3）。
  唯一做对"写后读回"的是 `format_cells` 的合并分支（`excel.js:511-526`）。
- **一个反向缺口**：`wps_update_chart` / `excel_update_chart`（`host=wps`）有工具但宿主无 RPC 分支，
  静态可判定必然失败（`compatibility-diff.md:73`）。
- **跨宿主不等价**：`excel_*` 的 `host` 参数给人"两边都能用"的印象，但 Word/PPT 在
  Microsoft 侧完全没实现，且 12 个表格方法未做参数适配（§2.4）。

## 1. 工具面枚举（按能力域分组）

统计口径：`docs/acceptance/2.1.0-p0p1/tools-snapshot.p5.json` 实测 **91 个工具**，
与 `src/bridge/tools/index.ts` 的 `assembleTools()` 顺序一致：
诊断 2 → 统一 `excel_*` 入口 27 → 网关原生 59 → 审计 3。

| 能力域 | 数量 | 来源 |
|---|---:|---|
| 诊断与会话 | 2 | `definitions/` 外，`tools/diagnostics.ts` |
| 统一表格入口 `excel_*` | 27 | `tools/excel.ts` 由 `EXCEL_METHODS` ∩ 网关工具派生 |
| WPS 表格 `wps_*`（不含 rollback） | 27 | `definitions/excel.ts` |
| Word `wps_word_*` | 12 | `definitions/word.ts` |
| PPT `wps_ppt_*` | 9 | `definitions/ppt.ts` |
| 目标锁 | 5 | `definitions/lock.ts`（3）+ `definitions/microsoft.ts`（2） |
| Microsoft 原生通道 | 3 | `definitions/microsoft.ts` |
| 原生脚本与反射 | 2 | `definitions/script.ts` |
| 审计与回滚 | 4 | `definitions/audit.ts`（1）+ `tools/audit.ts`（3） |
| **合计** | **91** | |

### 1.1 诊断与会话（2）

| 工具 | 职责 | 必填 |
|---|---|---|
| `bridge_get_capabilities` | 平台、连接状态、支持工具、验证程度、回滚限制 | — |
| `bridge_diagnose` | 检查 Bridge / WPS 连接与能力边界，不改文档 | — |

### 1.2 统一表格入口 `excel_*`（27，均额外必填 `host`）

由 `unifiedExcelTools()` 在 `wps_*` 表格工具上追加 `host: wps | microsoft` 派生，职责与同名的 `wps_*` 完全一致，
但**只在 `EXCEL_METHODS` 交集内派生**——这是下文 B 类缺口的结构性原因。

`excel_get_workspace_summary`、`excel_get_sheet_outline`、`excel_create_sheet`(sheetName)、`excel_delete_sheet`(sheetName)、
`excel_read_range`(address)、`excel_get_range_styles`(address)、`excel_search_cells`(query)、`excel_patch_cells`(address)、
`excel_format_cells`(address)、`excel_add_conditional_formatting`(address)、`excel_freeze_panes`(—)、
`excel_modify_rows_columns`(targetType,action,index)、`excel_auto_fit_columns`(—)、`excel_capture_sheet_preview`(—)、
`excel_add_chart`(—)、`excel_get_charts`(—)、`excel_update_chart`(—)、`excel_delete_chart`(—)、
`excel_create_pivot_table`(sourceRange,destCell)、`excel_set_filter_and_sort`(range)、
`excel_set_data_validation`(address,validationType)、`excel_manage_sheet`(sheetName,action)、
`excel_manage_rows_and_columns`(targetType,action,index)、`excel_manage_cell_comments`(action)、
`excel_find_and_replace`(searchQuery)、`excel_duplicate_sheet`(sourceSheetName,newSheetName)、`excel_save_workbook`(—)

### 1.3 WPS 表格 `wps_*`（27）

| 工具 | 职责 | 必填 |
|---|---|---|
| `wps_get_workspace_summary` | 已打开工作簿列表 / 工作表列表 / 当前鼠标选区 | — |
| `wps_get_sheet_outline` | UsedRange 边界 + 前 3 行表头样本 | — |
| `wps_create_sheet` | 新建并激活工作表 | sheetName |
| `wps_delete_sheet` | 删除工作表 | sheetName |
| `wps_read_range` | 切片读值与公式（可选数字格式） | address |
| `wps_get_range_styles` | 区域样式：summary 摘要 / cells 逐格 | address |
| `wps_search_cells` | 文本/公式关键词搜索定位 | query |
| `wps_patch_cells` | 原地改值/公式，带前后快照与 diff | address |
| `wps_format_cells` | 字体/色/对齐/行高/数字格式/边框/合并 | address |
| `wps_add_conditional_formatting` | 阈值高亮 / 数据条 / 色阶 | address |
| `wps_freeze_panes` | 冻结或解冻窗口窗格 | — |
| `wps_modify_rows_columns` | 整行整列插入/删除/隐藏/显示 | targetType,action,index |
| `wps_auto_fit_columns` | 自适应列宽 + 上下限 + 换行 | — |
| `wps_capture_sheet_preview` | 区域/图表截图（Base64，多模态自检） | — |
| `wps_add_chart` | 建原生矢量图表（含非连续数据源、Y 轴刻度） | — |
| `wps_get_charts` | 图表列表 / detail 系列与坐标轴 | — |
| `wps_update_chart` | 改图表位置、标题、图例 | — |
| `wps_delete_chart` | 删图表 / 清空全部（按名称、序号、标题、锚点） | — |
| `wps_create_pivot_table` | 聚合生成数据透视表 | sourceRange,destCell |
| `wps_set_filter_and_sort` | 自动筛选开关 + 多列排序 | range |
| `wps_set_data_validation` | 下拉列表 / 数值区间 + 提示与报错 | address,validationType |
| `wps_manage_sheet` | 重命名/移动/标签色/保护/解保护 | sheetName,action |
| `wps_manage_rows_and_columns` | 行列表增删隐显 + 行高列宽 | targetType,action,index |
| `wps_manage_cell_comments` | 批注增/读/删/清空 | action |
| `wps_find_and_replace` | 查找定位与批量替换 | searchQuery |
| `wps_duplicate_sheet` | 整表克隆（含格式/公式/图表） | sourceSheetName,newSheetName |
| `wps_save_workbook` | 落盘保存 | — |

### 1.4 Word `wps_word_*`（12）

| 工具 | 职责 | 必填 |
|---|---|---|
| `wps_word_create_document` | 新建空白或按模板建文档 | — |
| `wps_word_save_document` | 保存 / 另存 / 导出 PDF | — |
| `wps_word_close_document` | 关闭文档（可选保存） | — |
| `wps_word_manage_content` | 删段落/删表格/清空全文 | action |
| `wps_word_read_document` | 读段落、大纲、排版元数据、表格结构 | — |
| `wps_word_write_content` | 结构化写入（标题/正文/列表/引用/代码块） | content |
| `wps_word_format_document` | 段落级排版（公文/商务预设 + 自定义） | — |
| `wps_word_insert_table_of_contents` | 生成目录 | — |
| `wps_word_manage_table` | 表格透视/插入/写矩阵/单元格排版/增删行/合并 | action |
| `wps_word_review_and_comments` | 修订开关、全量接受/拒绝、批注增读 | action |
| `wps_word_page_layout_and_watermark` | 页眉页脚 + 文字水印 | — |
| `wps_word_find_and_replace` | 查找替换（通配符 + 替换格式） | searchQuery |

### 1.5 PPT `wps_ppt_*`（9）

| 工具 | 职责 | 必填 |
|---|---|---|
| `wps_ppt_read_presentation` | 大纲、页列表、文本要点、演讲者备注、页面尺寸 | — |
| `wps_ppt_get_slide_shapes` | 单页全部形状几何/文本/表格元数据/层级 | — |
| `wps_ppt_generate_deck` | 按 JSON 大纲批量生成整套胶片 | slides |
| `wps_ppt_manage_slides` | 增/删/移动/克隆/设背景 | action |
| `wps_ppt_manage_table` | 建表/读表/单元格文字/主题配色 | action |
| `wps_ppt_add_business_cards` | 2/3/4 栏商业卡片自动排版 | cards |
| `wps_ppt_insert_native_chart` | 插入原生图表并绑定数据 | categories,series |
| `wps_ppt_manage_shapes_and_media` | 文本框/形状/位置尺寸/互换/层级/对齐/删除 | action |
| `wps_ppt_capture_slide_preview` | 单页导出图片供多模态自检 | — |

### 1.6 目标锁（5）、Microsoft 原生通道（3）、脚本反射（2）、审计（4）

| 工具 | 职责 | 必填 |
|---|---|---|
| `wps_lock_target_document` | 强隔离锁定目标文档，防窗口漂移 | component,targetName |
| `wps_unlock_target_document` | 解除锁定 | — |
| `wps_get_locked_status` | 查询锁状态与已打开文件列表 | — |
| `office_lock_target` | Microsoft 侧锁定 | component,targetName |
| `office_unlock_target` | Microsoft 侧解锁 | — |
| `office_get_status` | 探测 Microsoft 进程与已打开文档 | — |
| `office_execute_script` | 走 JXA/AppleScript/PowerShell 的原生脚本 | component,script |
| `office_capture_slide_preview` | Microsoft PPT 单页快照 | slideIndex |
| `wps_execute_script` | WPS 原生 JS 逃生舱（万能兜底） | code |
| `wps_inspect_api` | 运行时反射探测宿主对象成员 | expression |
| `wps_rollback` | 按审计 ID 一键撤销修改 | auditId |
| `wps_get_audit_history` | 分页筛选修改记录 | — |
| `wps_get_audit_record` | 按 ID 读完整记录与快照 | auditId |
| `wps_clear_audit_history` | 清空本地留痕 | — |

> **逃生舱的存在改变了缺口判定的性质**：任何"宿主 API 支持但没工具"的能力，
> 理论上都能用 `wps_execute_script` / `office_execute_script` 触达。因此本报告的 B 类
> 按"证据硬度"分两级：
> - **B1（硬）**：加载项里已有专用 RPC 分支 / 网关已有 handler，只是没给 schema —— 零成本补齐即可用；
> - **B2（软）**：宿主 API 面支持，但 RPC 与工具都没有，需新写宿主实现。

## 2. 宿主实现对照

### 2.1 路由全景（决定"缺口"到底是什么性质的缺口）

```
MCP 工具 → gateway HANDLERS(registry) → callOffice(method, params)
             ├─ host=wps       → WPS 加载项 RPC 分发（wps-addon/src/dispatch.js）
             └─ host=microsoft → normalizer 参数适配 → Office.js 加载项（office-addon/src/rpc.js）
                                 └─ 失败且平台=win32 且 method ∈ EXCEL_METHODS 才可能回退原生 COM
```

三条硬边界（都在代码里可验证）：

1. **Word / PPT 工具在 Microsoft 宿主上没有实现**。`office-addon` 只有 `src/excel/**`……
   `rpc.js:200` 的 `default` 对未映射方法直接抛 `Office.js 暂未映射该工具：${method}`；
   而 `errors.ts:107` 规定不在 `EXCEL_METHODS` 内的方法**不允许**回退原生通道。
   → 21 个 `wps_word_*` / `wps_ppt_*` 工具在 `host=microsoft` 时是死路，只能靠 `office_execute_script`。
2. **`excel_*` 统一入口不是"跨宿主等价入口"**。`tools/excel.ts:20-35` 只做改名 + 追加必填 `host`；
   参数形状适配全在 `normalizer.ts`，而它只覆盖 14 个方法（`normalizer.ts:35-201`），
   `default:`（`normalizer.ts:204`）把其余方法**原样透传**给 Microsoft 加载项 → 见 §2.4。
3. **专用工具缺失 ≠ 宿主不能做**。`wps_execute_script` / `office_execute_script` 是万能逃生舱，
   所以下面的 B 类要按"是否已经写好 RPC/处理器"分级（B1 零成本 / B2 需新写宿主实现）。

### 2.2 已实现但"AI 永远调不到"的死分支（B1，最高性价比）

`tests/contract-consistency.test.ts:53-56` 自认的债务台账，与 `gateway.ts:44-134` 的 `HANDLERS` 实测一致：

| 分支 | 网关处理器 | 宿主实现 | 能力 | 为什么值得补 |
|---|---|---|---|---|
| `wps_clear_range` | `gateway.ts:68` → `gateway/excel.ts:59` | `excel.js:104`（`range.Clear()`） | 清空区域 | AI 重置脏数据块；现在只能删行或用 patch 写空字符串（留下格式残留） |
| `wps_get_style_token` | `gateway.ts:67` → `gateway/excel.ts:51` | `excel.js:7` | 读原表设计语言（字体/字号/表头底色） | AI 让新页沿用现有报表视觉规范，避免配色漂移 |
| `wps_word_capture_preview` | `gateway.ts:113` → `gateway/word.ts` | `word.js:932`（导出 PDF） | Word 页面视觉快照 | **Word 是唯一没有任何预览/自检手段的组件**，Excel 与 PPT 都有截图工具 |
| `wps_ppt_add_chart` | `gateway.ts:126`（别名） | `ppt.js:461` | 与 `wps_ppt_insert_native_chart` 同一实现 | 仅兼容名，价值低 |
| `wps_reload_addon` | `gateway.ts:130` | `script.ts` | 重载加载项 | 运维向，AI 不需要 |
| `wps_eval_code` | `gateway.ts:133` | `script.ts` | 与 `execute_script` 重复 | 价值低 |

另有一组"路由表声明了、但没有工具也没有网关处理器"的方法（`compatibility-diff.md:58` 的 `declaredNotCallable`）：
`clear_range`（与上面第 1 条同源）、`insert_dimension`（`excel.js:1026`，
能力已被 `manage_rows_and_columns` 的 insert 覆盖，属冗余）、`rollback_cells`（`excel.js:1003`，审计回滚内部用）。

### 2.3 宿主 API 面 vs 工具面（WPS）

用"宿主文件里实际出现的 API 成员"做证据，得出各组件的**能力空集**。

**Excel（`excel.js` 1857 行）**——出现：`Validation`(10) `Sort`(10) `Shapes`(10) `FormatConditions`(4)
`FreezePanes`(3) `SplitRow`(2) `Tab`(1) `Protect`(1) `PivotCaches`(1)。
**完全没出现**（宿主支持、MCP 内外皆无）：
`Hyperlinks`、`Names`、`PageSetup`、`Outline`/`Group`（分级显示）、`Locked`/`FormulaHidden`、
`TextToColumns`、`RemoveDuplicates`、`ListObjects`（结构化表格）、`PrintArea`、`AutoFill`、
`PivotTables`（只 Create 不读）、`ExportAsFixedFormat`/`SaveAs`（导出 PDF/另存）、`Cells.Comment` 之外的 `Comments`（线程批注）。

**Word（`word.js` 955 行）**——出现：`Range`(85) `Tables`(26) `Paragraphs`(16) `PageSetup`(10)
`Comments`(5) `Styles`(3) `Words`(2) `Find`(2) `ExportAsFixedFormat`(2) `Bookmarks`(2)
`TablesOfContents`(1, 只 Add) `Shapes`(1, 只水印) `Sections`(1, 只有 Item(1)) `Headers`(1) `Footers`(1)
`ListFormat`(1) `SaveAs2`/`SaveAs`(1)。
**完全没出现**：`Revisions`（只有 Accept/RejectAll）、`InlineShapes`（插图）、`Hyperlinks`、
`Fields`、`ContentControls`、`Bookmarks.Add`（只能读不能建）、`TablesOfContents.Update`（目录不可刷新）、
`Styles` 的自定义/套用、多节 `Sections`、`Range.HighlightColorIndex`（高亮）。

**PPT（`ppt.js` 1225 行）**——出现：`AddTextbox`(9) `NotesPage`(6) `AddShape`(5) `Delete`(3)
`Background`(2) `MoveTo`/`Export`/`Duplicate`/`AddTable`/`AddLine`/`AddChart`(各 1)。
**完全没出现**：`AddPicture`、`AddConnector`（连线）、`AddSmartArt`、`Group`/`Ungroup`、
`Layout`/`CustomLayout`（套版式）、`SlideShowTransition`（隐藏页）、`TimeLine`（动画）、
`Comments`、`Hyperlinks`、`Design`（母版）、`SaveAs`、图表集合的读取/修改。

**Microsoft（`office-addon/src/rpc.js` 77 个 RPC 方法）**是最大的一块"已实现、未暴露"资产，
详见 §3 的 B2 组：`list_named_items` `update_named_item` `get_document_properties`
`update_document_properties` `copy_range` `set_hyperlink` `insert_image` `list_shapes` `update_shape`
`clear_cells` `get_used_range` `calculate` `create_table` `update_table` `update_pivot_table`
`apply_filter` `sort_range` `list_sheets` `add_sheet` `update_sheet` `copy_sheet` `rename_sheet`
`get_workbook_info` `get_workbook_state` `export_chart_image`……

### 2.4 `host=microsoft` 时代码与 schema 的错位（12 个工具）

`normalizer.ts` 只适配了 14 个表格方法；下列 12 个 `excel_*` 工具走 `default:` 原样透传：

`excel_get_range_styles`、`excel_create_sheet`、`excel_delete_sheet`、`excel_capture_sheet_preview`、
`excel_update_chart`、`excel_create_pivot_table`、`excel_set_filter_and_sort`、`excel_set_data_validation`、
`excel_manage_sheet`、`excel_manage_cell_comments`、`excel_find_and_replace`、`excel_duplicate_sheet`

其中两类后果（**静态可判定，未实机复现**）：

| 工具（host=microsoft） | 加载项实际读的参数 | 实际后果 |
|---|---|---|
| `excel_set_data_validation` | `params.rule`（`range.js:231`） | 先 `dataValidation.clear()` 再无条件不设置 → **清除原有校验且不报错**，返回 `success:true` |
| `excel_manage_sheet` | action ∈ copy/duplicate/rename/hide/show/color（`sheets.js:56-77`） | 传 `move`/`protect`/`unprotect` 全部落到 if-chain 之外 → **静默返回 success 但什么都没做** |
| `excel_manage_sheet`（反向） | schema 只允许 rename/move/tab_color/protect/unprotect | 宿主支持 `hide`/`show`（`sheets.js:71-74`）却**传不进去** → 工作表隐藏能力双向不通 |
| `excel_set_filter_and_sort` | `params.address`、`action`、`sortColumn`（`filter-sort.js:6-27`） | schema 给的是 `range`/`enableAutoFilter`/`sortRules` → `getRange(undefined)` 报错 |
| `excel_create_pivot_table` | `sourceAddress`/`destinationAddress`（`pivot.js:7-8`） | schema 给 `sourceRange`/`destCell` → 源区域解析为 undefined |
| `excel_manage_rows_and_columns`（经 normalizer 改名 `modify_rows_columns`） | `dimension`∈rows/columns，action 仅 insert/delete（`range.js:149-190`） | schema 的 `targetType:"column"` **被忽略**（`dimension` 默认 `rows`）→ 想插列却插了行；`hide`/`unhide`/`set_size` 静默 success 无效 |
| `excel_update_chart`（host=wps） | `dispatch.js` 无 `update_chart` 分支 | 静态可判定必然失败（`HOST_IMPLEMENTATION_GAPS.wps`，`compatibility-diff.md:73`） |

对照：`excel_add_chart` 是唯一**做对了双向适配**的例子——`normalizer.ts:164-167` 把 `position.*`
补成顶层 `left/top/width/height` 给 Microsoft，而 WPS 宿主只认 `position.*`（见 §4）。

### 2.5 两个"看着能用、其实危险"的宿主别名

| 位置 | 现象 | 风险 |
|---|---|---|
| `office-addon/src/rpc.js:144-146` | `list_conditional_formats` / `update_conditional_format` 都路由到 `handleAddConditionalFormatting` | 想**读**条件格式会**新增**一条 `ruleType` 缺省为 `cell_value`、`formula1` 缺省为 0 的规则（`format.js:153-176`） |
| `office-addon/src/rpc.js:193-195` | `list_comments` / `update_comment` 都路由到 `handleManageComments`，而它 `action = params.action \|\| "add"`（`comment.js:7`） | 想**列**批注会在 A1（`comment.js:10` 默认 `"A1"`）**插入一条空批注** |

### 2.6 审计与回滚的实际覆盖

全仓库只有一处 `auditStore.addRecord`：`gateway/excel.ts:131`，即 `wps_patch_cells`。
快照只含 `values` + `formulas`（`excel.js:346-357`）。
→ **AI 的格式修改、批注、图表增删、删表、条件格式、冻结窗格全都没有审计、不可回滚**，
`wps_rollback` 只对"写值/写公式"有效。

## 3. 缺口表（按 AI 操控文档价值降序）

### 3.0 分类口径

| 类别 | 含义 | 判定依据 |
|---|---|---|
| **A 类·粒度不够** | 有工具，但只能粗粒度操作 | 宿主 API 支持更细定位，schema 却没给参数 |
| **B 类·完全没暴露** | 宿主能做、MCP 没有工具 | B1 = RPC/handler 已实现（零成本补 schema）；B2 = 宿主 API 支持但需新写实现 |
| **C 类·只写不读** | 写得进、读不回 | AI 无法自我验证，只能靠截图猜 |
| **D 类·只读不写** | 读得到、改不了 | 读到的问题无法定点修复 |

**主类计数**：A 6 / B1 4 / B2 24 / C 7 / D 4 = 45，另有 3 条同时属两类（#11 A+D、#22 B2+C、#28 A+B2），
合计 **48 条**。

### 3.1 排序总表

| # | 宿主能做的操作 | 现在是哪个工具 | 缺口 | 类 |
|---|---|---|---|---|
| 1 | 读出区域上的条件格式规则（类型/阈值/颜色/覆盖范围） | 只有 `wps_add_conditional_formatting`（`excel.js:536`） | 无任何读取入口 | C |
| 2 | 读出窗口冻结状态（`FreezePanes`/`SplitRow`/`SplitColumn`） | 只有 `wps_freeze_panes`（`excel.js:611`） | 无读取入口 | C |
| 3 | 把批注钉在"第 5 段"或某段文字上 | `wps_word_review_and_comments` 只能钉在**当前选区**（`word.js:755`） | 缺定位参数 | A |
| 4 | 把 Word 页面渲染出来给多模态模型看 | `wps_word_capture_preview` 已实现（`word.js:932`），**无 schema** | 死分支 | B1 |
| 5 | 读回数据有效性（下拉列表项/数值区间/提示语） | 只有 `wps_set_data_validation`（`excel.js:1770`） | 无读取入口 | C |
| 6 | 读回已写入的页眉/页脚/水印 | 只有 `wps_word_page_layout_and_watermark`（`word.js:779`） | 无读取入口 | C |
| 7 | 读回当前筛选条件与排序状态 | `wps_set_filter_and_sort` 只回显 `appliedFilterRange`（`excel.js:1642`） | 读不到筛选条件 | C |
| 8 | 读段落样式名/行距/首行缩进/段间距 | `wps_word_read_document` 只给 bold/italic/size/name/align（`word.js:189-198`） | 粒度不够 | A |
| 9 | 区域超链接读写 | 无 | Microsoft 已实现 RPC（`range.js:217`），WPS 无 | B2 |
| 10 | 命名区域读写 | 无 | Microsoft 已实现（`workbook.js:95/105`） | B2 |
| 11 | 改已有页里的**一句话**而不动整框格式 / 写已有页的备注 | `update_shape` 只能整框替换（`ppt.js:692`）；备注只在 `generate_deck` 里能写（`ppt.js:366`） | 粒度不够 + 只读不写 | A/D |
| 12 | 清空区域（值+格式），不留格式残渣 | `wps_clear_range` 已实现（`excel.js:104`），**无 schema** | 死分支 | B1 |
| 13 | 读出现有报表的设计语言（字体/字号/表头底色） | `wps_get_style_token` 已实现（`excel.js:7`），**无 schema** | 死分支 | B1 |
| 14 | 插图 / 读形状与图片位置 | 无 | Microsoft 已实现（`shape.js:4/19/30`） | B2 |
| 15 | 隐藏/显示工作表 | `wps_manage_sheet` 无此 action（`excel.js:1837`）；Microsoft 有 `hide`/`show`（`sheets.js:71-74`）但 schema enum 不含 | 双向不通 | B2 |
| 16 | 只加粗/高亮命中某关键词的文字 | 宿主已实现 `searchQuery`/`searchQueries`（`word.js:378-421`），schema 未暴露 | 能力被隐藏 | B1 |
| 17 | 区域复制/移动（`copyFrom`） | 无 | Microsoft 已实现（`range.js:206`） | B2 |
| 18 | 导出 PDF / 设置页面与打印区域 | 无（`excel.js` 全文无 `PageSetup`/`ExportAsFixedFormat`） | 交付类空白 | B2 |
| 19 | 行/列分组分级显示（折叠明细） | 无（`excel.js` 无 `Outline`/`Group`） | 无入口 | B2 |
| 20 | 单元格锁定 + 公式隐藏（保护模板） | `wps_manage_sheet` 只能整表 protect（`excel.js:1837`） | 无法指定可编辑区 | B2 |
| 21 | 结构化表格 Table（自动扩展/结构化引用） | 无 | Microsoft 已实现（`table.js:4/17`） | B2 |
| 22 | 透视表读取 / 刷新 / 删除 | 只有 `create`（`excel.js:1554`） | 只建不读不改 | B2/C |
| 23 | 删除重复项 / 文本分列 / 自动填充 | 无 | 数据清洗三件套缺失 | B2 |
| 24 | 建书签 / 插入图片 | `write_content` 能**按**书签定位（`word.js:267`）但不能**建**书签；`InlineShapes` 未出现 | 半残 | B2 |
| 25 | 插入超链接 | 无（`word.js` 无 `Hyperlinks`） | 无入口 | B2 |
| 26 | 读修订列表 / 逐条接受拒绝 | 只有全量 Accept/Reject（`word.js:745-752`） | 粒度不够 | A |
| 27 | 删除/回复批注 | 只有 add/list（`word.js:753-773`） | 写完不能撤 | B2 |
| 28 | 刷新目录 / 读页面设置 / 多节页眉页脚 | `TablesOfContents` 只 Add（`word.js:537`）；`Sections.Item(1)` 硬编码（`word.js:782`） | 目录不可刷新 + 只覆盖第 1 节 | A/B2 |
| 29 | 表格加列/删列 | Word `manage_table` 只有 add_row/delete_row（`word.js:704-717`） | 缺 action | B2 |
| 30 | 查找命中位置（而不只是命中数） | `word_find_and_replace` 只返回 `matchCount`（`word.js:920-929`） | 读不到位置 | A |
| 31 | 插入图片/Logo 到幻灯片 | 工具名 `manage_shapes_and_media` 承诺 media，但 `AddPicture` 未出现 | 名实不符 | B2 |
| 32 | 形状分组/取消分组 | 无（`ppt.js` 无 `Group`） | 无入口 | B2 |
| 33 | 连线/箭头连接（流程图） | 无（`ppt.js` 只有 `AddLine`，无 `AddConnector`） | 无入口 | B2 |
| 34 | 给已有页套用版式/母版 | 无（`ppt.js` 无 `Layout`/`CustomLayout`） | 只能手摆坐标 | B2 |
| 35 | 隐藏幻灯片 | 无（`ppt.js` 无 `SlideShowTransition`） | 无入口 | B2 |
| 36 | 读/改已有幻灯片里的图表 | `get_slide_shapes` 只给 `hasChart` 布尔（`ppt.js:533`） | 图表不可读不可改 | B2 |
| 37 | PPT 表格加行/删行/合并单元格 | 只有 create/read/set_cell_text/style（`ppt.js:937/1086/1121/1150`） | 缺 action | B2 |
| 38 | 读回形状内文字的 run 级格式 | `get_slide_shapes` 文本截断 200 字、单一 fontSize（`ppt.js:548-549`） | 改完无法核对 | A |
| 39 | 文档属性（标题/作者/关键词）读写 | 无 | Microsoft 已实现（`workbook.js:118/126`） | B2 |
| 40 | 工作簿另存为/关闭/重算控制 | 只有 `save_workbook` | 缺 SaveAs/Close/Calculate | B2 |
| 41 | 改图表位置/标题/图例（WPS） | `wps_update_chart` 存在但无宿主分支 | 有工具无能力 | D |
| 42 | 设置/移动当前选区（把用户视线引到目标） | `get_workspace_summary` 能读选区 | 不能写选区 | D |
| 43 | 改 PPT 页面尺寸/方向 | `ppt_read_presentation` 能读 PageSetup | 不能写 | D |
| 44 | 保存/关闭 Microsoft 文档 | `office_get_status` 能列出打开的文档 | 无保存/关闭工具 | D |
| 45 | 回滚格式/批注/图表等非值类修改 | `wps_rollback` 仅覆盖 `patch_cells`（`gateway/excel.ts:131`） | 安全网太窄 | C |
| 46 | 清空条件格式（跨区域） | 只有 `clearExisting` 作用于传入区域（`excel.js:557`） | 无独立清理入口 | C |
| 47 | 线程批注/批注富文本 | `manage_cell_comments` 只有传统批注（`excel.js:734`） | 语义降级 | B2 |
| 48 | 两个重复工具 schema 不一致 | `wps_modify_rows_columns` 与 `wps_manage_rows_and_columns` 同一宿主实现（`dispatch.js:80` 与 `116` 都指向 `modifyRowsColumns`） | 工具面冗余、易误用 | A |

### 3.2 前 20 条：具体补什么工具 + 参数

> 位置一律写"宿主实现"（真正要改的地方）与"schema 位置"。
> 行号取自 **2026-09-22 16:12** 的工作区快照（`wps-addon/src/excel.js` 正在被并行任务修改，见 §5）。

**① `excel_get_conditional_formatting`（C，最高优先）**
- 参数：`address`（必填）、`sheetName`、`workbookName`、`detail`（bool，是否逐条展开 `formula1/formula2/operator/colors/priority/stopIfTrue`）、`maxRules`
- 宿主实现：WPS 侧新增 `getConditionalFormatting`，遍历 `range.FormatConditions`（现有代码只在 `excel.js:570/581/587` 创建）；Microsoft 侧复用 `list_conditional_formats`，但必须**先修 `office-addon/src/rpc.js:144-146` 的错误别名**（现在指向写操作）
- 场景：AI 做完看板后核对"E5:E20 到底有没有挂上 <90% 标红"，而不是靠截图猜颜色

**② 冻结窗格/视图状态读回（C）**
- 方案：并入 `wps_get_sheet_outline` 返回值，新增 `view: { freezePanes, splitRow, splitColumn, frozenRows, frozenColumns }`；或独立 `excel_get_sheet_view`
- 宿主实现：`excel.js:611` 同处读 `app.ActiveWindow.FreezePanes/SplitRow/SplitColumn`（现已写但从不读）
- 场景：AI 反复调 `freeze_panes` 会互相覆盖，需要先读回再决定"设几行"

**③ Word 批注定位（A）**
- 方案：`wps_word_review_and_comments` 的 `add_comment` 增 `target` ∈ `selection|paragraph|range|text`，配 `paragraphIndex` / `paragraphRange:[start,end]` / `anchorText`
- 宿主实现：`word.js:755` 现在写死 `wordApp.Selection.Range`；改为按 `doc.Paragraphs.Item(n).Range` 或 `doc.Content.Find` 命中范围取 Range
- 场景：AI 做审核员时"给第 12 段的金额加批注：与附件不符"，现在只能要求用户先用鼠标选中

**④ 暴露 `wps_word_capture_preview`（B1，零成本）**
- 参数：`documentName`、`pageRange`（可选）、`outputPath`（由 bridge 注入，参见 `ppt_capture_slide_preview` 的 `params.outputPath` 约定，`word.js:935`）
- 实现位置：schema 加到 `definitions/word.ts`；从 `tests/contract-consistency.test.ts:53-56` 台账中移除
- 场景：Word 是唯一没有视觉自检的组件（Excel 有 `capture_sheet_preview`、PPT 有 `ppt_capture_slide_preview`），公文排版只能"盲写"

**⑤ `excel_get_data_validation`（C）**
- 参数：`address`（必填）、`sheetName`、`workbookName`
- 宿主实现：读 `range.Validation.Type/Operator/Formula1/Formula2/InCellDropdown/InputTitle/InputMessage/ErrorTitle/ErrorMessage`（`excel.js:1770` 同文件已写这些属性，只是没有读分支）
- 场景：AI 交付表单前确认下拉项真的是 `['已通过','待复测','已报废']`

**⑥ Word 页眉/页脚/水印读回（C）**
- 方案：`wps_word_read_document` 增 `includePageLayout`（默认 true）→ 返回 `sections[].headers/footers/watermarkText/differentFirstPage/differentOddEvenPages`
- 宿主实现：`word.js:787-802` 已能写；读 `section.Headers.Item(1).Range.Text`、`doc.Shapes` 里的文字水印
- 场景：AI 写完"内部机密"水印后要能确认，且要能判断是否与已有水印叠加

**⑦ 筛选/排序状态读回（C）**
- 方案：`excel_get_sheet_outline` 增 `filter: { autoFilterMode, filterRange, criteria[] }`，或 `excel_get_sheet_view` 一并返回
- 宿主实现：`excel.js:1665-1671` 已在读 `sheet.AutoFilterMode` / `sheet.AutoFilter.Range`，只是仅回显范围
- 场景：AI 交付前确认"用户看到的仍是全量数据还是被筛掉的子集"——这是数据看板最常见的误读

**⑧ Word 段落排版读回粒度（A）**
- 方案：`wps_word_read_document` 的 `paragraphDetails[]` 增 `styleName`、`lineSpacingRule`、`lineSpacing`、`firstLineIndentChars`、`spaceBefore`、`spaceAfter`
- 宿主实现：`word.js:189-198` 的读取块补齐字段（`para.Style.NameLocal`、`para.Format.*`）
- 场景：AI 按"公文规范"排版后逐段核验"仿宋_GB2312 / 三号 / 28 磅 / 首行缩进 2 字符"是否真的落地

**⑨ Excel 超链接读写（B2）**
- 新增：`excel_set_hyperlink(address, url, textToDisplay?, screenTip?, sheetName?)` + 读回并入 `get_range_styles` 的 `include` 枚举（新增 `"hyperlink"`）
- 现成实现：Microsoft `range.js:217`；WPS 需新写 `range.Hyperlinks.Add`
- 场景：AI 做目录页/跨表跳转导航

**⑩ Excel 命名区域读写（B2）**
- 新增：`excel_list_named_ranges` / `excel_upsert_named_range(name, reference, comment?)` / `excel_delete_named_range(name)`
- 现成实现：Microsoft `workbook.js:95`（list）、`workbook.js:105`（update/delete）；WPS 需新写
- 场景：把 `'明细数据'!$A$1:$K$5422` 命名为 `明细`，让公式与 AI 指令都可读

**⑪ PPT 已有页文本的段落级改写 + 备注写入（A/D）**
- 方案：`wps_ppt_manage_shapes_and_media` 的 `update_shape` 增 `paragraphs: [{ index, text, fontSize, fontColor, bold, alignment }]`；`wps_ppt_manage_slides` 增 `action: "set_notes"` + `notes`
- 宿主实现：文本用 `TextFrame.TextRange.Paragraphs(i)` 定位（`ppt.js:692` 现在整框赋值 `TextRange.Text`）；备注用 `slide.NotesPage.Shapes.Placeholders.Item(2)`（`ppt.js:366-375` 已有写备注的代码，只是只在 `generate_deck` 里调用）
- 场景：AI 微调第 7 页第 2 条要点，现在只能整框重写→丢失原有排版与配色

**⑫ 暴露 `wps_clear_range` / `excel_clear_range`（B1，零成本）**
- 参数：`address`（必填）、`sheetName`、`workbookName`、`applyTo` ∈ `all|contents|formats`（**建议新增**：现在宿主写死 `range.Clear()`，`excel.js:108`）
- 场景：AI 重置一块脏数据区；没有它只能删行（破坏结构）或写空串（留格式）

**⑬ 暴露 `wps_get_style_token` / `excel_get_style_token`（B1，零成本）**
- 参数：`sheetName`、`sampleAddress`（默认 `"A3"`）、`workbookName`
- 场景：AI 在既有报表里新增一块内容时沿用原字体/表头底色，避免"同一份报表两种蓝"

**⑭ Excel 图片/形状插入与读取（B2）**
- 新增：`excel_insert_image(base64Image, left?, top?, width?, height?, sheetName?)` / `excel_list_shapes` / `excel_update_shape(name, left/top/width/height/action:"delete")`
- 现成实现：Microsoft `shape.js:4/19/30`
- 场景：AI 贴公司 Logo、贴流程图示意、清掉压住数据的旧图片

**⑮ Excel 工作表隐藏/显示（B2，双向不通）**
- 扩 schema：`wps_manage_sheet` 的 `action` enum 增 `hide`/`show`（并把 `targetType` 与宿主 action 对齐）
- 宿主实现：WPS 新写 `sheet.Visible = false/true`（`xlsxSheetVisible=-1/xlSheetHidden=0`）；Microsoft 已有（`sheets.js:71-74`）
- 场景：AI 生成"只留汇总页给领导看"的交付版，把 12 张明细页收起来

**⑯ Word"只处理命中关键词"（B1，宿主已实现）**
- 方案：`wps_word_format_document` schema 增 `searchQueries: string[]`（宿主键名，`word.js:379`），并在 description 里说明"命中即加粗/改字号"
- 场景：AI 把全文所有"已完成"标绿、把所有金额加粗——现在只能整段改，粒度太粗

**⑰ Excel 区域复制/移动（B2）**
- 新增：`excel_copy_range(sourceAddress, destinationAddress, copyType ∈ all|values|formats|formulas, sheetName?)`
- 现成实现：Microsoft `range.js:206`
- 场景：AI 复制"本月模板区块"到新表、把汇总块搬到看板页

**⑱ Excel 导出 PDF / 页面设置（B2）**
- 新增：`excel_export_pdf(outputPath, address?, sheetName?)`、`excel_set_page_setup(orientation, paperSize, printArea, fitToPages, headerText, footerText)`
- 宿主实现：WPS `PageSetup` / `ExportAsFixedFormat` 全未出现（§2.3）
- 场景：AI 交付可打印的报表（A3 横向、一页宽、页脚带页码）

**⑲ Excel 行/列分组分级显示（B2）**
- 新增：`excel_group_rows_columns(targetType, index, count, level, action ∈ group|ungroup|collapse|expand)`
- 宿主实现：`Rows.Group` / `OutlineLevel` / `ShowDetail`（WPS 未出现）
- 场景：AI 生成"可折叠的明细+小计"结构，这是财务/经营报表的标准形态

**⑳ 单元格锁定 + 公式隐藏（B2）**
- 新增：`excel_set_cell_protection(address, locked?, formulaHidden?, sheetName?)`，并让 `manage_sheet protect` 支持 `allowEditRanges`
- 宿主实现：`range.Locked` / `range.FormulaHidden`（WPS 未出现）
- 场景：AI 交付模板时只让用户改黄色输入区，公式区锁定并隐藏

### 3.3 C 类集中说明（最影响 AI 自我验证）

这 7 项有共同的失败模式：**工具返回 `success:true` + 回显入参，AI 以为改成了，实际只能靠截图观察像素**。

| 写工具 | 回显内容 | 读不回来的关键信息 |
|---|---|---|
| `wps_add_conditional_formatting`（`excel.js:536`） | `ruleType` + `address` | 规则是否真的挂上、阈值、优先级、`stopIfTrue` |
| `wps_freeze_panes`（`excel.js:611`） | 入参 + 一句文案 | `SplitRow/SplitColumn` 实际值、是否被后续调用覆盖 |
| `wps_set_data_validation`（`excel.js:1770`） | `validationType` + `address` | 下拉项、区间、提示与报错文案 |
| `wps_set_filter_and_sort`（`excel.js:1642`） | `appliedFilterRange`、`sortedRuleCount` | 当前筛选条件、排序是否被用户改回 |
| `wps_manage_sheet` protect/tab_color（`excel.js:1837`） | `protected:true` / 颜色回显 | 是否真被保护、密码是否为无、标签色取值 |
| `wps_word_page_layout_and_watermark`（`word.js:779`） | 入参回显 | 页眉页脚实际文本、水印是否存在 |
| `wps_create_pivot_table`（`excel.js:1554`） | 字段计数 | 透视表实际布局、可刷新性、后续读取 |

补充：**唯一有真正读回校验的是 `format_cells` 的合并分支**（`excel.js:511-526` 写后读 `MergeCells`，不一致就抛错）——
这条已经证明"写后读回"在本仓库是可行且已被采用的模式，其余写工具只是没做。

## 4. 附：参数被实现忽略 / 与实现不符

判定方法：工具 schema（`src/bridge/tools/definitions/*.ts`）里声明、但宿主函数体从未解构或解构后从未使用；
以及 schema 的 `enum`/默认值承诺与宿主分支不一致。**全部为静态读码结论。**

### 4.1 声明了但宿主完全不读（静默无效）

| 工具 | 参数 | 宿主证据 | 后果 |
|---|---|---|---|
| `wps_auto_fit_columns` | `address`（"需要自适应调整的单元格或列区域 'A1:E5' 或 'A:E'"） | `excel.js:944` 只解构 `{ sheetName, columnRules, workbookName }` | 传了地址也不生效，只能全表或按 `columnRules` |
| `wps_capture_sheet_preview` | `chartName` / `name`（"单独捕获截图的原生图表名称"） | `excel.js:1086` 只解构 `{ sheetName, address, range, workbookName }` | WPS 上永远截区域/首图；**Microsoft 侧支持**（`chart.js:247-255`）→ 跨宿主行为不一致 |
| `wps_add_chart` | 顶层 `left` / `top` / `width` / `height` | `excel.js:1113-1129` 只认 `position.leftCell/width/height` | WPS 上像素定位失效；**Microsoft 侧由 `normalizer.ts:164-167` 补齐** → 跨宿主行为不一致 |
| `wps_add_chart` | `cellRange` / `startCell` / `endCell`（"刚性单元格吸附"） | 全文件无 `cellRange`/`startCell`/`endCell`（grep 仅命中 `word.js:722-724` 的合并单元格） | 承诺的"咬死区域内、杜绝像素漂移"完全没实现 |
| `wps_ppt_insert_native_chart` | `hasLegend` / `showDataLabels` | `ppt.js:462` 未解构；`ppt.js:473` 调用时只传 `{chartType,title,categories,series}` | 每个系列仍可用 `series[].hasDataLabels` 兜底，但顶层开关无效（`renderPptChart` 本身支持，`ppt.js:153`） |
| `wps_ppt_manage_table` | `zebra`（"是否启用交替斑马纹底色"） | `ppt.js:856` 解构后全文无引用 | 传了无效 |
| `wps_duplicate_sheet` | `position: "end"`（"工作簿最末尾"） | `excel.js:911` 只有 `if (position === "before")`，无 `end` 分支 | `end` 等同 `after`；源表不在末尾时位置不符预期 |
| `wps_format_cells` | `verticalAlignment: "top" \| "bottom"` | `excel.js:476` 只有 `if (verticalAlignment === "center")` | 垂直靠上/靠下静默无效（schema 还把 center 写成默认） |
| `excel_manage_rows_and_columns`（host=microsoft） | `targetType`、`action=hide/unhide/set_size` | `range.js:149-190` 只认 `dimension`(rows/columns) 与 `insert/delete` | `targetType:"column"` 被忽略 → **想插列却插行**；hide/set_size 静默返回 success |
| `excel_set_data_validation`（host=microsoft） | `validationType`/`listItems`/`operator`/`minVal`/`maxVal` | `range.js:231` 只认 `params.rule` | 先 `clear()` 再什么都不设 → **清掉原有校验还不报错** |

### 4.2 schema 默认值与实现不符

| 工具 | 参数 | schema | 实现 |
|---|---|---|---|
| `wps_word_read_document` | `includeFormatting` | "默认 true" | `word.js:175` 用真值判断 → **不传就不返回** `paragraphDetails` |
| `wps_word_read_document` | `maxParagraphs` | "默认 200" | `word.js:152` 大纲循环硬编码 `Math.min(paraCount, 200)`，不受 `maxParagraphs` 影响 |
| `wps_word_read_document` | `scope` | enum 含 `selection`/`tables` | 正文分支只判 `full`/`paragraphs`/undefined；`selectionText` 与 `tables` **总是**返回，`scope` 实际只影响 `previewText` 是否省略 |

### 4.3 声明的 enum 值在宿主没有分支

| 工具 | 参数值 | 宿主证据 | 后果 |
|---|---|---|---|
| `wps_word_format_document` | `preset: "academic"` | `word.js:474` 只有 `gov_standard` 与 `business_modern` 两个分支，其余落 `applyFormatToPara` | "学术排版"静默退化为自定义（且大部分参数未传 → 几乎不改动） |
| `wps_ppt_manage_table` | `action: "set_table_data"` | `ppt.js` 全文 `0` 处出现 `set_table_data`；switch 只有 create/read/set_cell_text/style | 描述承诺"批量注入二维数据"，实际必然抛 `未知的表格操作` |
| `wps_manage_sheet` | —（反向缺失） | Microsoft 支持 `hide`/`show`（`sheets.js:71-74`），WPS 无 | 见 §3 第 15 条 |
| `wps_modify_rows_columns` | —（反向缺失） | 与 `wps_manage_rows_and_columns` **同一个宿主函数**（`dispatch.js:80` 与 `dispatch.js:116` 都调 `modifyRowsColumns`），但前者 schema 没有 `set_size` | 同名能力两个工具、参数集不一致，AI 容易选错 |

### 4.4 宿主支持但 schema 没暴露的参数（能力被隐藏）

| 工具 | 隐藏的参数 | 宿主证据 | 价值 |
|---|---|---|---|
| `wps_word_format_document` | `searchQueries: string[]`（按关键词命中处加粗/改字号，穿透正文与全部表格） | `word.js:379-421`，逻辑完整 | 见 §3 第 16 条 |
| `wps_word_manage_table` | `columnIndex`（`format_cell` 可用它代替 `cellColumn`） | `word.js:562` 解构了 `columnIndex`，`word.js:684` 使用 | 低（`cellColumn` 已可用，属命名冗余） |
| `excel_format_cells`（host=microsoft） | `style`（命名样式模板）、`autofitRows`、`across`、嵌套 `font{}/fill{}/alignment{}` | `format.js:12-31`（`OPPO_Banner` 等 9 个模板）、`format.js:104/119` | 低，且**模板是客户定制硬编码**，不建议暴露为通用参数 |
| `excel_get_range_styles`（host=microsoft） | 返回结构是 `{font,fill,alignment,numberFormat}`，与 WPS 的 `{styles, mixedOrUnavailableFields}` 不同 | `range.js:36-60` vs `excel.js:249-286` | 中：同名工具两宿主返回形状不同，AI 需要写两份解析；建议在 `normalizer.ts` 的响应侧统一（`normalizer.ts:368` 目前没有 `get_range_styles` 分支） |

## 5. 附：方法与局限

### 5.1 本报告的取证方式

- **工具面**：以 `docs/acceptance/2.1.0-p0p1/tools-snapshot.p5.json`（实测 91 项）与
  `src/bridge/tools/definitions/**` + `tools/{diagnostics,excel,audit}.ts` 交叉核对，非人工清点。
- **宿主实现**：逐个读 `wps-addon/src/{excel,word,ppt,ppt-layout,dispatch,shared}.js`、
  `office-addon/src/**`、`src/bridge/{gateway.ts,gateway/*.ts,office/*.ts,errors.ts}`，
  引用一律给 `文件:行号`。
- **能力空集**：用"宿主文件中是否出现该 API 成员"的反向 grep 作为证据（§2.3），
  这比"我没看到"更强，但仍弱于实机调用。

### 5.2 局限（请连带阅读）

1. **全部结论为静态读码，未做真机验收。** 本任务授权"以读代码为主，必要时只读反射探测"，
   因此本报告不声称任何工具"实测可用/不可用"；涉及"必然失败/静默无效"的判断都标注了代码依据，
   属**静态可判定**，不等于运行验证。
2. **行号对应 2026-09-22 16:12 的工作区快照。** 同一目录下的并行任务正在修改
   `wps-addon/src/{excel,word,ppt}.js`（会话期间 `excel.js` 已从 1857 行增至 1879 行，
   `createPivotTable` 由 1532 行移到 1554 行）。引用时请以**函数名/代码片段**为准，
   行号仅作定位起点。
3. **B2 类的"用不上"判断带主观性。** §3 的排序按"AI 操控文档的价值"评，标准是
   使用频率 × 无替代路径 × 自我验证需求；不同业务（公文 / 财务 / 看板 / 汇报）会给出不同顺序。
4. **逃生舱口径。** 所有 B2 类能力理论上都能用 `wps_execute_script` 触达（§1.6 的说明），
   本报告的"缺口"指的是**没有专用工具、AI 必须写原生 JS、无参数校验、无读回、无审计**，
   而不是"完全做不到"。
5. **未覆盖范围**：`resources/office/**`（Windows 原生驱动 PowerShell）的能力面、
   `skills/` 里的技能封装、以及 `office_addon` 之外的非 Excel Microsoft 组件。
   `office_execute_script` 的 JXA/PowerShell 驱动是另一条独立能力通道，未纳入本次比对。

### 5.3 与同目录其他报告的边界

同目录另有 `06-wps-api-map.md`（宿主对象模型能力地图）与 `08-ms-vs-wps.md`（双平台对照）。
本报告（`07-tool-gap`）只回答一个问题：**"宿主能做、但 MCP 没给 AI 工具"的有哪些、按价值怎么排、
补的时候加什么参数、改哪个文件**。宿主 API 面的完整测绘与双平台差异矩阵不在本报告范围内。

### 5.4 最该优先补的 3 件事（结论）

1. **给已写入的格式类操作补读回**（§3.3 的 7 项），因为这是"AI 以为自己改对了"的唯一系统性漏洞；
   成本最低的入口是先把 `freeze_panes` / `data_validation` / `conditional_formatting` 三项的
   读取并入 `get_sheet_outline` 与 `get_range_styles` 的返回结构。
2. **清掉 B1 死分支**（`wps_clear_range`、`wps_get_style_token`、`wps_word_capture_preview`），
   宿主与网关都写好了，只差 schema —— 单位收益最高，且能顺带缩减 `contract-consistency` 台账。
3. **修 Word 的"定位"能力**（批注只能挂当前选区、格式只能整段、查找不返回位置、关键词加粗被隐藏），
   因为 Word 侧 AI 目前只能"整篇/整段"操作，是三个组件里粒度最粗的一个。
