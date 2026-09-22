# 09 · MCP 工具说明与契约修正（P5 批次）

> 任务：把 `issues.md` 中已证实的「工具说明 / 参数 schema」问题在**工具定义层**修掉。
> 写区：`src/bridge/tools/definitions/**`。本文件按硬约束「边改边更新」维护，末节为最终结果。
> 状态口径：`已改` / `已改（说明层）` / `写区外，已由协调方修复` / `不修（附理由）`。

## 1. 逐条结果（编号 → 文件 → 改了什么 → 状态）

| 编号 | 文件 | 改动 | 状态 |
|---|---|---|---|
| ISS-27 | `definitions/excel.ts` | 27 个 `wps_*` 表格工具全部写入选型规则：`wps_get_workspace_summary` 写完整规则（两个入口等价、不要重复调用、Microsoft 只能用 `excel_*`），其余 26 个用统一短标记「同名 wps_* 与 excel_* 二选一——wps_* 只走 WPS 表格（不传 host），excel_* 跨宿主（必传 host）」。该措辞在 `excel_*` 派生说明里同样读得通（`excel_*` 说明由 `wps_*` 文本派生）。另写清真重复对：`wps_modify_rows_columns`（无 `set_size`）与 `wps_manage_rows_and_columns`（多 `set_size`/`size`）互为区别。 | 已改 |
| ISS-28 | `definitions/audit.ts` | `wps_rollback` 说明写明覆盖边界：只回滚 `patch_cells` 的值/公式；样式、条件格式、图表、行列/工作表结构、Word/PPT、脚本执行均无快照、不可撤销；回滚前比对当前内容，有后续修改即拒绝。 | 已改 |
| ISS-30 | `definitions/audit.ts` + `src/bridge/tools/audit.ts` | `wps_rollback` 说明与 `auditId` 描述本批补齐（含 auditId 来源）。其余 3 个审计工具（10 个无说明参数）定义在我写区外，**已由协调方在 `59dd433` 修复**；终检「无说明参数 = 0」。 | 已改 / 写区外已修 |
| ISS-20 | `definitions/excel.ts` | `wps_update_chart` 说明首句即写宿主限制：WPS 未实现、调用在执行前被拒、不改文档；给出替代路径（`wps_execute_script` 或 `host=microsoft`），与 `HOST_IMPLEMENTATION_GAPS.wps=['update_chart']` 口径一致。并写清定位必传其一（`chartName`/`name`/`shapeName`）。 | 已改 |
| ISS-33 | `src/bridge/tools/excel.ts`（写区外） | 我无法改到；**已由协调方在 `59dd433` 修复**：`excel_*` 后缀收缩为「统一表格入口；host 必填。」，验收状态改挂到必填的 `host` 参数描述（"Microsoft 通道在 macOS 与 Windows 均尚未实机验收"）。终检确认 `excel_add_chart` 的 `host` 描述含该口径。 | 写区外已修 |
| ISS-31 | `definitions/excel.ts` / `word.ts` / `ppt.ts` | 9 处参数补 `type`（终检「缺 type = 0」）：`manage_rows_and_columns.index`→`["number","string"]`、`find_and_replace.searchQuery`→`["string","number"]`（各含 `excel_*` 派生副本）、`word_write_content.content`→`["string","array"]`+`items`、PPT `shapeId`/`shapeId1`/`shapeId2`→`["string","integer"]`。`borders` 的联合类型**保留** `["string","boolean"]` 并写清取值：JSON Schema draft-07 明确允许 `type` 为字符串数组，且 `catalog.validateArgs`（第 71–72 行）原生支持；改成单类型会拒绝现有合法调用 `borders: true`。17 项接受集回归全部符合预期（见 §3）。 | 已改 |
| ISS-32 | `definitions/microsoft.ts` | 残句根因是装配层正则删掉「实现 100% 任意操作无死角！」后留下悬空逗号；正则本身在写区外，故改写定义文本使清洗后为完整句。终检 `office_execute_script` 交付文本以「该通道在 macOS 与 Windows 均尚未实机验收。」收尾；全量 91 条无「以标点结尾 / 残留【】/ 连续空白」。 | 已改 |
| ISS-34 | `definitions/word.ts` / `ppt.ts` / `excel.ts` | 枚举与说明逐项对齐：`preset` 补 `academic`（并如实写明当前等同 `custom`）、`read_document.scope` 补 `paragraphs`/`tables`、`write_content.type` 补 `code_block`、`stylePreset` 补 `clean_minimal`（如实写明等同 `none`）与 `none`、PPT `insert_native_chart.chartType` 补 `bar_stacked`、`manage_shapes_and_media.action` 说明补 `delete_shape`、`verticalAlignment`/`operator`/`deck[].chart.chartType` 补全枚举释义、组件类枚举统一写 `'word'\|'excel'\|'ppt'`。终检「枚举未在说明中出现」仅剩自解释的 `component`。 | 已改 |
| ISS-35 | `definitions/*.ts` | 见 §2 决策：**未加任何 `required`**（快照 required 变化 = 0），17 个零必填工具的必需输入全是「多选一」，改为在说明首行写明必传集合；`clearAll` 写明是显式确认开关。 | 已改（说明层） |
| ISS-36 | `definitions/*.ts` | 营销句清零（`彻底/痛点/100% 完美/杜绝像素/极大提升/图灵/超级引擎/无死角/专业级/智能审核` 全仓 0 命中），`cellRange` 的「100% 完美行级对齐、杜绝像素漂移」等与实现不符的承诺改为实情；字数改花在约束（选型、必传项、宿主忽略参数、覆盖边界、读回要求）。`excel_*` 的 864 字样板由协调方在 `59dd433` 收缩。 | 已改 |
| ISS-37 | `definitions/excel.ts` / `audit.ts` | 报错文案在 `catalog.validateArgs`（写区外），按"以说明补充参数集"处理：`wps_rollback` 写明"只接受 auditId，传 host/sheetName 报未知参数"；`wps_save_workbook` 写明"参数集只有 workbookName，不接受 sheetName"；`create_sheet`/`delete_sheet` 写明参数；枚举参数说明内均列允许值。 | 已改（说明层） |
| ISS-25 | `definitions/excel.ts` | `wps_capture_sheet_preview` 写明优先级（`address` 优先于别名 `range`，都不传取 UsedRange；`chartName`/`name` 在 WPS 宿主被忽略）与返回结构 `{success, workbookName, sheetName, address, imageBase64, imageMimeType, imageSizeBytes}`，并提示失败时可能没有 `imageBase64`。 | 已改 |
| ISS-82 | `definitions/ppt.ts` | `layoutIndex` 写明是 **ppLayout 枚举**而非 `CustomLayouts` 序号，给出实测常用值（1=标题幻灯片、2=标题和文本、7=标题和图示或组织结构图、12=空白默认），并写明越界不报错、要精确套版式需用脚本操作 `slide.CustomLayout`。 | 已改 |
| ISS-83 | `definitions/ppt.ts` | 从 `wps_ppt_manage_table.action` 枚举移除宿主没有的 `set_table_data`，说明改写为"宿主没有批量改写已有表格数据的操作：重新 `create_table(data)` 或逐格 `set_cell_text`"。 | 已改 |
| ISS-84 | `definitions/ppt.ts` | `data` 参数写明「每个单元格必须是字符串，数字会被 schema 拒绝并报 `arguments.data[行][列]: 类型不正确`，请自行转字符串」；并注明只用于 `create_table` 初始填充。 | 已改 |
| ISS-88 | `definitions/ppt.ts` | `swap_shapes` 写明只交换 Top、Left 不变（横排形状交换后位置不变，水平换位请用 `update_shape`）；`align_shapes` 写明基准是 `shapeIds` 首个可解析形状、不是页面对齐；`shapeId`/`shapeId1`/`shapeId2`/`shapeIds` 写明解析顺序：数字先按 `Shape.Id`、未命中再按页内 1-based 索引兜底，字符串按名称。 | 已改 |

## 2. 两项决策（供复核）

1. **ISS-35 不加 `required`**：17 个零必填写工具的"真正必需输入"全部是**多选一**集合，单项 `required` 会破坏合法用法：
   - `add_chart`：`dataRange` ∪ `dataRanges`（宿主缺数据源直接报错，二者任一合法）；
   - `delete_chart`：目标选择器 ∪ `clearAll`；
   - `freeze_panes`：三个参数至少一个；`page_layout_and_watermark`：页眉/页脚/水印至少一个；`word_format_document`：预设或自定义参数至少一个；
   - `update_chart`：`chartName`/`name`/`shapeName` 三个别名任一；
   - `unlock_target`×2：`component` 省略即"解锁全部"是已文档化行为；`save_workbook`/`word_save_document`/`word_close_document`/`word_create_document`/`insert_table_of_contents`/`auto_fit_columns`/`capture_sheet_preview`/`ppt_capture_slide_preview`：省略即用已文档化的默认目标/默认页。
   因此按"说明层写明必传其一"处理。若要在 schema 层强制"至少一个"，需要 `anyOf`（`catalog.validateArgs` 目前不识别，等于只对客户端生效）或 `oneOf`（会拒绝"两个都传"这类现在合法的调用）——属契约语义变更，建议由协调方决定。
2. **`borders` 的联合类型保留数组写法**：JSON Schema draft-07 允许 `type` 为字符串数组，本仓 `validateArgs` 也按数组处理；改成单类型会拒绝 `borders: true`。已用 17 项接受集回归证明未收紧。（若坚持单一 `type`，只能改成 `oneOf`，但那样该参数又会回到"缺 `type`"。）

## 3. 验证证据

| 项 | 命令 | 结果 |
|---|---|---|
| 类型检查 | `npm run typecheck` | **0 错误**（exit 0） |
| 全量测试 | `npm test` | **86/86 pass**，exit 0 |
| 缺 `type` 参数 | 自检脚本遍历 91 个工具的 schema | **0**（改前 9 处） |
| 无说明参数 | 同上（含嵌套对象/数组元素） | **0**（改前 13 处，含协调方修复的审计族） |
| 说明完整性 | 检查"以标点结尾/未以句末标点结束/残留【】/连续空白/首尾空白" | 91 条全部通过（`office_execute_script` 残句已消除） |
| 枚举对齐 | 枚举成员必须出现在参数说明或工具说明中 | 仅剩自解释的 `component`（word/excel/ppt） |
| 接受集回归 | 17 项代表性入参（`shapeId` 数字/字符串、`shapeIds` 混合、`index` 数字/字母、`searchQuery` 数字/字符串、`content` 字符串/数组/非法对象、`borders` 字符串/布尔/数字、`data` 字符串/数字、`set_table_data` 移除后拒绝） | **17/17 符合预期**，无合法调用被收紧 |

## 4. 契约快照差异（`tools-snapshot.p5.json` → `09-tools-snapshot.after.json`）

- 工具数 **91 → 91**；只读/写 **21/70 → 21/70**；**顺序逐项一致**；**无新增、无删除**。
- `capabilities()` 与 `excelMethods` **无变化**。
- **description 变化 82 条**（其余 9 条未动）。
- **inputSchema 变化 59 条**，全部来自参数 `description` 文本；结构性变化只有下面两类：
  - **`type` 补全 9 处**（`index`×2、`searchQuery`×2、`content`、`shapeId`×3，成对含 `excel_*` 派生副本）；
  - **`enum` 收窄 1 处**：`wps_ppt_manage_table.action` 去掉 `set_table_data`。
- **`required` 变化：0 处**（§2 决策 1）。
- 快照文件：新快照另存为 [09-tools-snapshot.after.json](09-tools-snapshot.after.json)，**未覆盖** `tools-snapshot.p5.json`（保留改前基线以便复核）；复核通过后可提升为新基线。
- 说明总量：**6549 → 15213 字符**（均值 72 → 167）。增长几乎全部来自约束类信息（选型标记 26×62 字、必传集合、宿主忽略参数、覆盖边界、读回要求），营销句已清零；`excel_*` 的 864 字样板由协调方收缩为 27×14 字。

## 5. 附带修正（不在 18 条清单内，属同一段文字里的失实承诺，已一并改）

均为「说明与实现不符」且改动只在文本层，逐条有源码或台账证据：

| 位置 | 原文/问题 | 现说明 |
|---|---|---|
| `wps_search_cells` | 声称可搜公式，实际只读 `Value2`（ISS-29） | 写明只按显示值搜索，按公式搜会 0 命中仍返回 success；要找公式用 `wps_read_range` |
| `wps_add_chart` | `cellRange` 承诺"100% 完美行级对齐、杜绝像素漂移"（ISS-18） | 写明 WPS 只认 `position.leftCell`/`position.width`/`height`，顶层定位参数与 `cellRange`/`startCell`/`endCell` 被忽略（源码核对：`excel.js` `addChart` 只读 `position`） |
| `wps_delete_chart` | `leftCell` 参数名像精确锚点（ISS-19） | 写明是 ±30px 邻近匹配、命中多张即拒绝、一张未命中报错 |
| `wps_patch_cells` | 未写值/公式互斥与清空方式（ISS-01/21/45） | 写明互斥、空项语义、null 清空、字符串日期转序列号 |
| `wps_format_cells` | `borders:'none'` 说可去边框，实现 falsy 跳过（ISS-42 后半） | 写明当前不会去边框 |
| `wps_manage_cell_comments` | `author` 看似写作者，实际拼进正文（ISS-43/61） | 写明作者恒为当前 WPS 用户，`author` 只是签名文本 |
| `wps_find_and_replace` | `row/col` 是相对偏移未写（ISS-65） | 写明相对 `searchRange` 的偏移 |
| `wps_save_workbook` | 未写保存是工作簿级（ISS-16） | 写明会连同其他会话的在途改动一起落盘 |
| `wps_get_workspace_summary` | `hasOpenWorkbook` 语义易被误读（ISS-02/77） | 写明 false 只表示未解析到目标，判断已打开文件请读 `openWorkbooks` |
| `wps_auto_fit_columns` | `address` 被实现忽略（ISS-39 剩余部分） | 标注为兼容参数、当前被忽略 |
| `wps_create_pivot_table` | `destSheetName` 必须已存在未写（ISS-63） | 写明必须先建表，否则报错；空表需手动刷新 |
| `wps_word_page_layout_and_watermark` | 只作用第 1 节、水印只落第 1 页（ISS-58/72） | 写明边界与跨页水印的脚本做法 |
| `wps_word_write_content` | `bookmark` 定位实测不落位（ISS-70） | 写明须读回核对，勿只凭 success |
| `wps_ppt_insert_native_chart` | 宿主插入可能不建形状仍返回成功（ISS-80） | 写明必须读回/预览核实 |
| `wps_inspect_api` | 深挖原生 getter 会令 WPS 崩溃（ISS-89） | 写明一次只探一个表达式、列出已知危险 getter、反射不等于无副作用 |
| `wps_ppt_manage_slides`（add） | `add`/`duplicate` 不幂等未写（ISS-78） | 写明追加型操作不幂等 |
| `excel_*` 选型标记 | —— | 措辞刻意做成在 `wps_*` 与派生的 `excel_*` 文本中都读得通 |

## 6. 未完成 / 需协调方处理

1. **ISS-35 的 schema 级强制**：需 `anyOf` 支持或契约语义决策（见 §2），本批只在说明层解决。
2. **ISS-37 的报错文案本身**：`unknown argument`/`不在允许值中` 的提示由 `src/bridge/catalog.ts` 生成，不在写区；如需附允许值清单，需改该文件。
3. 本轮**只做静态验证**（typecheck / 测试 / 快照 / 自检脚本），**未连真实宿主**：说明文本描述的行为来自源码与既有台账证据，未在 WPS/Microsoft 实机上重新复跑。
