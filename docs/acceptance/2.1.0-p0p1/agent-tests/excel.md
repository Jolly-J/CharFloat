# WPS 表格（Excel）侧工具全面测试报告

- 候选标识：`2.1.0-p0p1`
- 执行者：teammate `excel-tester`（共享任务 `task-1`）
- 桥接服务：`http://127.0.0.1:19890/api/v1/tool/call`，WPS 加载项 2.1.0（macOS / WPS 12.1.28496）
- 测试对象：工具清单快照 `tools-snapshot.p17.json` 中的 **101 个表格侧工具**（50 个 `excel_*` + 51 个 `wps_*`，排除 word/ppt 与锁/审计中的无关项）

---

## 0. 结论摘要

**判定标准（本报告唯一的通过口径）**：写进去的值，能从**宿主直接读回来一致**才叫通过。
工具返回 `success: true` 一律不作为通过依据；本报告所有结论都附带 `wps_execute_script` 的宿主原始读数。

| 分层 | 数量 | 说明 |
|---|---|---|
| 已真机验证的断言 | 约 200 条 | 每条都有宿主读数 |
| 确认真实缺陷 | **21 条**（高 4 / 中 10 / 低 7） | 见 §4 |
| 确认正常（含易误判为缺陷者） | 见 §5 | 含 4 条"看起来像缺陷、实测正常"的澄清 |
| 测试脚本自身错误 | 19 类 | 见 §7，**均未计入缺陷** |
| 因事故 / 主动排除而未覆盖 | 见 §6 | 明确列明原因 |

### 最该先修的 4 条（高）

1. **`ruleType:'text_contains'` 的条件格式完全不生效**——公式相对引用被平移（`M1` 变 `Y1`），且同一调用会随"活动单元格"不同而语义不同。宿主 `DisplayFormat` 证明目标单元格没有被高亮。这是"调用成功但没生效"的教科书案例。
2. **`format_cells` 的 `verticalAlignment` 被忽略**——`top`/`bottom` 都落到 `-4108`(center)，而返回体 `appliedStyles.verticalAlignment` 照样回显请求值。
3. **8 个形状类工具的 schema 缺 `workbookName`**——参数被桥接层丢弃，宿主回退到**陈旧锁目标**，导致形状工具整体不可用（`未在 WPS 中找到目标工作簿 [工作簿1]`）。属"工具没暴露目标参数"，不是宿主不支持。
4. **`create_workbook` 的 `savePath` 指向已存在文件时谎报成功**——返回 `success:true` + `savedPath` + "并保存到 …"，但工作簿仍 `unsaved`、磁盘 mtime 未变。

---

## 1. 判定方法与证据标准

每一条断言固定三步，缺一不可：

1. 写入一个**与默认值不同的特征值**；
2. 用 `wps_execute_script`（`component:'excel'`，**显式带 `workbookName`**）**直接问宿主**读回真实状态；
3. 比对宿主读数与期望值。

**颜色常量一律用函数换算**（`bgr(hex) = r + g*256 + b*65536`），不手写 BGR 整数——手写必错，会造出假失败。
**删除类结论一律用集合枚举核对**（`Count` + 逐项 `Name`），不用单对象探针——单对象探针在本机给出过假阳性（见 §7.16）。

隔离方式：`wps_create_workbook` 新建 `AgentExcelProbe.xlsx` 作为唯一探测簿，全部用例在其临时表（`CF/CV/CC/CD/CP/CX/DO/D6/D7/D8`）上进行；测试结束关闭探测簿并删除文件。**全程未修改 `测试表格.xlsx` 的任何既有业务数据**（仅只读读取表名与已用区域）。

---

## 2. 环境事故记录（不计入产品缺陷）

### 2.1 WPS 主进程崩溃

- 时间：20:49:03，**SIGSEGV**，`wpsapi` 递归爆栈
- 触发者：teammate `word-ppt-tester` 执行 `wps_word_page_layout_and_watermark` 的 `watermarkText`（Word 侧，**与 Excel 操作无关**，我仅被连带）
- 表现：`bridge_diagnose` 显示 word/excel/ppt 三组件在**同一毫秒**（`1790081341109`）以 `code 1001` `clean-close` 断开；`ps aux` 中已无 `wpsoffice` 主进程
- 恢复：word-ppt-tester 重开 WPS；我随后重开探测簿续跑
- 后果：我已跑完的**批次 A（环境/读工具）与批次 B（表管理/行列）证据有效**；批次 C 首次执行因断连作废，恢复后重跑；探测簿未保存的内存态（批次 B 建的表）在崩溃中丢失，属预期

### 2.2 用户数据完整性核对（崩溃后，只读，未做任何修改）——原文保留

WPS 已恢复，excel 组件 `connected=true`。`测试表格.xlsx` 已打开，宿主直读结果：

- **工作表数量：22 张，与崩溃前一致**，表名逐个核对全部在位：
  1 行业概览 / 2 市场预测 / 3 企业现状总览 / 4 产业链与市场空间 / 5 政策与标准 / 6 数据来源 / 7 `__MCP验收测试__` / 8 `sweep图形_钙钛矿产业信息图` / 9 `sweep图表_产能与效率看板` / 10 `sweep表格_钙钛矿企业跟踪看板` / 11 `fault_case1_rollback` / 12 `fault_case2_scope` / 13 `fault_case2_tmp` / 14 `fault_case3_A` / 15 `fault_case3_B` / 16 `fault_case5_params` / 17 `adv_高级能力实测` / 18 `e2e_钙钛矿数据集` / 19 `adv_透视输出` / 20 `zz_总测` / 21 `矢量绘图演示` / 22 `运营驾驶舱`
- 已用区域也在（抽样：行业概览 `$A$1:$C$18`、企业现状总览 `$A$1:$K$39`、`adv_高级能力实测` `$A$1:$J$629`、`adv_透视输出` `$B$4:$N$507`、`sweep图表_产能与效率看板` `$A$1:$AC$43`），没有出现整表被清空的迹象。
- **未出现「文档恢复」对话框**——期间所有宿主调用都正常返回，没有被模态框阻塞。

一个需要判断的观察（当时未采取任何动作）：宿主读到 `w.Saved = false`，即该工作簿在内存中处于"未保存"状态。可能有两种解释，无法从宿主区分：(a) WPS 崩溃恢复后把工作簿标记为脏；(b) 打开后 WPS 惯常标记为脏。**未执行保存，也未选择任何恢复版本。**

> 处置结论（Lead 裁决）：磁盘 mtime = 18:51（崩溃前）、崩溃时间 20:49、内存 `Saved=false`，**无法判定内存版本相对磁盘是新还是旧**，因此**不保存** `测试表格.xlsx`。本报告提交时该工作簿已由他人关闭且**未保存**，磁盘文件保持 18:51 版本未被覆盖。

### 2.3 崩溃遗留的陈旧锁目标（放大了缺陷 #3）

崩溃前宿主活动簿是空白簿 `工作簿1`。崩溃后该名字被写入加载项侧 `lockedTargets.excel` 并**跨崩溃保留**：

- `wps-addon/src/shared.js:259`：`const targetName = workbookName || lockedTargets.excel;`
- 崩溃后所有**未在 schema 暴露 `workbookName`** 的工具都会解析到这个已不存在的目标 → 直接报错
- `wps_unlock_target_document` 返回 `success:true`「已成功解除 EXCEL 文档锁定」，但**复查后有效目标仍是 `工作簿1`**（见缺陷 #12）
- 我用 `wps_lock_target_document {component:'excel', targetName:'AgentExcelProbe.xlsx'}` 临时锁到正确目标，才得以完成形状类用例；测试结束已解锁

---

## 3. 覆盖清单

### 3.1 具备宿主读数证据的工具（71 个）

| 组 | 工具 | 覆盖的参数（有宿主读数的） |
|---|---|---|
| 环境/读 | `bridge_get_capabilities`、`bridge_diagnose` | — |
| | `excel_get_workspace_summary` / `wps_...` | `workbookName`、缺 `host` 报错、非法 `host`、不存在的簿 |
| | `excel_get_sheet_outline` | `sheetName`、`workbookName`、不存在对象、`sheetState` |
| | `excel_read_range` | `address`、`includeFormulas`、`includeNumberFormats`、空区域、非法地址、缺必填 |
| | `excel_search_cells` | `query`（值 / 公式文本）、`maxResults`、无命中 |
| | `excel_get_range_styles` | `mode(summary/cells)`、`maxCells`、`include`（合法/非法/`validation`） |
| 表管理 | `excel_create_sheet` | `sheetName`、重名、非法名、超长名、不存在簿、激活语义 |
| | `excel_delete_sheet` | 正常删除、不存在、不存在簿 |
| | `excel_duplicate_sheet` | `sourceSheetName`、`newSheetName`、`position(after/before/end)`、克隆保真（值/粗体/填充/数字格式/条件格式） |
| | `excel_manage_sheet` | `action(rename/move/tab_color/protect/unprotect/read)`、`newName`、`targetIndex`、`color`、`password`、缺参、非法 action、重名 |
| 行列 | `excel_modify_rows_columns` | `targetType(row/column)`、`action(insert/delete/hide/unhide)`、`index`、`count`、不支持 `set_size` |
| | `excel_manage_rows_and_columns` | 同上 + `action:set_size` + `size`、`index` 用字母、越界、非法 `targetType` |
| | `excel_auto_fit_columns` | `columnRules(colIndex/minWidth/maxWidth)`、无 `columnRules` 走已用区域、`address` 被忽略（已在描述声明）、不存在表 |
| 写入 | `excel_patch_cells` / `wps_...` | `values`、`formulas`、找不到目标簿 |
| | `excel_read_range`（见上） | — |
| 格式 | `excel_format_cells` | `backgroundColor`、`fontColor`、`bold`、`fontSize`、`horizontalAlignment(left/center/right)`、`verticalAlignment(top/center/bottom)`、`numberFormat`、`wrapText`、`rowHeight`、`borders`、`merge`、`unmerge`、非法字号、非法颜色 |
| | `excel_format_text_segment` | `find`、`start`、`length`、`occurrence`、`bold`、`italic`、`underline`、`fontName`、`fontSize`、`fontColor`、未命中/越界 |
| | `excel_get_style_token` | `sampleAddress`、不存在表 |
| 视图 | `excel_set_sheet_view` | `zoom`、`showGridlines`、`showHeadings`、越界 zoom |
| | `excel_freeze_panes` | `freezeRowIndex`、`freezeColumnIndex`、`action:read`、`unfreeze`、0 值边界 |
| 条件格式 | `excel_add_conditional_formatting` | `ruleType(cell_value/data_bar/color_scale/icon_set/top10/duplicate_values/text_contains/formula/clear)`、`operator(greater_than/less_than/between)`、`formula1`、`formula2`、`backgroundColor`、`fontColor`、`barColor`、`colorScaleMin/Max`、`iconSet`、`topRank/topBottom/topPercent`、`containsText`、`clearExisting`、`action:read`、非法 ruleType |
| 有效性 | `excel_set_data_validation` | `validationType(list/number_range)`、`listItems`、`operator(between/greater_than)`、`minVal`、`maxVal`、`promptTitle`、`promptMessage`、`errorTitle`、`errorMessage`、`action:read`、缺 `listItems`、非法类型 |
| 筛选排序 | `excel_set_filter_and_sort` | `range`、`enableAutoFilter`、`sortRules(colIndex/order asc+desc)`、`action:read`、越界 colIndex |
| 图表 | `excel_add_chart` | `dataRange`、`chartType(column_clustered)`、`title`、`hasLegend`、`hasDataLabels`、`seriesColors`、`position(leftCell/width/height)` |
| | `excel_get_charts` | `chartIndex`、`detail` |
| | `excel_update_chart` | `title`、`legendPosition`、`left/top/width/height`、非法枚举值 |
| | `excel_delete_chart` | `chartIndex`、越界 |
| | `excel_export_chart_image` | `outputPath`、`format(PNG)`（已核对落盘字节数） |
| 透视表 | `excel_create_pivot_table` | `sourceRange`、`destCell`、`rowFields`、`columnFields`、`dataFields(sum)`、`action:read`、不存在字段 |
| 形状 | `excel_add_shape` | `kind(geometric/line/textBox/wordart)`、`shapeType`、`left/top/width/height`、`fillColor`、`lineColor`、`lineWeight`、`text`、`fontSize`、`bold`、`italic`、`fontName`、`fontColor`、`textAlign(left/center/right)`、`textVAlign(top/middle/bottom)`、`x1/y1/x2/y2`、未知 shapeType、非法 kind |
| | `excel_update_shape` | `left/top/width/height`、`fillColor`、`text`、`rotation`、`newName`、`visible`、`action:delete` |
| | `excel_list_shapes` | `filterName`、`detail` |
| | `excel_group_shapes` | `shapeNames`、`groupName`（成员数） |
| | `excel_ungroup_shapes` | `shapeName`（释放数） |
| | `excel_set_shape_zorder` | `zOrder(bringToFront/sendToBack)`、不存在形状 |
| 图片 | `excel_manage_pictures` | `action(insert/list/delete)`、`filePath`、`left/top/width/height`、`pictureName`、不存在路径 |
| 超链接 | `excel_manage_hyperlink` | `action(add/list/delete)`、`address`、`url`、`displayText`、`tooltip`、`targetAddress` |
| 命名区域 | `excel_manage_named_range` | `action(add/list/delete)`、`name`、`refersTo`、`comment` |
| 文档属性 | `excel_manage_document_properties` | `action(apply/read/delete)`、`properties(Title/Author/自定义)`、`propertyNames` |
| 结构化表格 | `excel_manage_table` | `action(apply/list/delete)`、`address`、`tableName`、`hasHeaders`、`styleName`、`totalsRow` |
| 打印导出 | `excel_configure_print_layout` | `printArea`、`orientation`、`paperSize`、`centerHorizontally/Vertically`、`printGridlines`、`topMargin`、`leftMargin`、`centerHeader`、`rightFooter`、`zoom`、`fitToPagesWide/Tall`、`action:read` |
| | `excel_export_sheet_pdf` | `scope:sheet`（已核对返回 outputPath） |
| | `excel_capture_sheet_preview` | `address`（返回 imageBase64） |
| 工作簿 | `excel_create_workbook` / `wps_...` | `savePath`（新增 / 已存在两种）、`sheetName` |
| | `excel_save_workbook` / `wps_...` | `workbookName`、多余参数 `sheetName`、不存在簿 |
| 网关 | `wps_execute_script` | `component`、`code`、`params`、`workbookName`（本报告全部宿主读数的通道） |
| | `wps_inspect_api` | `component`、`expression`、`evaluate`、不存在成员 |
| 锁 | `wps_lock_target_document` | `component`、`targetName`（存在/不存在） |
| | `wps_unlock_target_document` | `component` |
| | `wps_get_locked_status` | — |
| 审计 | `wps_get_audit_history` | `limit` |
| | `wps_get_audit_record` | 不存在的 auditId |
| | `wps_rollback` | `auditId`（有/无中间脚本调用两种） |

### 3.2 已覆盖但**未逐参数穷尽**的工具（10 个）

`excel_add_chart`、`excel_update_chart`、`excel_delete_chart`、`excel_get_charts`、`excel_configure_print_layout`、`excel_add_conditional_formatting`、`excel_manage_table`、`excel_manage_document_properties`、`excel_manage_hyperlink`、`wps_execute_script`。
未覆盖的参数逐条列在 §6.2。

### 3.3 明确未覆盖（2 个，主动排除）

| 工具 | 排除原因 |
|---|---|
| `wps_reload_addon` | 会断开**所有队友**（word/ppt/ms）的加载项连接，属跨任务破坏性操作，按 Lead 新规"文档级/全局操作先报备"**主动排除** |
| `wps_clear_audit_history` | 会销毁**团队共享**的审计证据（其他队友的写入留痕），**主动排除** |

---

## 4. 缺陷清单（21 条）

每条给出：最简复现调用 → 宿主读数 → 期望值 → 根因 → 分类。

### 高

#### H-1 `ruleType:'text_contains'` 条件格式完全不生效（公式相对引用被平移）

**分类**：调用成功但没生效（宿主语义边界 + 实现未处理"相对引用基准是活动单元格"）
**最简复现**：

```
1) wps_execute_script:  wb.Worksheets.Item('D8').Range('A1').Select();   // 活动单元格 = A1
2) wps_execute_script:  wb.Worksheets.Item('D8').Range('M3').Value2 = '含特价';
3) excel_add_conditional_formatting { host:'wps', sheetName:'D8', address:'M1:M5',
       ruleType:'text_contains', containsText:'特价', backgroundColor:'#FEE2E2' }   → success: true
```
**宿主读数**：

```
FormatConditions.Item(1).Formula1 = =ISNUMBER(SEARCH("特价",Y1))     ← 期望引用 M1
Range('M1..M5').DisplayFormat.Interior.Color = 16777215 ×5           ← 无任何高亮
```
**期望值**：`Formula1` 引用 `M1`；`M3`（唯一含"特价"的单元格）`DisplayFormat` = `14869246`（= `bgr('#FEE2E2')`）。
**对照实验（证明是活动单元格依赖）**：把活动单元格设为 `M1` 后**同一个调用**：
```
Formula1 = =ISNUMBER(SEARCH("特价",M1))     → M3 的 DisplayFormat = 14869246 ✔ 正确高亮
```
**根因**：`FormatConditions.Add(xlExpression, ...)` 的公式相对引用以**活动单元格**为基准。实现用目标区域左上角地址拼公式（`M1`），却没有先把活动单元格移到该处，也未改用 `R1C1`/绝对引用；于是基准为 `A1` 时整体平移 `列号差`（M=13 → Y=25，平移 +12；N=14 → AA=27，平移 +13 且行 +1 → `AA3`，与实测完全吻合）。
**影响**：任何"按文字/公式自动生成公式"的条件格式（`text_contains`）在活动单元格不等于区域左上角时**静默失效**——而"活动单元格不是区域左上角"正是真实使用中的常态。

#### H-2 `format_cells` 的 `verticalAlignment` 被忽略，且返回体谎报

**分类**：参数被忽略 + 返回体谎报
**最简复现**：
```
excel_format_cells { host:'wps', sheetName:'D6', address:'C1', verticalAlignment:'top' }   → success: true
```
**返回体**：`appliedStyles: { ..., "verticalAlignment": "top" }`
**宿主读数**：`Range('C1').VerticalAlignment = -4108`（center）← **期望 `-4160`(top)**
`bottom` 同样落到 `-4108`；`center` 恰好等于默认值因而"通过"。
**排除宿主不支持**：宿主原生 API 可用——
```
wps_execute_script: Range('A1').VerticalAlignment = -4160; → 读回 -4160 ✔
```
**根因**：宿主原生可用、工具返回体回显了请求值，但目标属性未被写入。属实现漏写（桥接或宿主侧未落到 `VerticalAlignment` 即返回 `appliedStyles`）。

#### H-3 8 个形状类工具 schema 缺 `workbookName` → 回退陈旧锁目标 → 工具不可用

**分类**：工具没暴露（**不是**宿主不支持）
**受影响工具**：`excel_add_shape`、`excel_group_shapes`、`excel_ungroup_shapes`、`excel_set_shape_zorder`、`excel_export_shape_image` 及其 5 个 `wps_*` 同名声（共 8 个形状/导图工具缺 `workbookName`）。
**最简复现**：
```
excel_add_shape { host:'wps', workbookName:'不存在簿zzz.xlsx', sheetName:'DO',
                  kind:'geometric', shapeType:'rectangle', left:10, top:10, width:50, height:20 }
→ 未在 WPS 中找到目标工作簿 [工作簿1]。当前已打开: AgentExcelProbe.xlsx
```
**关键对照（证明参数根本没到宿主）**：
```
excel_patch_cells { host:'wps', workbookName:'不存在簿zzz.xlsx', ... }
→ 未在 WPS 中找到目标工作簿 [不存在簿zzz.xlsx]。当前已打开: AgentExcelProbe.xlsx   ← 提到了我传的名字
excel_add_shape   { host:'wps', workbookName:'不存在簿zzz.xlsx', ... }
→ 未在 WPS 中找到目标工作簿 [工作簿1]。当前已打开: AgentExcelProbe.xlsx          ← 完全没提我传的名字
```
**根因链**：
1. `excel_add_shape` 等 8 个工具 schema 里没有 `workbookName`（快照核对：101 个表格工具中 15 个无此参数）；
2. 桥接层只转发 schema 内的字段，`workbookName` 被静默丢弃（**未**像 `save_workbook` 那样返回 `ignoredParams` 提示）；
3. 宿主侧 `wps-addon/src/shared.js:259` 回退：`const targetName = workbookName || lockedTargets.excel;`
4. `lockedTargets.excel` 是崩溃前残留的 `'工作簿1'`，该工作簿已不存在 → 必然失败。
**期望值**：形状工具的 schema 应包含 `workbookName`（`excel_add_shape` 的 gateway 处理器其实**已经**在转发 `args?.workbookName`，见 `src/bridge/gateway/excel.ts:854`——只差 schema 暴露），或在无法指定目标时明确报错而非回退到陈旧锁。
**绕行方式（已用于完成本批用例）**：`wps_lock_target_document {component:'excel', targetName:'AgentExcelProbe.xlsx'}` 后形状工具立即恢复正常（已实测：锁定后 `add_shape` 全部参数落地，见 §5）。这反证了"宿主完全支持、只是目标参数没接上"。

#### H-4 `create_workbook` 的 `savePath` 已存在时谎报"已保存"

**分类**：调用成功但没生效 + 返回体谎报
**最简复现**：
```
1) wps_create_workbook { savePath:'/Users/jolin/Downloads/AgentExcelProbe.xlsx', sheetName:'S1' }  → 真的落盘
2) 记录磁盘 mtime；再次 wps_create_workbook { savePath:'<同一路径>', sheetName:'S1' }
```
**返回体（第 2 次）**：
```json
{"success":true,"workbookName":"工作簿23","savedPath":"/Users/jolin/Downloads/AgentExcelProbe.xlsx",
 "warnings":[],"message":"已新建工作簿 [工作簿23]（1 张表，首表 S1）并保存到 /Users/jolin/Downloads/AgentExcelProbe.xlsx"}
```
**宿主读数**：`工作簿23 | <unsaved> | false`（`Name | Path | Saved`）
**磁盘读数**：mtime **与第 1 次完全相同**（`1790081259093.6982`），文件未被覆盖。
**期望值**：三者之一——报错、`warnings` 非空、或真的落盘。实测三者皆无。
**根因**：`wps-addon/src/excel.js:2185`
```js
try { wb.SaveAs(String(savePath), 51); saved = String(savePath); }
catch (e) { saveError = String(e.message || e); }
```
`app.DisplayAlerts = false` 下 WPS 对已存在路径的 `SaveAs` **既不抛异常也不落盘**（覆盖确认被静默吞掉），实现以"没抛异常"判定成功，`savedPath` 直接取请求路径，未做任何落盘核对。`warnings` 只在 `SaveAs` 抛异常时才有内容。

### 中

#### M-1 `manage_sheet action:'rename'` 改到已存在名字 → 谎报改名成功

```
excel_manage_sheet { host:'wps', workbookName:WB, sheetName:'TBm2', action:'rename', newName:'TB1' }
```
**返回体**：`success:true, oldName:'TBm2', newName:'TB1', message:'工作表 [TBm2] 已成功重命名为 [TB1]'`
**宿主读数**：逐表枚举仍是 `...,TBm2`（`TBm2` 未被改名，也没有被合并）
**期望**：报错（目标名已存在），或返回 `success:false`。

#### M-2 `manage_sheet action:'unprotect'` 错密码 → 谎报已解锁

```
1) excel_manage_sheet { ..., action:'protect',   password:'pw123' }  → 宿主 ProtectContents=true ✔
2) excel_manage_sheet { ..., action:'unprotect', password:'WRONG' }
```
**返回体**：`success:true, protected:false, message:'工作表 [TBm2] 已解除锁定保护'`
**宿主读数**：`ProtectContents = true`（**仍处于保护**）
**根因（已证实为宿主机为 + 工具未读回）**：宿主原生 API 对错密码**不抛异常**——
```
wps_execute_script: try{ ws.Unprotect('WRONG2'); return 'THREW_NOTHING 保护态='+ws.ProtectContents }catch(e){...}
→ THREW_NOTHING 保护态=true
```
工具把"没抛异常"当成成功，没有在写后读回 `ProtectContents` 校验。正确密码的 `unprotect` 工作正常（`ProtectContents` → `false`）。

#### M-3 `duplicate_sheet position:'end'` 被忽略（等同 `after`）

**源表不是最后一张时**是稳定的复现条件（源表已是最后一张时 `end` 与 `after` 结果巧合相同，会掩盖问题）：
```
枚举: 1:... 16:Zsrc 17:Ztail          ← 源表 Zsrc 不在末尾
excel_duplicate_sheet { ..., sourceSheetName:'Zsrc', newSheetName:'Zend', position:'end' }
宿主读数: Zend Index=17 / 总数=18      ← 期望 18（末尾），实际插在源表之后
```
**根因**：`wps-addon/src/excel.js:1994`
```js
if (position === "before") { srcSheet.Copy(srcSheet); }
else { srcSheet.Copy(null, srcSheet); }     // 'end' 落进 else，从未实现
```

#### M-4 `create_sheet` 非法表名/超长表名 → 静默创建 `SheetN` 并 `success:true`

```
excel_create_sheet { host:'wps', workbookName:WB, sheetName:'非法[名]' }
→ success:true, sheetName:'Sheet10', isNew:true, sheetCount:11
宿主枚举: ... 11:Sheet10        ← 请求的名字没有生效，实际建了一张 Sheet10
excel_create_sheet { ..., sheetName:'x'.repeat(40) }  → 同样得到 Sheet11
```
**期望**：非法/超长表名应报错（不能让调用方以为建了 `非法[名]`）。返回体虽然如实回报了 `sheetName:'Sheet10'`，但 `success:true` 会让只看 `success` 的调用方误判。
**对照（正常行为，非缺陷）**：重名时返回 `success:true, isNew:false`，表数不变——是幂等语义，合理。

#### M-5 `set_data_validation validationType:'list'` 缺 `listItems` → 静默 no-op

```
excel_set_data_validation { host:'wps', workbookName:WB, sheetName:'D6', address:'L1', validationType:'list' }
→ success:true, err: undefined
宿主读数: Range('L1').Validation → Type=null, Formula1=null（没有任何有效性被写入）
```
**期望**：报错（`list` 缺 `listItems` 无从构造下拉）。同族的 `manage_sheet rename` 缺 `newName` 能正确报错，说明该工具缺同类必填-搭配校验。

#### M-6 `set_filter_and_sort` 的 `sortRules.colIndex` 越界 → 静默 no-op

```
excel_set_filter_and_sort { ..., range:'A10:C13', sortRules:[{colIndex:99, order:'asc'}] }
→ success:true
宿主读数: A11:A13 = 甲,乙,丙   ← 数据完全未变（区域只有 3 列，colIndex 99 无效）
```
**期望**：报错或明确告警"排序列超出范围"。

#### M-7 `create_pivot_table` 字段名不存在 → 静默丢字段

```
excel_create_pivot_table { ..., sourceRange:'A1:C4', destCell:'A20', rowFields:['不存在的字段'] }
→ success:true, warnings:[]
宿主读数: PivotTable_1790081539141  rows=0 cols=0 data=0   ← 三组字段全空
```
**期望**：字段不存在应报错或至少进 `warnings`。实测建了一个空透视表并报成功。

#### M-8 `wps_get_locked_status` 的字段不可信

```
1) wps_lock_target_document { component:'excel', targetName:'AgentExcelProbe.xlsx' }   // 该簿确实打开着
2) wps_get_locked_status {}
```
**宿主/桥接读数**：
```json
{"gatewayLocks":{"excel":"AgentExcelProbe.xlsx"},
 "addonStatus":{"openWorkbooks":[], "openWorkbookCount":0,
   "lockedTargets":{"excel":"AgentExcelProbe.xlsx"},
   "lockTargetExists":{"excel":{"target":"AgentExcelProbe.xlsx","exists":false}}}}
```
**两处与事实矛盾**：
- `openWorkbooks: []` / `openWorkbookCount: 0`，而宿主当下确实打开着 `AgentExcelProbe.xlsx`（`excel_get_workspace_summary` 与 `app.Workbooks` 都能读到）；
- `lockTargetExists.excel.exists: false`，而锁定的正是已打开的真实工作簿。
**影响**：该工具是判断"锁目标是否还在"的唯一入口（描述里也强调它用于此），字段恒为空/误判成 `false` 会让调用方误以为"没有锁"或"锁目标已丢"，进而在缺陷 H-3 的场景里无法自查。

#### M-9 `wps_rollback` 的前置守卫使"写→读回核对→回滚"标准流程不可用

```
1) excel_patch_cells {..., address:'Z2', values:[['原值']]}
2) excel_patch_cells {..., address:'Z2', values:[['改后值']]}          → auditId = 896d1ac4-…
3) wps_execute_script: return Range('Z2').Value2                       ← 读回核对（本报告强制的验证步骤）
4) wps_rollback { auditId:'896d1ac4-…' }
```
**返回**：
```json
{"success":false,"error":"本记录（…）之后，工作簿 [AgentExcelProbe.xlsx] 上发生过未经快照记录的操作：
 wps_execute_script：任意原生脚本执行（可能重建或删除工作表、整体改写内容）（…，无法归属到具体工作簿）。
 此时即使区域内容与记录一致，也无法确认它仍是同一次修改的结果…已拒绝覆盖。"}
```
**宿主读数**：`Z2 = 改后值`（未回滚）。
**对照（回滚本身正常）**：不插入第 3 步的读回，同样的"写→回滚"：
```
wps_rollback → success:true, message:'成功恢复区域 $Z$5 的值与公式', identityChecks:[无更晚的同区域记录, 无未审计的身份变更操作, 目标工作表仍存在, …]
宿主 Z5 = 前值 ✔
```
**判定**：守卫本身是**正确且安全**的（宁拒不破坏），但它的判据把"读回核对"也算成不可归属的身份变更，导致**任何按规范先读回再回滚的流程都会被拒**。属可用性缺陷，建议区分"只读脚本"与"可写脚本"，或让读回类脚本不进入该台账。

#### M-10 `manage_named_range` 的 `comment` 被静默忽略；`add` 不稳定

**（a）`comment` 被忽略**（返回体无任何提示）
```
excel_manage_named_range { host:'wps', workbookName:WB, action:'add', name:'fName',
                           refersTo:'DO!$B$2:$B$4', comment:'备注F' }
→ success:true, verified:true, warnings:[]
宿主读数: Names.Item('fName').RefersTo == "DO!$B$2:$B$4" | Comment == ""    ← 期望 "备注F"
```
**（b）`add` 不稳定**：同一调用形态在不同时刻结果不一致——
- 成功例：`add {name:'gName', refersTo:'DO!$B$2:$B$4'}` → 宿主枚举出现 `gName`，`delete` 后枚举消失 ✔
- 失败例：`add {name:'hR1', refersTo:'DO!$B$2:$B$4', comment:'备注H'}` → 返回 `success:true` 但 `verified:false, warnings:["写入后读不回该名称"]`；宿主枚举 `CX!_FilterDatabase,DO!Print_Area` **完全没有 hR1**；单格引用 `hR2` 同样失败
**判定**：`verified:false` + warning 是诚实的，但 `success:true` 与"什么都没建"并存容易误判。**该条标「待确认」**——未定位到稳定触发条件（差异可能来自 `comment` 参数、或前序 `delete` 不存在的名字留下的状态）。

### 低

#### L-1 `format_cells` 的非法颜色值被静默接受

```
excel_format_cells { ..., address:'D1', backgroundColor:'不是颜色' }   → success:true, warnings: undefined
宿主读数: Range('D1').Interior.Color = 3351057 = 原值 bgr('#112233')（未变）
同样：'#GGGGGG' → success:true、无 warning
```
**期望**：非法颜色应报错（对比：非法 `fontSize:-5` 与非法 `include` 值都能正确报错，说明该校验缺失或位置不对）。

#### L-2 `freeze_panes freezeRowIndex:0`（越界值）产生随机冻结

```
基线: app.ActiveWindow → FreezePanes=false, SplitRow=0, SplitColumn=0
excel_freeze_panes { ..., freezeRowIndex:0, freezeColumnIndex:0 }
→ success:true, freezeRowIndex:1, freezeColumnIndex:1, message:'已成功锁定第 0 行之上的表头吸顶显示'
宿主读数: FreezePanes=true, SplitRow=0, SplitColumn=13     ← 冻结在无意义的位置
```
**期望**：`0` 应从 1 开始校验（描述明确"从 1 开始"）→ 报错，或视作"不冻结"。

#### L-3 `manage_pictures insert` 的 `left/top/width/height` 不被精确兑现

两次独立测量（返回体本身**如实**回报了实际值）：

| 请求 `left,top,width,height` | 宿主实际 | 偏差 |
|---|---|---|
| 600, 300, 64, 64 | 598.2, 286.4, 64, 60.8 | 左 −1.8 / 上 −13.6 / 高 −3.2 |
| 100, 100, 100, 50 | 99.2, 96, 99, 47.6 | 左 −0.8 / 上 −4 / 宽 −1 / 高 −2.4 |
| 400, 700, 120, 120 | 398.2, 667.2, 120, 114.4 | 左 −1.8 / **上 −32.8** / 高 −5.6，`LockAspectRatio=0` |

`LockAspectRatio` 为 0，排除"锁定纵横比"解释。同尺寸需求下 `add_shape`（几何形状）的 `left/top/width/height` **完全精确**（`20,200,180,70` 分毫不差），所以是图片插入路径特有。
**判定**：几何参数未精确生效；工具返回体诚实（回读了实际值）。根因**待确认**（WPS `AddPicture` 的缩放/锚点取整行为）。

#### L-4 `save_workbook` 对从未保存过的工作簿首次保存误报失败

```
excel_save_workbook { host:'wps', workbookName:'工作簿17' }
→ success:false, error:"WPS save_workbook 超时，结果未知；先读取状态，不要自动重放写入"
宿主读数（紧接着）：工作簿17.xlsx | full=/Users/jolin/Downloads/工作簿17.xlsx | saved=true
磁盘：文件确实生成
```
**判定**：保存**实际成功**，但调用方拿到失败。错误文案自身是诚实的（"结果未知，先读回"），问题是超时阈值对"未保存过的工作簿首次 SaveAs"偏紧。已落盘工作簿再次保存正常（`success:true`，很快返回）。

#### L-5 `manage_hyperlink`：schema 暴露 `targetAddress` 但单独传它不可用

```
excel_manage_hyperlink { ..., action:'add', address:'E2', targetAddress:'#A1', displayText:'内部跳转' }
→ success:false, error:"add 需要 url"
```
`url` + `targetAddress` 同传时两者都正确落到宿主（`Address=https://wps.cn | SubAddress=#Sheet1!A1 | TextToDisplay=两用`）。
**判定**：契约未声明 `targetAddress` 依赖 `url`；错误信息本身可读，属文档/契约缺陷（低）。

#### L-6 契约快照 p17 与运行中的服务不一致（3 处）

| 工具 | 快照 `tools-snapshot.p17.json` | 运行中的服务实测 | 影响 |
|---|---|---|---|
| `excel_update_chart.legendPosition` | `Top/Bottom/Left/Right/Corner` | 只接受 `right/left/top/bottom`（传 `Right` 报"不在允许值中（允许值：right / left / top / bottom）"） | 按快照调用必失败；`scripts/param-effect-test.mjs` 用的正是小写 `right`，说明脚本与快照已脱节 |
| `excel_export_shape_image` | 有 `outputPath`（必填） | 无 `outputPath`；允许 `shapeName/shapeId/format/scale/sheetName/host`，传 `outputPath` 报"未知参数" | 按快照调用必失败 |
| `excel_export_shape_image`（能力本身） | 标注为已实现 | `excel_export_shape_image` 在 WPS 宿主上未实现（加载项缺少该 RPC 分支），已拒绝且未执行 | 该工具在 WPS 侧**确实不可用**（属"宿主侧未接"，非我的测试问题） |

#### L-7 `get_range_styles` 的 `include:['validation']`：宿主能做，工具没暴露

**分类**：工具没暴露（**不是**宿主不支持）——本任务特别要求区分的情形。
```
excel_get_range_styles { ..., address:'H1', mode:'summary', include:['validation'] }
→ success:false, error:"arguments.include[0]: 不在允许值中（允许值：fontName / fontSize / bold / fontColor /
   backgroundColor / numberFormat / horizontalAlignment / verticalAlignment / wrapText / rowHeight /
   columnWidth / merged / mergeArea / borders）"
wps_get_range_styles 同样被拒
```
但：
- **工具自己的描述宣称可用**：同一定义文件里写着"可通过 include 追加 **validation**（数据有效性读回：类型名/操作符/公式1-2/…）"；
- **宿主实现已支持**：`wps-addon/src/excel.js:365` 的 `allowed` 数组**包含** `"validation"`（第 365 行），`excel.js:352` 起有对应读回分支；
- **宿主确实能读**：`wps_execute_script: Range('H1').Validation → Type=3, Formula1='已通过,待复测,已报废'` ✔
- **卡点在 schema**：`src/bridge/tools/definitions/excel.ts:118` 的 `items.enum` 少了 `validation`，被 `src/bridge/catalog.ts:113` 的枚举校验先行挡下。
**期望**：在 `excel.ts:118` 的 enum 补上 `"validation"`（一行），能力即贯通。

#### L-8 `save_workbook`：文档称"传 sheetName 会报未知参数"，实测不报错

描述原文：`参数集只有 workbookName（不接受 sheetName，传了会报"未知参数"）`。
实测：
```json
{"success":true, "ignoredParams":["sheetName"],
 "ignoredParamsNote":"以下参数对 excel_save_workbook 不适用，已忽略：sheetName（本工具允许的参数见 schema）"}
```
**判定**：实现行为比文档**更安全**（显式告知被忽略的参数，属正确做法），但描述与实现不符，建议改描述。真正的未知参数（如 `bogusParam`）确实会报错：`arguments.bogusParam: 未知参数；本工具允许的参数：workbookName`。

#### L-9 审计字段名不一致

`wps_get_audit_history` 的记录主键是 `id`，而写类工具（如 `excel_patch_cells`）返回的是 `auditId`。按 `auditId` 去匹配历史记录会取不到（我第一次就取错了）。建议统一字段名或双向兼容。

---

## 5. 已确认正常的行为（含 4 条易误判为缺陷的澄清）

这些**不是**缺陷，写出来是为了避免下一个会话重复排查：

| 项 | 现象 | 为什么不是缺陷 |
|---|---|---|
| `merge:true` | 宿主回 `$A$3:$C$3`，我断言 `A3:C3` 失败 | 绝对引用格式差异，合并本身正确 |
| `search_cells` / `set_filter_and_sort` `filterRange` | 回 `$A$3` / `$A$1:$C$4` | 同上；范围与筛选开关都正确 |
| `data_validation` number_range | 宿主 `Type=2`（`xlValidateDecimal`）而非我预期的 `Type=1` | 实现选了 decimal，语义正确；`Operator=1(between)/5(greater_than)`、`Formula1/2` 全部落地 |
| `data_validation` prompt | 我探测 `v.PromptTitle` 得 `undefined` | 宿主真实属性名是 `InputTitle`/`InputMessage`；用正确属性名验证**完全落地** |
| `set_sheet_view` zoom=1000 | `success` 或钳制 | 宿主 `Zoom` 未超上限 |
| `format_text_segment` | `find`/`start`+`length`/`occurrence` 全部按预期命中指定字符区间，未选区不受影响 | 局部富文本实现正确 |
| `add_chart` `position.leftCell:'F1'` | 我不是按 60px/列而是按 48px/列算，实际 `Left=240` | 默认列宽 ≈48px，工具正确（与既有 `param-effect-test` 的 `J1→432` 一致） |
| `add_shape` 的 `textAlign/textVAlign` | 空文字的形状上不生效 | **有文字时完全生效**（`right→-4152`、`bottom→-4107` 实测通过）；空 `TextFrame` 无对齐可言 |
| `update_shape` 的 `text` | 曾见 `形状E改改`（疑似追加） | 干净复现为**替换**（`AAAA` → `BBBB` ✔）。先前那次**未复现**，标「待确认」 |
| `update_shape visible/action:delete`、`manage_named_range delete`、`manage_document_properties delete` | 单对象探针显示"删了还在" | **集合枚举证明删除全部生效**：形状 `9→8` 且 `gDel` 消失；`Names 3→2` 且 `gName` 消失；`CustomDocumentProperties 4→3` 且 `gCustom` 消失。**是我的探针不可靠**（见 §7.16） |
| `configure_print_layout` `zoom` + `fitToPagesWide/Tall` 同传 | 宿主 `Zoom=false` | Excel 语义：`FitToPages` 与 `Zoom` 互斥。单独传 `zoom:88` → 宿主 `Zoom=88` ✔ |
| `configure_print_layout action:'read'` | 回 `settings.printArea = "$A$1:$C$4"` | 我断言漏了 `$`；读回正确 |
| `manage_table` | `apply/list/delete` + `tableName/hasHeaders/styleName/totalsRow` 全部落地，删除后 `ListObjects.Count=0` | 正确 |
| `manage_hyperlink` `url+targetAddress` | 两者都落地（`Address` + `SubAddress`） | 正确 |
| `pivot` 正常路径 | 名称/`SourceData`/`RowFields=1`/`DataFields=1` 都落地 | 正确（只有"字段不存在"是缺陷 M-7） |
| `export_shape_image` | 报"WPS 宿主上未实现（加载项缺少该 RPC 分支）" | **宿主侧未接**，如实拒绝、未静默执行——错误处理正确 |
| `wps_inspect_api` | `expression` 求值、`evaluate:true` 展开成员、不存在成员报"目标对象为空 (null/undefined)" | 正确 |
| `wps_get_audit_record` | 不存在的 auditId 报"找不到对应的审计记录" | 正确 |
| `wps_rollback` | 无中间脚本调用时回滚成功且宿主读数一致 | 正确（M-9 是守卫过严，不是回滚坏了） |
| `wps_execute_script` | `code`/`params`/`component`/`workbookName` 组合可用 | 全部宿主读数的通道，行为可信 |

---

## 6. 未覆盖清单

### 6.1 因事故未覆盖

| 范围 | 原因 |
|---|---|
| 其余用例在崩溃前的初跑结果 | 批次 C 首跑在 `工作簿1` 断连窗口内执行，全部宿主读数为 `<<HOST_ERR:WPS excel 加载项未连接>>`，**已作废并在恢复后重跑**，未计入任何结论 |
| 探测簿中批次 B 建立的工作表状态 | 崩溃时未保存，内存态丢失（预期的连带损失，非数据损坏） |

### 6.2 已覆盖但未逐参数穷尽（这些参数本次**没有**宿主读数证据）

- `excel_add_chart`：`cellRange`、`sourceAddress`、`dataRanges`、`startCell`、`endCell`、`left`、`top`、`width`、`height`、`smoothLine`、`seriesSettings`、`yAxis{min,max,step,title,numberFormat}`、`replaceExisting`
- `excel_update_chart`：`cellRange`、`startCell`、`endCell`、`name`、`shapeName`、`leftDelta` 类偏移
- `excel_delete_chart`：`chartName`、`chartTitle`、`name`、`shapeName`、`leftCell`、`clearAll`
- `excel_get_charts`：`chartTitle`、`shapeName`、`shapeSelfTest`、`selfTestKeep`（后者为 MS 侧诊断入口）
- `excel_add_conditional_formatting`：`iconThresholds`
- `excel_configure_print_layout`：`addHorizontalPageBreak`、`addVerticalPageBreak`、`clearPageBreaks`、`clearPrintArea`、`headerMargin`、`footerMargin`、`leftHeader`、`rightHeader`、`leftFooter`、`centerFooter`、`printTitleRows`、`printTitleColumns`
- `excel_manage_table`：`newName`
- `excel_manage_document_properties`：`properties` 的其他内置字段
- `excel_format_cells`：`borders: true`（只测了十六进制色）、`borders: 'none'/false`（描述已声明会被跳过）
- `wps_execute_script`：`documentName`、`presentationName`
- `excel_auto_fit_columns`：`minWidth` 的下限生效条件（本次自适应结果均大于 `minWidth` 或小于 `maxWidth`，未构造到钳制边界）

### 6.3 明确未测（非本任务范围 / 主动排除）

- `excel_*` 在 `host:'microsoft'` 侧的行为 → 归 `task-3` / `ms-tester`
- Word / PPT 工具 → 归 `task-2` / `word-ppt-tester`
- `office_lock_target` / `office_unlock_target`（`wps_*` 同名版本已测）
- `wps_reload_addon`、`wps_clear_audit_history` → 见 §3.3，主动排除

---

## 7. 测试脚本自身的错误（19 类，**均已从缺陷中剔除**）

按任务要求，如实区分"我测错了"。这些最初都表现为失败，逐条复核后确认是脚本问题：

1. **宿主读回未绑定 `workbookName`** → 活动簿漂移（当时活动簿是刚新建的空白簿），读回一片 `null`，误判"写入没生效"。修法：所有宿主读回显式带 `workbookName`。
2. `get_sheet_outline` 字段名猜错（实际是 `usedRangeAddress`/`rowCount`/`columnCount`，不是 `usedRange`）。
3. `search_cells` 断言忽略绝对引用（`$A$3`）。
4. `get_range_styles` 的 `maxCells` 用 `"address"` 出现次数判断（区域地址本身也算一次）——实际 `returnedCells=2 / truncated=true` 完全正确。
5. `get_range_styles include:['validation']` 我按"应该成功"断言——**应先质疑自己的期望**：这正是缺陷 L-7。
6. `get_sheet_outline` 列数期望 6（当时只写了 3 列）。
7. `duplicate_sheet` 的 `position` 断言公式写错（比较 `new === src` 而非 `new === src±1`）。
8. `manage_sheet rename` 断言用子串匹配：`includes('TBm')` 会被 `'TBm2'` 命中。
9. `modify_rows_columns insert` 期望串漏了空值（宿主打印 `null`）。
10. **`auto_fit_columns` 用了不存在的字段名 `column`**（正确是 `colIndex`）→ 验证器已正确报错 `arguments.columnRules[0].colIndex: 必填`，我却差点把它当成"参数被忽略"。改用 `colIndex` 后全部生效。
11. `cond_format` 探针在**重叠区域**上读 `FormatConditions.Item(1)`（读到的是外层区域的规则），并读取数据条规则不存在的 `.Interior`/`.Color` 属性 → 3 个假失败。修法：换到互不重叠的干净区域（`P1:P6`/`Q1:Q6`），此后 `between`(Type=1/Operator=1/F1=15/F2=25/色) 与 `data_bar`(Type=4/BarColor) **全部正确**。
12. `text_contains`/`formula` 规则断言写错（要求 `Type=9`，实际实现用 `Type=2` + `ISNUMBER(SEARCH(...))`）——**实现方式正确**；真正的问题（H-1）是**活动单元格依赖**，与 `Type` 无关。
13. `data_validation` prompt 探针用了宿主不存在的属性名 `PromptTitle`/`PromptMessage`（真实是 `InputTitle`/`InputMessage`）→ 假失败。
14. `data_validation` number_range 期望 `Type=1`，宿主用 `Type=2`（decimal）→ 假失败。
15. `set_filter_and_sort`/`add_chart position` 的期望值算错（忽略 `$`；按 60px/列 而非实测的 48px/列）。
16. **单对象探针判断删除结果**（`Shapes.Item(name)` / `Names.Item(name)` / `CustomDocumentProperties.Item(name)`）→ 假阳性"删了还在"。**改用集合枚举后，三处删除全部证明正确**。这条最值得记：判断"对象是否还在"必须用 `Count` + 逐项枚举，不能用 `Item(name)` 是否抛异常。
17. `configure_print_layout` 把 `zoom` 与 `fitToPagesWide/Tall` 混传 → 宿主按 Excel 语义把 `Zoom` 置 `false`，假失败。单独测 `zoom:88` 正常。
18. `export_shape_image` 按快照传 `outputPath` → 运行版无此参数，假失败（真问题是**快照漂移**，缺陷 L-6）。
19. 批次 D 首跑时形状工具全部失败，我最初怀疑"形状能力坏了"——实际是 H-3 的目标解析缺陷叠加崩溃残留锁，**不是形状功能本身**。

---

## 8. 复现材料说明

- 本报告所有调用都可通过 `POST http://127.0.0.1:19890/api/v1/tool/call`（`Authorization: Bearer $(cat ~/.wps-bridge/token)`，body `{"name":…,"arguments":{…}}`）原样重放。
- 宿主读数统一用 `wps_execute_script {component:'excel', code:'…', workbookName:'<目标簿>'}` 取得；片段已内联在每条缺陷的"宿主读数"里。
- 颜色期望值一律按 `bgr(hex) = r + g*256 + b*65536` 换算，本报告出现的关键值：`#FEE2E2 → 14869246`、`#112233 → 3351057`、`#0F9D58 → 5807375`、`#CBD5E1 → 14800331`、`#1F3864 → 6585503`… 均在正文标注了换算来源，未手写。
- 测试脚本为临时文件（`/tmp/excel-tests/*.mjs`，含 `lib.mjs` 的 `call`/`host`/`bgr` 三个核心函数），**未落入仓库**（本任务的写入范围只有本报告）。关键是 `host()` 必须带 `workbookName`，以及"删除类结论用集合枚举"这一条。

---

## 9. 处置与遗留

**已完成**

- 101 个表格侧工具的覆盖盘点，其中 71 个具备宿主读数证据，另有 10 个部分覆盖、2 个主动排除
- 21 条缺陷（高 4 / 中 10 / 低 7）全部带最简复现 + 宿主读数 + 期望值
- 崩溃后用户数据完整性核对（22 张表全部在位），`测试表格.xlsx` **未保存、未被覆盖**
- 全部探测产物已清理：`AgentExcelProbe.xlsx` 已关闭并删除文件、临时导出图片已删除、锁已解除

**遗留 / 建议下一步**

1. **无需依赖即可修**：H-2（`verticalAlignment` 落属性）、H-4（`create_workbook` 落盘核对）、M-1/M-2/M-5/M-6/M-7（补"写后读回/入参校验"）、M-3（`position:'end'` 分支）、L-1（颜色校验）、L-7（`excel.ts:118` enum 补 `validation`，一行）
2. **需要先修契约**：H-3（8 个形状工具 schema 补 `workbookName`）；同时建议让"参数被忽略"像 `save_workbook` 那样返回 `ignoredParams`，而不是静默丢弃
3. **需要产品决策**：M-9（回滚守卫是否放行只读脚本）、M-8（`wps_get_locked_status` 字段失真会掩盖 H-3 这类问题）
4. **待确认 2 条**：M-10 命名区域 `add` 的不稳定触发条件、L-3 图片几何偏差的宿主根因；`update_shape text` 疑似追加（未复现）
5. **快照需要刷新**：L-6 说明 `tools-snapshot.p17.json` 已与运行版脱节，`scripts/param-effect-test.mjs` 里 `update_chart.legendPosition:'right'`、`excel_add_shape` 的 `host` 等用法也需同步核对

**风险提示（给下一个会话）**

- 形状类工具在**存在陈旧锁目标**时整体不可用；先跑 `wps_get_locked_status` 无助于自查（M-8），可直接 `wps_lock_target_document` 锁到目标簿绕开。
- 判断"写入/删除是否真的生效"必须用宿主集合枚举，单对象探针会给出假阳性（§7.16）。
- 不要跑 `wps_reload_addon`（断全部队友连接）与 `wps_clear_audit_history`（销毁团队证据）。
