# 对外 MCP 能力全景盘点（P5 / mcp-sweep-01）

> 面向没参与过本项目的人：读完这一份就能知道这套 MCP 能干什么、分成哪些域、每个域有哪些工具、**哪些能真跑、哪些只是声明**，以及**一个 AI 只看工具说明能不能把它用起来**。

| 元数据 | 值 |
|---|---|
| 盘点时间 | 2026-09-22 15:17–15:22 CST |
| 仓库锚点 | `cac6d38`（工作区有未提交改动，未参与本次判定） |
| 服务版本 | Bridge 2.1.0（`service.json`），监听 `127.0.0.1:19890` |
| 宿主 | macOS（darwin），WPS 加载项已连接（`isWpsConnected: true`），Microsoft Office **未连接**（`isMsOfficeConnected: false`） |
| 目标工作簿 | `~/Downloads/钙钛矿各家企业现状.xlsx`（8 张工作表，含 2 张测试/盘点痕迹表） |
| 清单来源 | `GET /api/v1/mcp-tools`（正在运行的服务实例，非源码快照）；能力状态来自 `GET /api/v1/capabilities` 与 `bridge_diagnose` |
| 证据口径 | 只读探测：清单拉取 + 只读工具真实调用 + 源码静态核对。**本次未执行任何写入类工具**，也未保存/修改任何文件 |
| 内容边界 | 本报告不记录工作簿的业务数据；仅出现工作表名称、结构、计数与调用元数据 |

---

## 一、30 秒速览

- **91 个对外工具**，按能力域分成 **9 组**（§二、§三）。
- 其中 **21 个标注只读**、**70 个标注写入**（`annotations.readOnlyHint`）。
- 按"能不能真跑"分四档：

| 档位 | 工具数 | 含义 |
|---|---|---|
| 本次实测调用返回 `success`（全部只读） | 11 | 有真实返回体为证（§4.1） |
| 本次验证了错误路径（只读） | 1 | `wps_get_audit_record` 对不存在的 ID 正确报错 |
| 有真实写入留痕 | 2（`excel_patch_cells` / `wps_patch_cells`） | 审计留痕 34 条，见 §4.2 |
| **声明了但宿主没实现（必然失败）** | **2** | `excel_update_chart` / `wps_update_chart`（§五 P-03） |
| 声明可实现、本次未验证 | 75 | 写入类为主，受"只读"约束未执行 |
| （另一维度）依赖本机未连接的通道 | 3 个 `office_*` ＋ 27 个 `excel_*` 的 `host=microsoft` 分支 | Microsoft Office 未运行（§七） |

- **核心结论**：一个 AI 只靠 `tools/list` 里的 `description` + `inputSchema`，**不能**在本会话把任务做起来。最致命的一条是目标定位自举失败：说明把"目标工作簿"写成可选参数，而省略它时工具会回落到**上一个会话遗留的陈旧锁目标**并直接报错，而唯一能列出已打开文件的工具恰好就是失败的那两个（§五 P-02）。
- **两处"说明承诺了宿主没有的能力"**：`update_chart` 在 WPS 侧未实现却写成可用功能（P-03）；`search_cells` 声称能搜"文本或公式"，实测只搜显示值、**搜公式返回 0 命中且报成功**（P-05）。
- 说明质量整体偏薄：**91 条说明合计仅 6549 字符，均值 71 字**；592 个参数里 13 个完全没有说明，11 个连 JSON Schema `type` 都没有。

---

## 二、能力域分组总表

| # | 能力域 | 工具数 | 读/写 | 说明均值 | 能跑判定 | 代表工具 |
|---|---|---|---|---|---|---|
| A | 连接与诊断 | 2 | 2 / 0 | 52 字 | 全部实测跑通 | `bridge_get_capabilities` |
| B | 目标锁定 | 5 | 1 / 4 | 44 字 | WPS 侧 1 项实测、其余未验证；`office_*` 两项离线 | `wps_lock_target_document` |
| C | Excel 统一入口（`excel_*`） | 27 | 6 / 21 | 106 字 | 6 个只读全部实测跑通；21 个写入未验证（其中 1 个宿主未实现） | `excel_read_range` |
| D | Excel WPS 兼容名（`wps_*` 表格） | 27 | 6 / 21 | 59 字 | 同 C，与 C 一一同构 | `wps_read_range` |
| E | Word 文档（`wps_word_*`） | 12 | 1 / 11 | 51 字 | 未验证（WPS Word 组件本次 `connected: false`） | `wps_word_read_document` |
| F | PPT 演示（`wps_ppt_*`） | 9 | 2 / 7 | 62 字 | 未验证（PPT 组件在线但本次无打开的文稿） | `wps_ppt_generate_deck` |
| G | 原生脚本与 API 反射 | 2 | 0 / 2 | 159 字 | 未验证（会改文档，未执行） | `wps_execute_script` |
| H | Microsoft 原生通道（`office_*`，不含锁） | 3 | 1 / 2 | 78 字 | 1 个实测跑通；2 个离线 | `office_get_status` |
| I | 审计与回滚 | 4 | 2 / 2 | 18 字 | 查询实测跑通；`wps_rollback` 有 1 条真实回滚留痕 | `wps_get_audit_history` |
| | **合计** | **91** | **21 / 70** | **71 字** | | |

**读/写口径**：`读` = `annotations.readOnlyHint === true`（服务自己声明的口径）。语义上另有 2 个"不改文档但会写本地文件/本地状态"的工具（`wps_capture_sheet_preview`、`wps_clear_audit_history`）被算作写，这是有意的（见 `src/bridge/contracts/host-methods.ts:45-46,56-58`），但从调用方视角看，"只读"标注与实际副作用并不一致，说明里也没有交代。

**C 与 D 是同构双份**：27 个 `excel_*` 由 27 个 `wps_*` 机械派生——换前缀、加一个必填 `host`、追加一句固定样板（`src/bridge/tools/excel.ts:20-35`）。除了 `host`，两者 schema 完全相同，说明文字只差那句样板。因此**91 个工具里真正独立的能力只有 64 个左右**。

---

## 三、逐域工具清单

图例：`读`/`写` = `readOnlyHint`；能跑状态 = `实测`（本次只读真实调用成功）/ `留痕`（有真实运行记录）/ `未验`（声明可实现，本次未执行）/ `缺口`（声明存在但宿主未实现）/ `离线`（依赖本机未连接的通道）。

### A. 连接与诊断（2）

| 工具 | 读/写 | 能跑 | 说明字数 | 备注 |
|---|---|---|---|---|
| `bridge_get_capabilities` | 读 | 实测 | 40 | 一次拿全平台/连接/实现状态/回滚覆盖/限制 |
| `bridge_diagnose` | 读 | 实测 | 65 | 与上者高度重叠；说明里指路 `office_get_status` 是有效的 |

### B. 目标锁定（5）

| 工具 | 读/写 | 能跑 | 说明字数 | 备注 |
|---|---|---|---|---|
| `wps_lock_target_document` | 写 | 未验 | 69 | 必填 `component` + `targetName` |
| `wps_unlock_target_document` | 写 | 未验 | 27 | 零必填 |
| `wps_get_locked_status` | 读 | 实测 | 42 | 说明承诺"及已打开的文件列表"，实测未返回（P-02） |
| `office_lock_target` | 写 | 离线 | 54 | 与 `wps_lock_target_document` 说明几乎同义 |
| `office_unlock_target` | 写 | 离线 | 27 | `component` 参数无说明 |

### C. Excel 统一入口（27，全部必填 `host`）

| 工具 | 读/写 | 能跑 | 说明字数 | 备注 |
|---|---|---|---|---|
| `excel_get_workspace_summary` | 读 | 实测 | 115 | 省略 `workbookName` 会失败（P-02） |
| `excel_get_sheet_outline` | 读 | 实测 | 99 | 同上 |
| `excel_read_range` | 读 | 实测 | 73 | 默认值与返回量控制写得清楚 |
| `excel_get_range_styles` | 读 | 实测 | 109 | `mode`/`maxCells` 说明清楚 |
| `excel_search_cells` | 读 | 实测 | 70 | 说明称能搜"文本或公式"，实测搜不到公式（P-05）；无检索范围参数 |
| `excel_get_charts` | 读 | 实测 | 131 | |
| `excel_create_sheet` | 写 | 未验 | 68 | |
| `excel_delete_sheet` | 写 | 未验 | 65 | |
| `excel_patch_cells` | 写 | 留痕 | 90 | 值与公式互斥规则未写（P-07） |
| `excel_format_cells` | 写 | 未验 | 136 | 参数级建议最丰富的一条 |
| `excel_add_conditional_formatting` | 写 | 未验 | 113 | |
| `excel_freeze_panes` | 写 | 未验 | 95 | 零必填 |
| `excel_modify_rows_columns` | 写 | 未验 | 104 | 与下一行功能重叠（P-01） |
| `excel_manage_rows_and_columns` | 写 | 未验 | 108 | 多 `set_size`；`index` 无 type |
| `excel_auto_fit_columns` | 写 | 未验 | 84 | |
| `excel_capture_sheet_preview` | 写 | 未验 | 117 | 会落盘本地图片；说明未交代 |
| `excel_add_chart` | 写 | 未验 | 129 | 零必填，数据源非必填（P-12） |
| `excel_update_chart` | 写 | **缺口** | 108 | WPS 宿主未实现（P-03） |
| `excel_delete_chart` | 写 | 未验 | 97 | |
| `excel_create_pivot_table` | 写 | 未验 | 116 | `dataFields` 聚合方式未说明 |
| `excel_set_filter_and_sort` | 写 | 未验 | 103 | `sortRules` 结构未说明 |
| `excel_set_data_validation` | 写 | 未验 | 112 | |
| `excel_manage_sheet` | 写 | 未验 | 114 | |
| `excel_manage_cell_comments` | 写 | 未验 | 122 | `address` 非必填但 `add/read/delete` 都需要 |
| `excel_find_and_replace` | 写 | 未验 | 145 | 写工具却承担查询职责；`searchQuery` 无 type |
| `excel_duplicate_sheet` | 写 | 未验 | 129 | |
| `excel_save_workbook` | 写 | 未验 | 101 | 不接受 `sheetName`（ISS-03） |

### D. Excel WPS 兼容名（27，与 C 一一同构，无 `host`）

`wps_get_workspace_summary`、`wps_get_sheet_outline`、`wps_read_range`、`wps_get_range_styles`、`wps_search_cells`、`wps_get_charts`（读，前 6 项能跑状态同 C）；`wps_create_sheet`、`wps_delete_sheet`、`wps_patch_cells`、`wps_format_cells`、`wps_add_conditional_formatting`、`wps_freeze_panes`、`wps_modify_rows_columns`、`wps_manage_rows_and_columns`、`wps_auto_fit_columns`、`wps_capture_sheet_preview`、`wps_add_chart`、`wps_update_chart`（**缺口**）、`wps_delete_chart`、`wps_create_pivot_table`、`wps_set_filter_and_sort`、`wps_set_data_validation`、`wps_manage_sheet`、`wps_manage_cell_comments`、`wps_find_and_replace`、`wps_duplicate_sheet`、`wps_save_workbook`（写，未验）。
说明文字 = C 列同名工具说明去掉那句样板，因此**信息量严格少于 C**，其余问题完全一致。

### E. Word 文档（12）

`wps_word_read_document`（读，未验）｜`wps_word_create_document`、`wps_word_save_document`、`wps_word_close_document`、`wps_word_manage_content`、`wps_word_write_content`、`wps_word_format_document`、`wps_word_insert_table_of_contents`、`wps_word_manage_table`、`wps_word_review_and_comments`、`wps_word_page_layout_and_watermark`、`wps_word_find_and_replace`（写，未验）。

这一族**最突出的契约问题是"目标文档"**：12 个工具里 11 个有 `documentName`，**必填数为 0**；另 1 个（`wps_word_create_document`）根本没有这个参数。说明统一写"目标文档名称，不传则默认当前活动文档"。本次 WPS Word 组件 `connected: false`，未实测；但按 §五 P-02 的同一套目标解析逻辑（`wps-addon/src/shared.js:188`），省略文档名会回落到粘性 `lockedTargets.word`。

### F. PPT 演示（9）

`wps_ppt_read_presentation`、`wps_ppt_get_slide_shapes`（读，未验）｜`wps_ppt_generate_deck`、`wps_ppt_manage_slides`、`wps_ppt_manage_table`、`wps_ppt_add_business_cards`、`wps_ppt_insert_native_chart`、`wps_ppt_manage_shapes_and_media`、`wps_ppt_capture_slide_preview`（写，未验）。

9 个工具全部有 `presentationName`，**必填数为 0**。这一族的**说明质量是全清单最好的**：`wps_ppt_manage_shapes_and_media`（135 字）写明了单位（pt）、要求先读页面尺寸、新增对象必须显式给坐标字号、编辑后读回并预览——这是"AI 能照着做"的范例。`wps_ppt_generate_deck` 也写明了要检查 `layoutWarnings` 并逐页预览。

### G. 原生脚本与 API 反射（2）

| 工具 | 读/写 | 能跑 | 说明字数 | 备注 |
|---|---|---|---|---|
| `wps_execute_script` | 写 | 未验 | 266 | **全清单说明最完整的一条**：写清了"专用工具未覆盖时使用"、注入变量、返回约定、先探测后执行 |
| `wps_inspect_api` | 写 | 未验 | 52 | 标注为非只读是有意的（执行任意表达式可能有副作用，`host-methods.ts:48-49`），但说明只写"自省探测"，读起来像安全只读 |

### H. Microsoft 原生通道（3）

| 工具 | 读/写 | 能跑 | 说明字数 | 备注 |
|---|---|---|---|---|
| `office_get_status` | 读 | 实测 | 78 | 返回 `runningComponents.excel: false`，word/ppt 报 "Application can't be found."，且**不区分"没安装"与"没运行"** |
| `office_execute_script` | 写 | 离线 | 106 | **交付说明被装配层正则截断成残句**（P-08） |
| `office_capture_slide_preview` | 写 | 离线 | 50 | 会落盘本地文件 |

### I. 审计与回滚（4）

| 工具 | 读/写 | 能跑 | 说明字数 | 备注 |
|---|---|---|---|---|
| `wps_get_audit_history` | 读 | 实测 | **17** | 9 个参数**全部无说明**（P-06） |
| `wps_get_audit_record` | 读 | 实测 | **16** | `auditId` 无说明 |
| `wps_rollback` | 写 | 留痕 | 23 | 只接受 `auditId`；覆盖范围未写（P-04、P-14） |
| `wps_clear_audit_history` | 写 | 未验 | **15** | 零必填、零参数说明、不可恢复 |

---

## 四、能跑与不能跑：证据

### 4.1 本次只读实测返回 `success`（11 个）

全部带显式目标（`host=wps` + `workbookName` + `sheetName`），返回 `success: true`：

| 工具 | 返回体关键字段 |
|---|---|
| `bridge_get_capabilities` | `version/platform/connection/hosts/audit/lifecycle/limitations` |
| `bridge_diagnose` | 同 capabilities + `hosts.wps.excel.unimplementedOnHost` |
| `office_get_status` | `runningComponents` / `openDocuments` / `errors` |
| `wps_get_locked_status` | `gatewayLocks: {}`、`addonStatus: null` |
| `excel_get_workspace_summary` | `workbookName/fullName/openWorkbooks/activeSheetName/sheetCount/sheets/selection` |
| `excel_get_sheet_outline` | `usedRangeAddress/rowCount/columnCount/headerPreview` |
| `excel_read_range` | `values/formulas/numberFormats` |
| `excel_get_range_styles` | `mode/styles/mixedOrUnavailableFields` |
| `excel_search_cells` | `totalFound/matches` |
| `excel_get_charts` | `totalChartCount/charts` |
| `wps_get_audit_history` | `records/total`（当前 `total: 34`） |

另 1 个只读工具只验证了错误路径：`wps_get_audit_record` 对不存在的 `auditId` 正确返回失败。

**未按名验证的对应项**：6 个 `wps_*` 只读兼容名（`wps_get_workspace_summary` 等）与 `excel_*` 同源实现（`wps-addon/src/excel.js` 同一函数），但我实测时 `wps_*` 侧只成功调用了审计工具，故**不把它们计入"已实测"**。同理，上表 11 个只读工具全部实测成功，但这**不代表**其余 80 个工具可用。

### 4.2 有真实运行留痕（写入侧）

`~/.wps-bridge/audit_history.json` 与 `wps_get_audit_history` 共同显示：**34 条记录、33 条 `applied`、1 条 `rolled_back`，`actionType` 只有 `update_values`(26) 与 `update_formulas`(8)**，即只有 `patch_cells` 被真正跑过。这同时**反证了 `wps_rollback` 至少成功执行过一次**。

工作簿侧的独立佐证：8 张工作表中，2 张带明显测试/盘点痕迹（`__MCP验收测试__`、`sweep图形_钙钛矿产业信息图`），说明结构性写入（建表一类）确实在真实宿主上落到过工作簿，不是纸面能力。

### 4.3 声明了但宿主没实现（必然失败）

`GET /api/v1/capabilities` 的 `hosts.wps.excel.unimplementedOnHost: ["update_chart"]`。契约层注释写得很明确（`src/bridge/contracts/host-methods.ts:32-34`）：

> `wps-addon/addon-core.js` 的 RPC 分发 switch 只有 add_chart / get_charts / delete_chart，没有 update_chart；未识别方法会走到 `default:` 抛出"未知的 RPC 方法"。因此 `excel_update_chart`（host=wps）**必然失败**。

我按只读约束**没有实际触发**它（它会改图表），判定依据是能力声明 + 路由表 + 宿主分发的静态核对。

### 4.4 宿主方法声明了但没有对应工具（AI 看不到）

`declaredNotCallable: ["clear_range", "insert_dimension", "rollback_cells"]`——这三个是宿主方法名，`tools/list` 里不存在同名工具，调用方无法感知也不会误用。列为"声明与可调用工具的差额"，不构成缺陷。

---

## 五、说明质量判断：一个 AI 只看说明能不能用起来？

**结论：不能完整自举，能勉强干活。** 读取类工具的说明基本够用（参数、默认值、返回量控制写得清楚）；一旦进入"定位目标 → 写入 → 校验"的闭环，说明层的缺口会直接导致失败或数据风险。按严重度排列如下：**P-01 至 P-05 是本次最差的 5 条**（选型无指引、目标自举死锁、两条"假能力"、一条"假承诺"），P-06 紧随其后。

---

### P-01（高）没有"我该用哪个"的指引；同能力双份工具在 `tools/list` 里无法区分

**证据**：27 对工具完全同构，说明文字只差一句样板。原文对照：

```
wps_read_range   : "切片读取指定区域(如 A1:C10)的单元格值与公式"
excel_read_range : "切片读取指定区域(如 A1:C10)的单元格值与公式 统一表格入口；host 必填。Windows Microsoft Excel 尚待实机验收。"
```

后台唯一的选型规则写在 **MCP 服务端 `instructions`** 字段里（握手时下发，不在 `tools/list` 中）：

> 优先使用 excel_* 结构化工具，显式选择 host=wps 或 microsoft，并指定 workbookName、sheetName。…… wps_* 兼容工具只控制 WPS。

**问题**：
1. `tools/list` 里没有任何字段告诉 AI "这两个是同一件事，选一个即可"。AI 要么随机选，要么把 27 对全读一遍（清单 118 KB，其中大量是这份重复）。
2. 同域内还有**真正意义重复**的一对：`excel_modify_rows_columns`（insert/delete/hide/unhide）与 `excel_manage_rows_and_columns`（多了 `set_size`，`index` 还能传列字母）。两份说明都没写"该用哪个、另一个是否已废弃"；`wps_` 侧同样成对存在。4 个工具做 2 件事。
3. `excel_find_and_replace` 一个"写"工具同时承担查询职责（`replaceText` 可选），而只读的 `excel_search_cells` 只能搜文本、不能搜公式——说明里没有交叉指路。

**影响**：选型成本直接转成 token 成本与误用风险；同一个任务在不同会话可能选中不同工具，行为不一致。

---

### P-02（高）"目标文档"声明为可选、实际必填；省略时回落到陈旧锁目标，且自举路径被自己堵死

**证据 1（schema 声明）**：27 个 `excel_*` **全部**有 `workbookName` 参数，**必填数为 0**。它的说明原文：

> 目标工作簿名称，例如 '明细.xlsx'，防止多文件焦点漂移，请先从工作区摘要确认目标。

`wps_get_workspace_summary` 与 `wps_get_sheet_outline` 更是**零必填参数**。Word 族 11 个 `documentName`、PPT 族 9 个 `presentationName`，必填数同样全是 0。

**证据 2（实测行为）**：省略目标名时全部失败，且错误里出现的是**另一个文件**：

| 调用 | 结果 |
|---|---|
| `wps_get_sheet_outline {}` | `success:false`，`未在 WPS 中找到目标工作簿 [经营&信息化 绩效完成情况.xlsx]。当前已打开: 钙钛矿各家企业现状.xlsx` |
| `excel_read_range {host:"wps", address:"A1:B2"}` | 同上 |
| `excel_get_sheet_outline {host:"wps", sheetName:"行业概览"}` | 同上（给了工作表名也救不回来） |
| `wps_get_workspace_summary {}` | 同上 |
| 对照：`excel_get_workspace_summary {host:"wps", workbookName:"钙钛矿各家企业现状.xlsx"}` | `success:true` |

**根因**（静态核对）：`wps-addon/src/shared.js:259` 的目标解析是 `workbookName || lockedTargets.excel`，而 `lockedTargets.excel` 是一个**跨会话残留的粘性锁**（`shared.js:285`：只有恰好 1 个工作簿时才回填）。锁指向的文件已被关闭，于是所有"不传目标"的调用都撞在这把陈旧锁上。

**为什么这是最严重的一条**：
- 一个新会话自然的第一步是"先看看打开了什么"，而**唯一能列出已打开文件的工具就是失败的那两个**；
- `wps_get_locked_status` 的说明写着"查询当前 Word、Excel、PPT 各组件的目标文档锁定状态**及已打开的文件列表**"，实测返回 `{"gatewayLocks":{},"addonStatus":null}`，**没有任何文件列表**——承诺的自举路径不存在；
- 实际可用的自举方式是从 `bridge_get_capabilities` 的 `components.excel.summary.message` 里读那句"当前已打开: …"，即**从错误文案里反推文件名**，而同一响应里的 `hasOpenWorkbook` 还是 `false`。这是偶然可用的通道，不是设计出来的路径。

**影响**：AI 会判定"用户没打开表格"并直接放弃任务（与已登记的 ISS-02 同源，但本条是**说明层**问题：schema 说可选、说明没说省略会怎样、实际必填且回落到脏状态）。修法方向应是三选一：把 `workbookName` 提为必填并同步说明；或省略时改为"当前活动工作簿"而非粘性锁；或提供真正的 `bridge_list_open_documents` 工具。

---

### P-03（高）`update_chart` 说明写成可用能力，WPS 宿主根本没实现

**说明原文**（`excel_update_chart`，108 字；`wps_update_chart` 为其去掉样板版）：

> 更新工作表中已有图表的位置、标题或图例。支持传入 cellRange（如 'I8:P20'）实现实机单元格刚性重定位吸附。 统一表格入口；host 必填。Windows Microsoft Excel 尚待实机验收。

**实际**：`capabilities.hosts.wps.excel.unimplementedOnHost: ["update_chart"]`；契约层注释写明"必然失败"（见 §4.3）。

**问题**：说明里没有任何"WPS 侧不支持"的字样，反而给了详尽的参数用法与"实机单元格刚性重定位吸附"的效果承诺。AI 会按说明去调用 → 报"未知的 RPC 方法" → 为了达成"更新图表"的目标，很可能改走"删掉重建"（`delete_chart` + `add_chart`），产生说明里完全没提示的副作用链。

**建议**：至少在两条说明里写明 WPS 缺口与替代路径（先 `get_charts` 读回、删旧建新时保留锚点），或在 `host=wps` 时直接拒绝并给出可执行建议。

---

### P-04（高）`wps_rollback` 说"原地恢复表格"，没说只覆盖一类修改

**说明原文**（23 字，全文）：

> 根据留痕记录 ID 一键撤销修改，原地恢复表格

**实际**（`capabilities.audit`）：

```
covered:   ["patch_cells 值与公式"]
uncovered: ["样式", "图表", "结构修改", "Word/PPT", "原生脚本"]
rollback:  "检查当前值和公式与记录的修改后快照一致；有后续修改则拒绝覆盖"
```

**问题**：说明里"一键撤销修改，原地恢复表格"是完全无条件的一般化承诺。AI 做完一轮格式化、建完图表之后，若想"反正能回滚"，就会以为有安全网——实际格式、图表、行列结构、Word/PPT、脚本全部**不可回滚**。此外"有后续修改则拒绝覆盖"这条重要的失败条件也没写进说明。

**建议**：把覆盖范围与拒绝条件写进 `wps_rollback` 说明本身（而不是只放在 `bridge_get_capabilities` 里），并在 `wps_get_audit_history` 的返回说明里标注"本条记录是否可回滚"。

---

### P-05（高）`search_cells` 声称能搜"公式"，实测搜不到；且报错要求"缩小检索范围"却不给这个参数

**说明原文**：

> 在表格中快速搜索包含指定文本或**公式**的单元格坐标

**实测反证**（只读，严格复现）：先在 `__MCP验收测试__`（测试表）用 `excel_read_range` 定位到一个**确认含公式**的单元格——该格 `formulas` 字段返回形如 `=B2+B3` 的运算式，而 `values` 字段返回其计算值，两者不同，足以证明它是公式格而非文本。随后用 `excel_search_cells` 检索该公式文本：

| 步骤 | 调用 | 结果 |
|---|---|---|
| 1 | `excel_read_range {sheetName:"__MCP验收测试__", address:"A1:H40"}` | 找到公式单元格 R4C2，公式文本与显示值不同 |
| 2 | `excel_search_cells {query:"<该公式去掉等号后的文本>", sheetName:"__MCP验收测试__"}` | `success:true`，**`totalFound: 0`**，未命中 R4C2 |

**根因**（静态核对）：`wps-addon/src/excel.js:289-322` 的 `searchCells` 只取 `usedRange.Value2`（显示值），**从不读 `.Formula`**。说明里的"或公式"是纯虚假能力。

**附带问题**：同一函数的守卫是

```js
if (usedRange.Rows.Count * usedRange.Columns.Count > 100000)
  throw new Error("已用范围超过 100000 单元格，请先缩小检索范围");
```

而 `excel_search_cells` 的参数只有 `query` / `sheetName` / `workbookName` / `maxResults`——**根本没有检索范围参数**。也就是说，在大表上这条错误信息要求调用方做一件该工具做不到的事；有 `searchRange` 的是另一个工具 `excel_find_and_replace`，而它同时是个写工具，说明里没有任何指路。

**影响**：AI 想"找出所有含公式的单元格"时会得到"0 命中"这个**看似成功的错误答案**，进而得出"表里没有公式"的错误结论——比报错更危险。加上 P-02 的目标问题，这是本次盘点中第二处"说明承诺了宿主没有的能力"。

---

### P-06（高）审计族说明近乎空白：15–17 字说明 + 9 个参数全部无说明

**说明原文（全文）**：

```
wps_get_audit_history  : "分页筛选修改记录；默认 5 条摘要"      （17 字）
wps_get_audit_record   : "按审计 ID 读取完整记录和快照"          （16 字）
wps_clear_audit_history: "清空本地存储的全部修改记录留痕"        （15 字）
```

**问题**：
- `wps_get_audit_history` 有 9 个参数（`limit`、`offset`、`view`、`workbookName`、`sheetName`、`clientName`、`actionType`、`status`、`fromTimestamp`、`toTimestamp`），**全部没有 description**——这 10 个无说明参数占了全清单 13 个"无说明参数"的 10 个。
- `view`/`status` 有枚举值但无解释；`actionType` 的取值域（`update_values`/`update_formulas`）只能从返回值里反推。
- 说明提到"默认 5 条摘要"，但返回体里的 `total` 字段、`offset` 与 `total` 的分页配合方式都没写。
- `wps_get_audit_record` 的 `auditId` 无说明，也没说"传 ID 会返回修改前后快照，可能包含单元格原文"——调用方不知道这个接口会把数据内容读出来。
- `wps_clear_audit_history` 是不可逆的破坏性操作，零参数说明，也没有"不可恢复"的警示。

**影响**：AI 无法构造有效的分页/筛选查询，只能盲试；对审计接口的返回内容（含数据快照）没有预期。

---

### P-07（中高）`patch_cells` 的值/公式语义没写清，混用会静默清空且仍报成功

**说明原文**：

> 直接在当前打开的 WPS 表格中原地修改数值或公式，自动抓取快照并实时呈现在用户屏幕上
> `values`: 二维数组数值
> `formulas`: 二维数组公式，例如 [['=A2*1.1']]

**问题**：两个参数都没有 `required`，说明**完全没讲同时传入会怎样**，也没讲"每个单元格二选一"。已登记的 ISS-01 实测：同时传 `values` 与含空串的 `formulas`，第二步把第一步刚写的值覆盖成空，而返回给调用方的文案是"已成功修改 4 个单元格"。说明层只要能加一句"两者互斥，不要同时传"，这条数据丢失风险就能被调用方规避。

**附带**：`excel_patch_cells` 说"实时呈现在用户屏幕上"——这是宿主激活的副作用，不是可依赖的保证，说明把它写成了功能承诺。

---

### P-08（中）11 个参数没有 JSON Schema `type`，1 个用了非标准类型

完整清单：

| 工具 | 参数 | 现状 |
|---|---|---|
| `excel_manage_rows_and_columns` / `wps_manage_rows_and_columns` | `index` | 无 `type`，说明写"数字……或列标识（数字 2 或字母 'B'）"——真值域是 number ∪ string，schema 里是空的 |
| `excel_find_and_replace` / `wps_find_and_replace` | `searchQuery` | 无 `type`，说明写"关键词、数值或错误标识" |
| `wps_word_write_content` | `content` | 无 `type`，说明写"单行字符串或多行字符串数组" |
| `wps_ppt_manage_table` | `shapeId` | 无 `type`，无说明 |
| `wps_ppt_manage_shapes_and_media` | `shapeId` / `shapeId1` / `shapeId2` | 无 `type`，说明只写"形状 ID / 标识" |
| `excel_format_cells` / `wps_format_cells` | `borders` | `type: ["string","boolean"]`——JSON Schema 的 `type` 不接受数组，严格校验的客户端会直接拒绝整份 schema |
| `wps_word_format_document` | `margins` | `object` 类型但**无说明**，属性未知 |

**影响**：按 schema 自动生成入参（表单、函数调用约束、类型检查）的客户端拿不到信息；`shapeId` 到底是数字还是字符串 ID 只能靠试。**注**：`excel_manage_rows_and_columns.index` 缺 `type` 可能是有意兼容 number/string，但既然 schema 无法表达联合类型，就更应该在说明里给出确切形态。

---

### P-09（中）交付说明被装配层正则误伤，出现残句

**交付清单里 `office_execute_script` 的说明原文**（注意结尾是逗号）：

> 直接向 Microsoft Word、Excel、PowerPoint 运行实例执行原生自动化代码（macOS 下为 JXA/AppleScript，Windows 下为 PowerShell COM 自动化），

**源码里的完整说明**（`src/bridge/tools/definitions/microsoft.ts:62`）：

> 【Microsoft Office 原生脚本引擎】直接向 …（同上）…，实现 100% 任意操作无死角！

清洗规则在 `src/bridge/tools/index.ts:22-23`：删 `【…】` 标记、删"实现 100% 任意操作无死角！"这类营销尾巴。规则本身合理，但**没有处理删除后留下的悬空逗号与断句**，于是交付给 AI 的是一条读起来像被截断的说明。同一规则对 `office_get_status`（"…完全免配置插件。"）等其它条的清洗是干净的，只有这一条断在标点上。

**影响**：AI 会怀疑说明不完整，进而无法确认脚本语法/返回约定（这条工具恰恰是 Microsoft 侧唯一的逃生通道）。

---

### P-10（中）"尚待实机验收"的措辞把范围写窄了，读者会以为 macOS 已验证

27/27 个 `excel_*` 的说明都追加同一句：

> Windows Microsoft Excel 尚待实机验收。

而 `capabilities` 的原话是：

> 代码路径已实现，但本候选版本**尚未在真实 Microsoft Excel 上做实机验收**；静态与模拟通道不代表实机通过

以及 WPS 侧：

> 既有功能曾在 macOS 使用；**2.0 回归需实机确认**

**问题**：说明只点名 Windows，读者（AI 或人）很容易推断"macOS 的 Microsoft Excel 已验证过"。实际是**全平台都没做实机验收**；WPS 侧也还有"2.0 回归待确认"的保留。这是"说明与实际状态不符"的典型：不是写错，而是写漏导致误导。

---

### P-11（中）枚举成员与说明不齐，同义别名优先级未定义

| 工具 | 缺口 |
|---|---|
| `wps_word_format_document` | `preset` 枚举含 `academic`，参数说明只解释了 `gov_standard` / `business_modern` / `custom` 三项 |
| `wps_ppt_insert_native_chart` | `chartType` 枚举含 `bar_stacked`，参数说明列了另外 6 项，独缺这一个 |
| `wps_ppt_manage_shapes_and_media` | `action` 枚举含 `delete_shape`，工具说明正文没有提到删除动作 |
| `excel_capture_sheet_preview` | 同时提供 `address` / `range` 与 `chartName` / `name` 两组同义别名，说明未写优先级，也未写同时传入会怎样 |
| `excel_add_chart` | 同时接受 `dataRange` / `sourceAddress`；说明只在 `dataRange` 处写"与 dataRanges 选其一传入"，`sourceAddress` 未纳入该约束 |

**影响**：AI 无法确定这些分支是否可用；别名冲突时只能靠试错。

---

### P-12（中）一批写工具零必填参数（连目标都不是必填）

**零必填参数的写工具共 17 个**（全清单 28 个零必填工具中，11 个是只读工具）：`wps_freeze_panes`、`wps_auto_fit_columns`、`wps_capture_sheet_preview`、`wps_add_chart`、`wps_update_chart`、`wps_delete_chart`、`wps_save_workbook`、`wps_word_create_document`、`wps_word_save_document`、`wps_word_close_document`、`wps_word_format_document`、`wps_word_insert_table_of_contents`、`wps_word_page_layout_and_watermark`、`wps_ppt_capture_slide_preview`、`wps_clear_audit_history`、`office_unlock_target`、`wps_unlock_target_document`。

最突出的是 **`wps_add_chart` / `excel_add_chart`**：说明承诺"创建与数据源动态绑定的原生矢量图表"，但 `dataRange`、`dataRanges`、`chartType` **一个都不是必填**——说明里那句"与 dataRanges 选其一传入"只是正文提示，schema 层没有任何强制。`wps_delete_chart` 同样零必填而带 `clearAll` 开关。

**影响**：AI 在缺少目标/数据源的状态下也能发出写调用。本次由于 P-02 的陈旧锁会先失败，反而"救"了它；一旦锁状态健康，这类调用就会真的落到工作簿上。

---

### P-13（低-中）营销化措辞挤占有用信息

全清单说明合计 **6549 字符（均值 71 字）**，其中每个 `excel_*` 固定追加 32 字样板（27 × 32 ≈ 864 字，占 13%），另有大量营销句，例如：

- `excel_delete_chart`："彻底解决旧图表无法清除、留下空白残缺边框的痛点。"
- `excel_add_chart`：`cellRange` 参数说明"由 Office 渲染引擎底层直接将图表咬死在该区域内，实现 **100% 完美**的行级对齐与等高排版，**杜绝像素漂移**！"
- `wps_ppt_add_business_cards`："现代化商业信息卡片"

**影响**：说明预算是有限的注意力资源。这些句子不提供可执行信息，却稀释了真正稀缺的内容（选型规则、参数约束、失败条件、单位与量纲）。

---

### P-14（低）错误提示不给出允许值，参数集不一致时尤其费轮次

实测错误文案：

| 触发 | 文案 |
|---|---|
| `excel_read_range {host:"foo"}` | `arguments.host: 不在允许值中`（不列允许值） |
| `excel_search_cells {host:"wps"}`（缺 query） | `arguments.query: 必填` |
| `excel_read_range {… , foo:1}` | `arguments.foo: 未知参数` |
| `wps_get_audit_record {auditId:"…", host:"wps"}` | `arguments.host: 未知参数` |

错误本身是机器可解析的（好），但都不给"正确形态"。结合已登记的 ISS-03（参数集不统一：`wps_rollback` 只收 `auditId`、`excel_save_workbook` 不收 `sheetName`），AI 按同一套约定批量调用时只能靠错误信息逐个试探。

---

## 六、反面对照：说明写得好的样例

为了说明"这不是做不到"，以下是本清单中说明质量的正面样例，可作为改写其余 80 余条的模板：

| 工具 | 好在哪 |
|---|---|
| `wps_execute_script`（266 字） | 唯一一条明确写了**何时用**（"专用工具未覆盖时使用本工具"）、**前置动作**（"先用 wps_inspect_api 或只读脚本检查目标 API"）、**注入变量**（`app`/`doc`/`wb`/`pres`/`wps`/`params`）、**返回约定**（"返回普通 JSON，勿返回宿主对象"） |
| `wps_ppt_manage_shapes_and_media`（135 字） | 写明单位（pt）、要求先读页面尺寸与现有形状、新增时必须显式给 `left/top/width/height`/`fontSize`、编辑后读回并预览 |
| `excel_format_cells`（136 字） | 参数级给出**建议值与禁用值**（"大标题设 15-18"、"严禁设 100pt 以上产生大空白框"、"严禁让日期显示为 46249 这类五位数字序列号"） |
| `excel_read_range` / `excel_get_range_styles` | 写明默认值与返回量控制（`includeFormulas 默认 true`、`maxCells 默认 100，最大 500`） |
| `wps_ppt_generate_deck` | 写明返回 `layoutWarnings` 且**必须逐页预览验证** |

**共性**：这 5 条都说清了"什么时候用 / 先做什么 / 单位与量纲 / 返回什么 / 怎么验证"——正是其余工具缺的四件事。

---

## 七、找不到、判断不了的部分

1. **70 个写入工具的真实行为未验证**。本次受"只读"约束未执行任何写入，除 `patch_cells`（有 34 条审计留痕）外，格式化、图表、透视表、冻结、筛选排序、数据验证、批注、查找替换、克隆表、保存，以及 Word/PPT 全部工具，只能给"声明可跑"的结论。**本报告中所有"未验"一律不等于"可用"**。
2. **Microsoft 通道整体判不了**。`isMsOfficeConnected: false`，`office_get_status` 返回 `runningComponents.excel: false`，word/ppt 报 `Application can't be found.`——该文案不区分"未安装"与"未运行"，因此 27 个 `excel_*` 的 `host=microsoft` 分支与 3 个 `office_*` 工具全部无法实测，且 `capabilities` 自述"尚未实机验收"。
3. **`update_chart` 的"必然失败"未实际触发**（触发即写文档）。判定依据是能力声明 + 路由表 + 宿主 RPC 分发的**静态**核对，属"高置信度推断"，不是实测复现。
4. **Word / PPT 的宿主支持范围无法逐工具判定**。`capabilities.hosts.wps.word` 与 `.ppt` 各只有一句自然语言（"有限支持：…"），没有像 Excel 那样给 `implemented` 数组；ISS-09 的 D9 已登记"capabilities 不能作为支持哪些工具的完整来源"。它建议补的 `bridge_get_tools` 工具在本次 91 项清单中**并不存在**。
5. **锁的来源查不到**。`wps_get_locked_status` 返回 `gatewayLocks: {}` 与 `addonStatus: null`，看不出 P-02 里那把陈旧锁是谁、何时设置的，也无法判断"锁是否仍会生效"。按已登记的 ISS-02，`hasOpenWorkbook` 字段语义本身也在待修状态。
6. **"说明够不够详细"含主观成分**。本报告只对**可验证项**给出证据（缺参数说明、缺 `type`、枚举与说明不齐、实测行为与说明不符、重复工具无选型指引）；§五 P-13 一类"信息密度"判断属评审意见，不是事实断言。
7. **未核对契约快照**。`docs/acceptance/2.1.0-p0p1/tools-snapshot.p5.json` 已存在，本次清单直接取自在运行的服务实例，两者是否逐字节一致未核对（不在本次范围）。

---

## 八、口径与复现

```sh
# 工具清单（正在运行的服务实例）
curl -s "http://127.0.0.1:19890/api/v1/mcp-tools" \
  -H "Authorization: Bearer $(cat ~/.wps-bridge/token)" -o tools.json

# 能力声明（implemented / declaredNotCallable / unimplementedOnHost / audit 覆盖）
curl -s "http://127.0.0.1:19890/api/v1/capabilities" \
  -H "Authorization: Bearer $(cat ~/.wps-bridge/token)"

# 单次调用
curl -s -X POST "http://127.0.0.1:19890/api/v1/tool/call" \
  -H "Authorization: Bearer $(cat ~/.wps-bridge/token)" \
  -H "Content-Type: application/json" \
  -d '{"name":"bridge_diagnose","arguments":{}}'
```

统计口径：工具数 91 = `GET /api/v1/mcp-tools` 返回数组长度；只读 21 = `annotations.readOnlyHint === true`；参数总数 592 = 各工具 `inputSchema.properties` 计数之和；说明字数 = `description` 的字符数（含空格与标点）。

源码侧判定引用：`src/bridge/tools/index.ts:22-23`（说明清洗）、`src/bridge/tools/excel.ts:20-35`（`excel_*` 派生）、`src/bridge/contracts/host-methods.ts:19-66`（路由表 / 宿主缺口 / 只读与重放集合）、`wps-addon/src/shared.js:188,223,259,285`（三组件粘性目标解析）。

**与其他验收文档的关系**：本文件只做"说明层可用性"盘点。能力声明与宿主映射差异见 [tool-inventory.md](../../tool-inventory.md) 与 [compatibility-diff.md](../../compatibility-diff.md)；已登记问题见 [issues.md](../../issues.md)。交叉关系：本文件的 P-02 与 ISS-02 同源（本条记的是 schema 声明可选 / 实际必填 / 自举路径被堵），P-07 与 ISS-01、P-14 与 ISS-03、P-10 与 ISS-09 分别互补但角度不同（说明层 vs 实现层）。**P-03、P-04、P-05、P-06、P-08、P-09、P-11、P-12、P-13 为本次新发现**，待并入台账 ISS-11 起。
