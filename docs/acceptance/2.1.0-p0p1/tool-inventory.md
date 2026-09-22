# 对外工具清点与宿主映射（改造计划 P0.3）

> **修订适用性**：本清单在 P1 修改过程中产出，正文引用的**行号对应 P0 基线源码**（HEAD `cac6d38` + 前序未提交修改），不对应 P1 之后的 `src/bridge/catalog.ts` / `src/bridge/office/adapter.ts`。
>
> 修正状态：**D1、D2、D3、D4、D5、D6、D7 已在 P1.5 按声明层修正**（对外工具名与 schema 零变化，见 [compatibility-diff.md](compatibility-diff.md) §3）；**D8、D9 仍待处理**，A1–A4 与未确认项亦未处理。
>
> **P3.2/P3.3 后行号全部失效**：`wps-addon/addon-core.js` 与 `office-addon/public/taskpane.js` 已改为由 `*/src/**` 构建生成的部署入口，本表中针对这两个文件的 `文件:行号` 引用（如 `addon-core.js:1397`）**不再对应当前产物**。函数名与行为未变，需要定位时请在 `wps-addon/src/**`、`office-addon/src/**` 中按函数名搜索；本表未重算行号，以免制造第二份易失效的映射。
>
> **评审补正（D6 收窄）**：D6 把 `wps_inspect_api` 也列为"被误标的只读工具"，这一条**不成立并已撤回**——它执行调用方给出的任意表达式，可能改文档，因此保持 `readOnlyHint: false`。实际更正的是另外 **4** 个工具，快照只读计数为 **17 → 21**（不是 22）。理由见 [compatibility-diff.md](compatibility-diff.md) §2.1。

## 一、数据来源与口径

| 项 | 值 | 依据 |
|---|---|---|
| 权威清单 | `docs/acceptance/2.1.0-p0p1/tools-snapshot.baseline.json` | 由 `scripts/snapshot-tools.ts` 从真实 `catalog.getTools()` / `capabilities()` 导出 |
| 导出命令 | `npm run snapshot:tools -- docs/acceptance/2.1.0-p0p1/tools-snapshot.baseline.json` | `package.json` 的 `snapshot:tools` 脚本 |
| 工具总数 | **91** | 快照 `toolCount` |
| 只读 | **17** | 快照 `readOnlyCount`（口径为 `annotations.readOnlyHint === true`） |
| 写入 | **74** | 快照 `writeCount`（`readOnlyHint !== true`，其中 4 个实际只读，见差异 D6） |

主表的“读写”列给的是**功能语义判断**（读 = 不改变文档/宿主状态；写 = 会改变文档、宿主或本地留痕），与快照的 `readOnlyHint` 口径不同：按语义统计为 **读 21 / 写 70**，比快照基线多出 4 个只读工具（`wps_get_locked_status`、`wps_word_read_document`、`wps_ppt_read_presentation`、`wps_ppt_get_slide_shapes`），差异原因见 D6。`wps_inspect_api` **不计入只读**（见 D6 补正），`wps_clear_audit_history`、`wps_lock_target_document` 等虽不改文档，但会改本地状态，仍按“写”计。`wps_clear_audit_history`、`wps_lock_target_document` 等虽不改文档，但会改本地状态，仍按“写”计。

**声明：本表由源码与真实 `getTools()` 快照推导，未做实机执行验证。** 宿主路由、审计覆盖来自静态源码事实；宿主内部逐项行为来自 `wps-addon/addon-core.js` 的实现阅读（已标行号）；仅凭 schema 形态推导的“部分成功语义”属推断，凡未逐函数核对的条目在表中以 `（推断）` 标注。

**91 的口径复算（源码事实）**：`catalog.getTools()` = 2 个诊断工具（`catalog.ts:37`）+ 27 个 `excel_*`（`catalog.ts:33-36`，由 `EXCEL_METHODS` 与已注册工具名求交集后改名）+ 59 个 gateway 原生 schema（`gateway.ts:1041-2763`：`wps_*` 表格 28 + Word 12 + PPT 9 + 脚本/反射 2 + 锁 3，`office_*` 5）+ 3 个审计工具（`catalog.ts:38-42`）= 91，与快照逐名一致。注意 `excel_*` 是**新增**的工具名，原 `wps_*` 同名工具同时保留。

## 二、宿主路由判定规则

| 标签 | 含义 | 判定依据（源码） |
|---|---|---|
| **双宿主** | 经 `callOffice()` 派发，实际宿主由请求上下文 `currentHost()` 决定：`wps` → WPS 加载项；`microsoft` → Office.js 加载项，失败且在 Windows 时回退原生驱动（仅限 `EXCEL_METHODS` 内的方法） | `office/adapter.ts:8-23`；`context.ts:5`（无上下文时默认 `wps`） |
| **WPS** | 同样经 `callOffice()`，但宿主被固定为 `wps`：只有 `excel_` 前缀且 `args.host === 'microsoft'` 才会置为 microsoft，兼容名 `wps_*` 永远走 WPS | `ws-server.ts:320-324`、`mcp-server.ts:32`、`catalog.ts:77` |
| **MS** | 不经 `callOffice()`，直接调用 `MsOfficeDriver`（macOS 走 JXA `osascript`，Windows 走 `powershell runner.ps1`） | `gateway.ts:147-198`、`office/ms-office-driver.ts:11-48` |
| **本地** | 无宿主调用：纯本地聚合/内存锁/本地审计库 | `catalog.ts:75-76`、`gateway.ts:84-118`、`gateway.ts:955-1004`、`audit-store.ts` |

补充事实：
- `wps_rollback` 是特例：它按审计记录里的 `host` 重新进入请求上下文（`gateway.ts:889`），所以可以为 microsoft 宿主记录的写入执行回滚，尽管工具名是 `wps_` 兼容名。
- `currentHost()` 只区分 `wps | microsoft`；`TargetLockStore` 的锁按 `session + host` 键控（`gateway.ts:26-63`），`MsOfficeDriver` 的锁只按 `session` 键控（`ms-office-driver.ts:7-10`）。
- 91 个工具中**没有**安装/部署/客户端配置类工具；安装部署走主进程与 HTTP 路由（如 `ws-server.ts:255` `/api/v1/office/install-addon`），不在本清单范围。

## 三、主表（91 行，顺序与快照一致）

分类列：`统一`=统一 `excel_*`；`兼容`=兼容 `wps_*`；`MS`=`Microsoft office_*`；`诊断`=`bridge_*`；`审计`=`wps_*audit*`。
“部分成功语义”中带行号者已核对宿主实现（源码事实），带 `（推断）`者仅由 schema 与网关分支推导。

| 工具名 | 分类 | 宿主路由 | 读写 | 部分成功语义 | 审计覆盖 |
|---|---|---|---|---|---|
| bridge_get_capabilities | 诊断 | 本地 | 读 | 否（纯本地聚合，不触达宿主） | 否 |
| bridge_diagnose | 诊断 | 本地 | 读 | 否（只读连接状态与组件清单） | 否 |
| excel_get_workspace_summary | 统一 | 双宿主 | 读 | 否 | 否 |
| excel_get_sheet_outline | 统一 | 双宿主 | 读 | 否 | 否 |
| excel_create_sheet | 统一 | 双宿主 | 写 | 否（单表创建，失败即无表；可用 excel_get_sheet_outline 读回） | 否 |
| excel_delete_sheet | 统一 | 双宿主 | 写 | 否（单表删除；连同内容一起生效、无审计、不可回滚） | 否 |
| excel_read_range | 统一 | 双宿主 | 读 | 否 | 否 |
| excel_get_range_styles | 统一 | 双宿主 | 读 | 否（`maxCells` 截断属结果不全，非写入部分生效） | 否 |
| excel_search_cells | 统一 | 双宿主 | 读 | 否（`maxResults` 截断属结果不全） | 否 |
| excel_patch_cells | 统一 | 双宿主 | 写 | 是（批量单元格矩阵，单次上限 10000 格 `addon-core.js:1397`；写入异常时宿主先恢复原值，恢复失败会明确抛“写入部分失败且恢复失败” `addon-core.js:1415-1421`；可用 excel_read_range 与 auditId 读回判定） | 是（仅值与公式，含修改前后快照；`gateway.ts:303-321`，全仓唯一 `addRecord` 调用点）。host=microsoft 时记录仍写入，但 Office.js 通道不返回快照（`normalizer.ts:243-258`、`office-addon/public/taskpane.js:750-756`），记录缺 before/after 快照，回滚会被拒绝（`gateway.ts:891`） |
| excel_format_cells | 统一 | 双宿主 | 写 | 是（同一区域多项样式一次下发；宿主返回无逐项计数 `addon-core.js:1562`，是否全部属性生效需 excel_get_range_styles 读回） | 否（样式不在 audit 覆盖内，与 capabilities.audit.uncovered 一致） |
| excel_add_conditional_formatting | 统一 | 双宿主 | 写 | 是（`clearExisting=true` 时先清后建，中途失败可能只剩清空效果；无专用条件格式读回工具） | 否 |
| excel_freeze_panes | 统一 | 双宿主 | 写 | 否（单次视图设置；无读回工具，需人工看界面）（推断） | 否 |
| excel_modify_rows_columns | 统一 | 双宿主 | 写 | 否（单次结构操作；插入/删除会移位既有数据且无审计，不可回滚） | 否 |
| excel_auto_fit_columns | 统一 | 双宿主 | 写 | 是（`columnRules` 逐条应用并返回 `results` 数组，可判定哪些规则生效 `addon-core.js:1963-2000`） | 否 |
| excel_capture_sheet_preview | 统一 | 双宿主 | 写 | 是（宿主预览失败会回退系统剪贴板取图；两条路径都失败仍返回 `success: true` 且无图，只有 message 提示 `gateway.ts:924-941`，此时不能据返回值判定） | 否 |
| excel_add_chart | 统一 | 双宿主 | 写 | 是（`replaceExisting` 默认 true，先删后建，中途失败可能只留下删除结果；可用 excel_get_charts 读回） | 否 |
| excel_get_charts | 统一 | 双宿主 | 读 | 否 | 否 |
| excel_update_chart | 统一 | 双宿主 | 写 | 是（多属性一次更新，无逐属性结果；host=wps 时必然失败，见差异 D7）。可用 excel_get_charts 读回 | 否 |
| excel_delete_chart | 统一 | 双宿主 | 写 | 是（`clearAll=true` 批量删除并返回 `deletedCount` `addon-core.js:2517-2527`；删除不可回滚，可用 excel_get_charts 读回） | 否 |
| excel_create_pivot_table | 统一 | 双宿主 | 写 | 是（行/列/数据字段逐个添加，仅返回字段计数 `addon-core.js:2607-2614`；无专用读回工具，失败判定未确认） | 否 |
| excel_set_filter_and_sort | 统一 | 双宿主 | 写 | 是（筛选与排序可视为两步，返回 `sortedRuleCount` `addon-core.js:2667-2672`；可用 excel_read_range 读回顺序） | 否 |
| excel_set_data_validation | 统一 | 双宿主 | 写 | 是（整区域批量下发，无逐格结果；无专用读回工具）（推断） | 否 |
| excel_manage_sheet | 统一 | 双宿主 | 写 | 否（单动作 rename/move/tab_color/protect/unprotect；保护密码无法读回，无审计） | 否 |
| excel_manage_rows_and_columns | 统一 | 双宿主 | 写 | 否（单次结构/尺寸操作；会移位数据，无审计） | 否 |
| excel_manage_cell_comments | 统一 | 双宿主 | 写 | 是（`clear_all` 批量删除全部批注且不可回滚 `addon-core.js:1754-1832`；同工具 `action=read` 可读回判定） | 否 |
| excel_find_and_replace | 统一 | 双宿主 | 写 | 是（逐格替换，返回 `replacedCount` 与逐项 `replaced` `addon-core.js:1852-1910`；可用 excel_search_cells 读回判定） | 否 |
| excel_duplicate_sheet | 统一 | 双宿主 | 写 | 是（含内容克隆，大表中断可能留下不完整副本；可用 excel_get_sheet_outline 读回） | 否 |
| excel_save_workbook | 统一 | 双宿主 | 写 | 是（整簿多表一次落盘；WPS 侧仅调用 `wb.Save()` 后即返回成功，未校验磁盘结果 `addon-core.js:1950-1959`；超时结果未知，需读回或查文件时间） | 否 |
| wps_execute_script | 兼容 | WPS | 写 | 是（任意多步原生脚本，中断即部分生效；返回值由脚本自身决定，需脚本内读回或后续工具读回） | 否（原生脚本不在审计覆盖内） |
| wps_inspect_api | 兼容 | WPS | 写 | 是（执行调用方给出的**任意表达式**，可取属性触发宿主求值、也可直接改文档，如 `Worksheets.Add()`；静态无法保证只读） | 否 |
| office_get_status | MS | MS | 读 | 否（只探测 Office 进程与已打开文档） | 否 |
| office_lock_target | MS | 本地（供 MS 通道使用的内存锁） | 写 | 否（只写内存锁表 `ms-office-driver.ts:9`，进程重启即丢） | 否 |
| office_unlock_target | MS | 本地（内存锁） | 写 | 否（只改内存锁表 `ms-office-driver.ts:10`） | 否 |
| office_execute_script | MS | MS | 写 | 是（任意多步脚本，中断即部分生效；结果由脚本返回，可再次执行只读脚本读回） | 否 |
| office_capture_slide_preview | MS | MS | 写 | 是（单页导出；未生成文件即报错，可判定该页成败 `ms-office-driver.ts:40-47`） | 否 |
| wps_lock_target_document | 兼容 | 本地（`TargetLockStore`，session+host 键控） | 写 | 否（只改内存锁，随进程结束丢失 `gateway.ts:37-43`） | 否 |
| wps_unlock_target_document | 兼容 | 本地（`TargetLockStore`） | 写 | 否（只改内存锁 `gateway.ts:45-51`） | 否 |
| wps_get_locked_status | 兼容 | 本地 | 读 | 否（只读内存锁；`addonStatus` 恒为 `null`，未反映加载项侧锁 `gateway.ts:111-118`） | 否 |
| wps_get_workspace_summary | 兼容 | WPS | 读 | 否 | 否 |
| wps_get_sheet_outline | 兼容 | WPS | 读 | 否 | 否 |
| wps_create_sheet | 兼容 | WPS | 写 | 否（单表创建；可用 wps_get_sheet_outline 读回） | 否 |
| wps_delete_sheet | 兼容 | WPS | 写 | 否（单表删除；连同内容生效、无审计、不可回滚） | 否 |
| wps_read_range | 兼容 | WPS | 读 | 否 | 否 |
| wps_get_range_styles | 兼容 | WPS | 读 | 否（`maxCells` 截断属结果不全） | 否 |
| wps_search_cells | 兼容 | WPS | 读 | 否（`maxResults` 截断属结果不全） | 否 |
| wps_patch_cells | 兼容 | WPS | 写 | 是（与 excel_patch_cells 同一条实现：10000 格上限、失败先恢复原值、恢复失败明确报“部分失败” `addon-core.js:1395-1421`；有 auditId 与前后快照可读回判定） | 是（仅值与公式，含修改前后快照；`gateway.ts:303-321`） |
| wps_format_cells | 兼容 | WPS | 写 | 是（多属性样式批量应用，无逐项计数 `addon-core.js:1562`；可用 wps_get_range_styles 读回） | 否 |
| wps_add_conditional_formatting | 兼容 | WPS | 写 | 是（`clearExisting=true` 先清后建；无专用读回工具） | 否 |
| wps_freeze_panes | 兼容 | WPS | 写 | 否（单次视图设置；无读回工具）（推断） | 否 |
| wps_modify_rows_columns | 兼容 | WPS | 写 | 否（单次结构操作；会移位既有数据且无审计） | 否 |
| wps_rollback | 兼容 | 双宿主（按审计记录的 host 重放 `gateway.ts:889`） | 写 | 是（先读回当前区域与 `afterSnapshot` 比对，不一致即拒绝覆盖 `gateway.ts:890-893`；回滚本身也是批量单元格写入，仍可能部分生效） | 部分（不新增记录，只把原记录置为 `rolled_back` `gateway.ts:897`、`audit-store.ts:159-165`；无回滚自身的独立留痕） |
| wps_auto_fit_columns | 兼容 | WPS | 写 | 是（逐条规则应用并返回 `results` `addon-core.js:1963-2000`） | 否 |
| wps_capture_sheet_preview | 兼容 | WPS | 写 | 是（剪贴板回退失败时仍返回 `success: true` 且无图 `gateway.ts:924-941`） | 否 |
| wps_add_chart | 兼容 | WPS | 写 | 是（`replaceExisting` 默认 true 先删后建；可用 wps_get_charts 读回） | 否 |
| wps_get_charts | 兼容 | WPS | 读 | 否 | 否 |
| wps_update_chart | 兼容 | WPS | 写 | 是（多属性更新无逐项结果；host=wps 必然失败，见差异 D7） | 否 |
| wps_delete_chart | 兼容 | WPS | 写 | 是（`clearAll=true` 返回 `deletedCount` `addon-core.js:2517-2527`；不可回滚） | 否 |
| wps_create_pivot_table | 兼容 | WPS | 写 | 是（逐字段添加，只返回字段计数 `addon-core.js:2607-2614`；无专用读回工具） | 否 |
| wps_set_filter_and_sort | 兼容 | WPS | 写 | 是（返回 `sortedRuleCount` `addon-core.js:2667-2672`；可用 wps_read_range 读回顺序） | 否 |
| wps_set_data_validation | 兼容 | WPS | 写 | 是（整区域批量下发；无专用读回工具）（推断） | 否 |
| wps_manage_sheet | 兼容 | WPS | 写 | 否（单动作；保护密码不可读回，无审计） | 否 |
| wps_manage_rows_and_columns | 兼容 | WPS | 写 | 否（单次结构/尺寸操作；会移位数据） | 否 |
| wps_manage_cell_comments | 兼容 | WPS | 写 | 是（`clear_all` 批量删除批注不可回滚 `addon-core.js:1754-1832`；`action=read` 可读回） | 否 |
| wps_find_and_replace | 兼容 | WPS | 写 | 是（逐格替换，返回 `replacedCount` 与逐项 `replaced` `addon-core.js:1852-1910`；可用 wps_search_cells 读回） | 否 |
| wps_duplicate_sheet | 兼容 | WPS | 写 | 是（内容克隆，中断可能留下不完整副本；可用 wps_get_sheet_outline 读回） | 否 |
| wps_save_workbook | 兼容 | WPS | 写 | 是（整簿落盘，仅 `wb.Save()` 后返回成功、未校验磁盘结果 `addon-core.js:1950-1959`；超时结果未知） | 否 |
| wps_word_create_document | 兼容 | WPS | 写 | 否（新建或按模板创建单文档，失败即无文档；可用 wps_word_read_document 读回） | 否 |
| wps_word_save_document | 兼容 | WPS | 写 | 是（原地保存/另存为/导出 PDF 三条路径，均只返回 success、不校验文件 `addon-core.js:2809-2849`；超时结果未知，需查文件） | 否 |
| wps_word_close_document | 兼容 | WPS | 写 | 是（`saveChanges=true` 时保存与关闭由宿主一步完成，返回不含保存结果 `addon-core.js:2852-2864`；关闭后无法读回，只能查文件时间） | 否 |
| wps_word_manage_content | 兼容 | WPS | 写 | 是（`delete_table`/`clear_all` 批量删除不可回滚；可用 wps_word_read_document 读回） | 否 |
| wps_word_read_document | 兼容 | WPS | 读 | 否（`maxParagraphs` 截断属结果不全） | 否 |
| wps_word_write_content | 兼容 | WPS | 写 | 是（单次插入，中断可能写入部分内容；可用 wps_word_read_document 读回）（推断） | 否 |
| wps_word_format_document | 兼容 | WPS | 写 | 是（多属性/多段落格式化，宿主返回无计数；可用 wps_word_read_document 的 includeFormatting 读回） | 否 |
| wps_word_insert_table_of_contents | 兼容 | WPS | 写 | 是（目录域生成后需宿主更新域，返回不含条目数 `addon-core.js:3305-3342`；可用 wps_word_read_document 读回） | 否 |
| wps_word_manage_table | 兼容 | WPS | 写 | 是（`write_matrix`/`update_data` 逐格写入，返回行列数 `addon-core.js:3385-3458`；可用 `action=inspect` 读回） | 否 |
| wps_word_review_and_comments | 兼容 | WPS | 写 | 是（`accept_all_revisions`/`reject_all_revisions` 为全体修订操作但返回无计数 `addon-core.js:3530-3534`；批注可 `list_comments` 读回，修订数无法读回判定） | 否 |
| wps_word_page_layout_and_watermark | 兼容 | WPS | 写 | 是（页眉/页脚/水印多项一次下发，无逐项结果 `gateway.ts:710-721`；无专用读回工具，需预览或人工确认）（推断） | 否 |
| wps_word_find_and_replace | 兼容 | WPS | 写 | 是（正文与表格逐处替换，返回 `matchCount` `addon-core.js:3616-3691`；可用 wps_word_read_document 读回） | 否 |
| wps_ppt_read_presentation | 兼容 | WPS | 读 | 否（`maxSlides` 截断属结果不全） | 否 |
| wps_ppt_get_slide_shapes | 兼容 | WPS | 读 | 否 | 否 |
| wps_ppt_generate_deck | 兼容 | WPS | 写 | 是（多页逐页生成，返回 `createdSlidesCount` 等于请求页数而非实测值 `addon-core.js:3985-4080`；中断会留下部分页，可用 wps_ppt_read_presentation 或逐页预览读回） | 否 |
| wps_ppt_manage_slides | 兼容 | WPS | 写 | 否（单动作 add/delete/move/duplicate/set_background；删除页不可回滚、无审计） | 否 |
| wps_ppt_manage_table | 兼容 | WPS | 写 | 是（`set_table_data` 整表写入、`style_table` 多项样式，逐格写入可能部分生效；可用 `action=read_table` 读回） | 否 |
| wps_ppt_add_business_cards | 兼容 | WPS | 写 | 是（`cards` 数组逐张绘制，无逐卡结果；可用 wps_ppt_get_slide_shapes 或预览读回）（推断） | 否 |
| wps_ppt_insert_native_chart | 兼容 | WPS | 写 | 是（多系列数据写入，无逐系列结果；可用 wps_ppt_capture_slide_preview 读回）（推断） | 否 |
| wps_ppt_manage_shapes_and_media | 兼容 | WPS | 写 | 是（`align_shapes` 按 `shapeIds` 逐个对齐，返回 `alignedCount` 等于请求数量而非实测成败 `addon-core.js:4485-4508`；可用 wps_ppt_get_slide_shapes 读回） | 否 |
| wps_ppt_capture_slide_preview | 兼容 | WPS | 写 | 是（未生成文件即报错，单页成败可判定 `gateway.ts:872-877`） | 否 |
| wps_get_audit_history | 审计 | 本地（audit-store 读） | 读 | 否（分页 + `view` 截断属结果不全，`gateway.ts:955-992`） | 否（本工具不写 audit-store；detail 视图单次上限 5 条 `gateway.ts:958`） |
| wps_get_audit_record | 审计 | 本地（audit-store 读） | 读 | 否 | 否（本工具不写 audit-store） |
| wps_clear_audit_history | 审计 | 本地（audit-store 写） | 写 | 是（整库清空且不可回滚 `audit-store.ts:144-147`、`gateway.ts:994-997`，无二次确认） | 否（本身是审计库操作，执行即删除全部留痕，不产生新记录） |

### 主表统计核对

| 维度 | 计数 | 说明 |
|---|---|---|
| 总计 | 91 | 与快照 `toolCount` 一致（逐名比对，无缺失、无重复、无多余，顺序与快照一致） |
| 分类：统一 / 兼容 / MS / 诊断 / 审计 | 27 / 54 / 5 / 2 / 3 | 快照前缀统计为 `excel_*` 27、`wps_*` 57、`office_*` 5、`bridge_*` 2；其中 `wps_*` 57 = 兼容 54 + 审计 3 |
| 路由：双宿主 / WPS / MS / 本地 | 28 / 50 / 3 / 10 | 双宿主 = 27 个 `excel_*` + `wps_rollback`；WPS = 50（表格 27 + Word 12 + PPT 9 + 脚本/反射 2）；MS = `office_get_status`、`office_execute_script`、`office_capture_slide_preview`；本地 = 3 个 `wps_*lock*` + 2 个诊断 + 3 个审计 + `office_lock_target`/`office_unlock_target`（这两个虽属 `office_*`，实现只写本地内存锁 `ms-office-driver.ts:7-10`） |
| 读写（主表语义口径） | 读 21 / 写 70 | 与快照基线 `readOnlyHint` 口径（读 17 / 写 74）的 4 个差额见 D6 |
| 审计覆盖：是 / 部分 / 否 | 2 / 1 / 86 | “是”仅 `wps_patch_cells`、`excel_patch_cells`（`gateway.ts:303` 是全仓唯一 `addRecord` 调用点）；“部分”为 `wps_rollback`（只改原记录状态）；审计读取工具不写 audit-store，按“否”计 |
| 明确存在部分成功语义的写工具 | 51 | 主表“部分成功语义”列以“是”开头者 |

## 四、与能力声明的差异

以下逐条为 `catalog.ts` 的 `capabilities()` 声明与代码事实的不一致，均附文件:行号。

| 编号 | 声明 | 代码事实 | 证据 |
|---|---|---|---|
| D1 | `microsoft.excel.validation: '全量 42 项结构化能力支持 (macOS / Windows 统一)'` | `EXCEL_METHODS` 实际 **30** 项，`getTools()` 实际产出的 `excel_*` 工具 **27** 个，`wps_*` 表格工具 28 个。全仓 grep “42 项”只出现在这一行声明里，没有任何可对应清单 | `catalog.ts:21` vs `catalog.ts:6-13`（30 项，快照 `excelMethods` 同为 30）、`catalog.ts:33-36`（交集改名后 27 个）、快照 `tools` 前缀统计 |
| D2 | 声明 WPS 与 Microsoft 的 excel `implemented` 都含 `clear_range`、`insert_dimension`、`rollback_cells` | 三者**都没有**可直接调用的工具：`getOpenAiTools()` 的 59 个 schema 中无 `wps_clear_range`/`wps_insert_dimension`；`wps_clear_range` 只有 switch 分支没有 schema，`insert_dimension` 连分支都没有；`catalog.ts:72-73` 会以“未知工具”直接拒绝。宿主侧能力其实存在（WPS `addon-core.js:556`、`addon-core.js:589`；Office.js `office-addon/public/taskpane.js:296`、`office-addon/public/taskpane.js:294`），即“宿主能做但对外工具不存在” | `catalog.ts:8-9`、`catalog.ts:20-21`、`gateway.ts:237-244`、`gateway.ts:1041-2763`、`catalog.ts:72-73` |
| D3 | `rollback_cells` 被计入两宿主 `implemented` | 对外只有兼容名 `wps_rollback`，不存在 `excel_rollback_cells`：`EXCEL_METHODS` 用 `rollback_cells` 匹配 `t.name.replace(/^wps_/,'')`，而注册名是 `wps_rollback`，永远匹配不上 | `catalog.ts:12`、`catalog.ts:33`、`gateway.ts:1484`、`gateway.ts:879` |
| D4 | `microsoft.excel.validation` 声称 macOS / Windows **统一** | 结构化 Microsoft 通道在 macOS 只有 Office.js 一条路：加载项报错时回退原生驱动被 `process.platform === 'win32'` 限制，macOS 直接抛出；Windows 回退也仅限 `EXCEL_METHODS` 内的方法。MCP 说明自身也写“Microsoft Excel 结构化工具目前面向 Windows COM，须经过实机验收”，与“macOS / Windows 统一”互相矛盾 | `office/adapter.ts:18-22`、`mcp-server.ts:10`、`catalog.ts:21` |
| D5 | `microsoft.word`、`microsoft.ppt` 标为 `'Office.js 原生通道'`，`microsoft.transport` 标为 `'Office.js Web Add-in / WebSocket'` | 对外不存在任何 microsoft 侧 Word/PPT 结构化工具：`host=microsoft` 仅对 `excel_` 前缀生效；Microsoft 的 Word/PPT 只能经 `office_execute_script`/`office_capture_slide_preview`，而这两者走的是原生驱动（JXA / PowerShell），不经 Office.js | `catalog.ts:21`、`mcp-server.ts:32`、`ws-server.ts:321`、`gateway.ts:178-198`、`office/ms-office-driver.ts:11-48` |
| D6 | `annotations.readOnlyHint` 表达工具的读写属性（快照 17 只读 / 74 写入） | `READ_TOOLS` 漏登记 **4** 个真实只读工具，导致 `wps_get_locked_status`、`wps_word_read_document`、`wps_ppt_read_presentation`、`wps_ppt_get_slide_shapes` 被标为可写；同一集合却把 `get_audit_history`/`get_audit_record` 记为只读。因此这 4 项应为只读。**初版 D6 另把 `wps_inspect_api` 也列为误标，评审补正已撤回**：该工具执行任意表达式、可能改文档，`readOnlyHint` 保持 false 是正确的 | `catalog.ts:14`、`catalog.ts:43`、快照中这 4 项 `readOnlyHint: false` |
| D7 | WPS 的 excel `implemented` 含 `update_chart`（`catalog.ts:10`、`catalog.ts:20`） | WPS 加载项**没有** `update_chart` 分支：全文件只有 `add_chart`/`get_charts`/`delete_chart`，未知方法会抛“未知的 RPC 方法”；Microsoft Office.js 侧反而有实现。即该能力声明对 WPS 宿主不成立 | `catalog.ts:10`、`catalog.ts:20`、`wps-addon/addon-core.js:595-603`、`wps-addon/addon-core.js:933-934`、`office-addon/public/taskpane.js:351` |
| D8 | `wps.ppt: '有限支持：基础页面、文字、形状、表格；复杂图表与母版需实机验证'` | WPS 侧实际暴露 `wps_ppt_insert_native_chart` 且有原生图表实现，声明未反映已存在的能力面（属宣传口径不准，非功能缺失） | `catalog.ts:20`、`gateway.ts:2589`、`wps-addon/addon-core.js:692` |
| D9 | `capabilities().hosts` 只给出 excel 的 `implemented` 数组，Word/PPT 仅一句描述 | Word/PPT 的工具覆盖范围无法从 capabilities 推得，客户端只能改查 `getTools()`；`bridge_get_capabilities` 因此不能作为“支持工具”的完整来源 | `catalog.ts:20-21`、`catalog.ts:37` |

**核对一致项（无差异）**：`audit.covered: ['patch_cells 值与公式']` 与代码一致——全仓只有 `gateway.ts:303` 一处调用 `auditStore.addRecord`；`audit.rollback` 描述的“检查当前值/公式与修改后快照一致，有后续修改则拒绝覆盖”与 `gateway.ts:890-893` 一致；`audit.uncovered` 列出的样式、图表、结构修改、Word/PPT、原生脚本确实没有审计写入。

### 同源附带发现（不属于 capabilities 声明，但为独立可复现事实）

| 编号 | 事实 | 证据 |
|---|---|---|
| A1 | 网关读取 `args.searchQuery` / `searchQueries` 做 Word 按文本格式化，但该工具 schema 标记 `additionalProperties: false` 且无这两个属性，参数会被入参校验拒绝 | `gateway.ts:654-655` vs `gateway.ts:2211-2251`、`catalog.ts:59` |
| A2 | 网关转发 `pageNumberFormat` / `watermarkColor`，但 `wps_word_page_layout_and_watermark` 的 schema 没有这两个属性，同样不可传 | `gateway.ts:715`、`gateway.ts:719` vs `gateway.ts:2346-2353` |
| A3 | 存在 6 个无 schema 的 switch 死分支，无法经 MCP/HTTP 调用（`catalog.ts:72-73` 先拦截）：`wps_get_style_token`、`wps_clear_range`、`wps_word_capture_preview`、`wps_ppt_add_chart`（与已暴露的 `wps_ppt_insert_native_chart` 共用分支的同义名）、`wps_reload_addon`、`wps_eval_code`。除该同义名外的 5 个对应的宿主 RPC 真实存在，属“宿主能做但对外工具不存在” | `gateway.ts:229`、`gateway.ts:237`、`gateway.ts:737`、`gateway.ts:825`、`gateway.ts:1006`、`gateway.ts:1010`；宿主侧 `wps-addon/addon-core.js:547`、`:556`、`:669`、`:773`、`:929` |
| A4 | 91 个工具中没有安装/部署类工具；“安装部署”能力经 HTTP 路由与主进程实现，不在本清单 | `ws-server.ts:255`、`ws-server.ts:305` |

## 五、未能确认项

以下无法从静态源码确定，明确记为**未确认**，不作推断：

1. **实机可用性未确认**：本清单未连接任何宿主、未执行任何工具；所有“宿主路由/审计/部分成功”结论都是静态推导。
2. **部署版本与源码是否一致未确认**：`wps-addon/addon-core.js`、`office-addon/public/taskpane.js` 的分支结论只在“运行时加载项 = 当前源码构建”时成立；版本一致性仅在连接时由加载项上报版本与 `VERSION` 比对（`ws-server.ts:365-380`），本次未做连接验证。
3. **“42 项”的来源未确认**：全仓仅 `catalog.ts:21` 一处出现该数字，无对应清单或计数逻辑可核对，无法判断它原本指哪一组能力。
4. **各工具的宿主超时阈值未确认**：多数网关分支不传 `timeout`（`office/adapter.ts:8`），实际超时由调用链默认值决定，未逐条核对。
5. **无专用读回工具的写操作失败判定未确认**：条件格式、数据验证、透视表、冻结窗格、Word 页眉页脚水印等没有对应的读取工具，“失败时能否读回判定”在表中标为推断，需实机确认。
6. **host=microsoft 下审计快照缺失的实际后果未确认**：源码显示 Office.js 通道不返回 before/after 快照（`normalizer.ts:243-258`、`office-addon/public/taskpane.js:750-756`），据此 `gateway.ts:891` 会拒绝回滚；本次未实机验证该记录的实际落库形态。
7. **Microsoft 侧 Word/PPT 能力边界未确认**：`office_execute_script` 的能力等于宿主 JXA/COM 脚本能力，没有静态清单可枚举；“Office.js 原生通道”的说法是否另有未暴露实现，未确认。
8. **8 行的“部分成功语义”未逐函数核对**：`excel_freeze_panes`、`excel_set_data_validation`、`wps_freeze_panes`、`wps_set_data_validation`、`wps_word_write_content`、`wps_word_page_layout_and_watermark`、`wps_ppt_add_business_cards`、`wps_ppt_insert_native_chart` 在表中标为 `（推断）`，仅由 schema 与网关分支推导，宿主实现未逐函数阅读。
9. **审计记录条数上限的副作用未确认**：`audit-store.ts:47-49` 只保留最新 500 条，超过后旧记录被丢弃；这对“回滚可用性”的实际影响未验证。
