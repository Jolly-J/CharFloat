# 07-tool-gap 中间产物（工作笔记）

产出文件：`docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/07-tool-gap.md`
快照时间：2026-09-22 16:1x（注意：同目录并行任务正在改 `wps-addon/src/*.js`，行号会漂）

## 一级结论

- 工具面 91 = 诊断2 + excel_*27 + wps_表格27 + word12 + ppt9 + 锁5 + ms原生3 + 脚本2 + 审计4
  （与 `docs/acceptance/2.1.0-p0p1/tools-snapshot.p5.json` 实测一致）
- 缺口 48 条：A 7 / B 28（B1 4 + B2 24）/ C 7 / D 4，跨类 2

## B1：宿主+网关都写好、只差 schema（零成本）

| 分支 | 网关 | 宿主 |
|---|---|---|
| wps_clear_range | gateway.ts:68 → gateway/excel.ts:59 | excel.js:104 |
| wps_get_style_token | gateway.ts:67 → gateway/excel.ts:51 | excel.js:7 |
| wps_word_capture_preview | gateway.ts:113 → gateway/word.ts | word.js:932 |
| wps_ppt_add_chart（别名，低价值） | gateway.ts:126 | ppt.js:461 |
| wps_reload_addon / wps_eval_code（低价值） | gateway.ts:130/133 | script.ts |

台账位置：tests/contract-consistency.test.ts:53-56（KNOWN_UNREGISTERED_BRANCHES）
路由表声明但不可调用：compatibility-diff.md:58 → clear_range / insert_dimension / rollback_cells

## C 类（只写不读）7 项

1. wps_add_conditional_formatting — excel.js:536（写完只回显 ruleType/address）
2. wps_freeze_panes — excel.js:611（无 SplitRow/FreezePanes 读回）
3. wps_set_data_validation — excel.js:1770
4. wps_set_filter_and_sort — excel.js:1642（只回显 appliedFilterRange）
5. wps_manage_sheet protect/tab_color — excel.js:1837
6. wps_word_page_layout_and_watermark — word.js:779
7. wps_create_pivot_table — excel.js:1554

正面对照：format_cells 的合并分支做了写后读回（excel.js:511-526），证明模式可行。
另外 wps_rollback 只覆盖 patch_cells（gateway/excel.ts:131 是唯一 addRecord 点）。

## 参数被实现忽略（§4）

- wps_auto_fit_columns 的 `address` → excel.js:944 未解构
- wps_capture_sheet_preview 的 `chartName`/`name` → excel.js:1086 未解构（MS 支持，chart.js:247）
- wps_add_chart 顶层 `left/top/width/height` + `cellRange/startCell/endCell` → excel.js:1113 只认 position.*
  （MS 侧 normalizer.ts:164-167 会补 left/top/width/height → 跨宿主不一致）
- wps_add_chart `chartType` bar/area/scatter → chartTypeMap（excel.js:1191）无键 → 静默降级簇状柱形
- wps_ppt_insert_native_chart 的 `hasLegend`/`showDataLabels` → ppt.js:462 未解构、473 未透传
- wps_ppt_manage_table 的 `zebra` → ppt.js:856 解构未用；`action:"set_table_data"` → 宿主 0 处实现，必抛错
- wps_duplicate_sheet 的 `position:"end"` → excel.js:911 只有 before/else
- wps_format_cells 的 `verticalAlignment: top|bottom` → excel.js:476 只处理 center
- wps_word_read_document 的 `includeFormatting` 默认 true 但实现是真值判断（word.js:175）；
  `maxParagraphs` 不影响大纲（word.js:152 硬编码 200）；scope=selection/tables 无实际分支
- wps_word_format_document：preset=gov_standard 时 margins 被忽略（word.js:345-355）；
  target=selection 时 alignment/行距/缩进全部忽略（word.js:451-466）；preset=academic 无分支
- 反向（宿主有 schema 没有）：word_format_document 的 searchQueries（word.js:379-421）

## host=microsoft 的错位（§2.4）

normalizer.ts 只适配 14 个方法，default（:204）原样透传。落到 MS 加载项后：
- set_data_validation：只认 params.rule（range.js:231）→ 先 clear 再啥都不设
- manage_sheet：只认 copy/duplicate/rename/hide/show/color（sheets.js:56-77）→ move/protect/unprotect 静默 success
- manage_sheet 反向：宿主有 hide/show，schema enum 没有
- set_filter_and_sort：要 params.address（filter-sort.js:7），schema 给 range → 报错
- create_pivot_table：要 sourceAddress/destinationAddress（pivot.js:7-8），schema 给 sourceRange/destCell
- manage_rows_and_columns：要 dimension（range.js:152），schema 给 targetType → 默认 rows → **想插列却插行**；
  hide/unhide/set_size 静默 success
- update_chart（host=wps）：dispatch.js 无分支，必失败（compatibility-diff.md:73）

## 危险别名（§2.5）

- office-addon/src/rpc.js:144-146：list_conditional_formats / update_conditional_format → handleAddConditionalFormatting
  → 「读条件格式」会新增一条 cell_value 规则（format.js:153-176）
- office-addon/src/rpc.js:193-195：list_comments / update_comment → handleManageComments，action 默认 "add"（comment.js:7）
  → 「列批注」会在 A1（comment.js:10）插一条空批注

## 通道边界

- office-addon 只有 src/excel/**；rpc.js:200 default 抛「暂未映射」
- errors.ts:107：不在 EXCEL_METHODS 内的方法不允许回退原生通道
  → 21 个 wps_word_*/wps_ppt_* 工具在 host=microsoft 时是死路
