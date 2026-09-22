# 12 桥接侧修复记录（ISS-46/93/97/77/02/48/49/50/51/59/90/16/03）

> 写区：`src/bridge/**`，排除 `src/bridge/tools/definitions/**` 与 `src/bridge/tools/excel.ts`
> （这两处由其他任务负责；本任务对它们只读）。
> 不在写区内的目标（加载项上报构建指纹、Python skill 的错误文案）改为等价手段或在文末单列。
> 状态：✅ 完成 / ⚠️ 只改文案或只能部分完成 / ❌ 修不动

## 逐条结果

| 编号 | 文件 | 改了什么 | 状态 |
|---|---|---|---|
| ISS-46 | `gateway/audit.ts`、`audit-store.ts`、`gateway.ts` | 回滚判定从"内容相等"升级为**四道判据**：① 同区域没有更晚的未回滚记录（区域重叠判定）；② 记录之后没有发生过未审计的"身份类"操作（脚本/清表/工作表增删改/行列增删/批量替换，`AuditStore.markUntrackedMutation` 按工作簿登记时间戳）；③ 目标工作表仍存在（宿主报"找不到工作表"即拒绝）；④ 内容仍等于记录里的修改后快照。任一不满足即拒绝，并在错误里给出"用 `wps_get_audit_record` 取回旧快照 + `patch_cells` 显式写回"的替代路径 | ✅ |
| ISS-93 | `office/normalizer.ts` | 补字段映射与**只做归一、不重复实现**：`create/delete/duplicate_sheet`、`manage_sheet`（补 `newName`/`tabColor` 别名）、`manage_rows_and_columns`（`targetType`→`dimension`）、`set_filter_and_sort`（`range`→`address`、`sortRules[0]`→`sortColumn`(0 基)+`sortOrder`）、`create_pivot_table`（`sourceRange`/`destCell`→`sourceAddress`/`destinationAddress`）、`manage_cell_comments`（`read`→`list`）、`find_and_replace`（`searchQuery`→`text`、`searchRange`→`address`、`maxResults` 截断）、`capture_sheet_preview`、`add_chart`（`dataRange`→`sourceAddress`、定位字段归一）、`get_charts`（宿主未过滤时桥接侧补选择器语义）、`get_range_styles`（叠加 WPS 形状 `styles`）。仍被宿主忽略、因而**显式报错或如实回报**的：透视表字段编排/跨表、批注 `clear_all`/`author`/无 id 删除、只切筛选（无 sortRules 会被宿主默认按第 1 列排序）、多级排序、空搜索串；`duplicate_sheet.position≠after` 与 `enableAutoFilter` 记入 `unsupportedFields`（照做 + 如实回报） | ✅ |
| ISS-99 | `office/normalizer.ts` | 响应转换从"白名单式重建"改为"**先铺开宿主原始响应**再覆盖统一字段"：`warnings` / `dataRangesApplied` / `seriesCount` / `yAxisApplied` / `totalCount` / `filtered` / `selector` / `changed` / `truncated` / `kind` / `renderedBy` 等宿主新增的"如实提示"不再被丢掉；`get_charts` 在宿主已自行过滤时不再重复过滤（避免把 `totalCount` 语义搞乱） | ✅ |
| ISS-97 | `contracts/host-methods.ts`、`office/adapter.ts` | 新增 `COM_IMPLEMENTATION_GAPS=['update_chart','save_workbook']` 与 `COM_EXCEL_METHODS`（30−2=28），adapter 的回退白名单改用它；不再把整张路由表当白名单 | ✅ |
| ISS-77 + ISS-02 | `gateway/excel.ts`、`gateway/lock.ts`、`gateway/types.ts`、`gateway.ts` | `get_workspace_summary` 拆字段：`hasOpenWorkbook` = 宿主真实"有没有打开工作簿"，`openWorkbooks` = 真实打开列表，`lockTarget {name,source,exists,missing}` = 锁目标是否存在；目标未命中时**再查一次不带目标**的摘要（这是唯一能列出已打开文稿的路径，解开自举死锁），并返回 `targetResolution` 回退顺序。`gateway.ts` 记录目标来源（request / session-lock / none）供说明使用。`wps_get_locked_status` 的 `addonStatus` 不再是 `null`：返回真实 `openWorkbooks` + 锁目标是否存在的比对 | ✅ |
| ISS-48 | `gateway/excel.ts`、`gateway/audit.ts`、`audit-store.ts`、`types.ts`、`tools/audit.ts`、`catalog.ts` | 12 类写操作补审计记录（format/conditional_format/freeze_panes/行列结构/图表/数据有效性/筛选排序/批注/工作表结构/批量替换/透视表/保存），返回体带 `auditId`+`rollbackable:false`+`rollbackScope`+说明；`AuditRecord` 增加 `rollbackable`、`sessionId`，`beforeSnapshot` 变为可选；`wps_rollback` 对不可回滚记录明确拒绝；`get_audit_history` 的 summary 带 `rollbackable`，能力里给出"可回滚/只留痕"清单 | ✅ |
| ISS-49 | `mcp-server.ts`、`audit-store.ts`、`gateway/excel.ts` | HTTP `/mcp` 会话的审计 `clientName` 改为从 MCP `initialize` 的 `clientInfo` 取名（`name version`）；显式指定的仍是显式值（stdio 通道保持 `stdio MCP` 不变，既有测试契约不破）。审计记录新增 `sessionId`，同名客户端也能区分 | ✅ |
| ISS-50 | `service-client.ts`、`cli.ts` | 新增 `serviceAddress()`/`serviceRecovery()`；`serviceRequest` 的网络层失败统一为"服务名 + 地址 + 恢复动作 + 原始错误"，不再抛裸 `fetch failed`；`probeService` 返回 `service/address/recovery/error` 结构化字段；`--doctor` 与 `--status` 分离输出（结论 / 可执行动作 / `rawError` 分开，未就绪时退出码 1） | ✅ Node 侧 |
| ISS-51 | `ws-server.ts` | `/mcp` 会话记录最后活动时间并按空闲阈值回收（默认 15 分钟，`WPS_BRIDGE_MCP_IDLE_MS` 可调；上限 `WPS_BRIDGE_MCP_MAX_SESSIONS`，默认 64）；建新会话前先扫一次空闲；429 文案给出：自动回收时间、`DELETE /mcp` 立即释放、环境变量调整、重启后台；`/api/v1/status` 暴露 `mcpSessions` | ✅ |
| ISS-59 | `build-fingerprint.ts`（新增）、`ws-server.ts`、`catalog.ts` | 新增构建指纹采集：桥接入口产物、WPS 加载项随包产物与**每个已部署副本**、Office.js 任务窗格产物的 sha256/mtime/size/路径；`deployedMatchesResource` 判定"部署副本是否等于当前构建"；连接上报版本与当前版本的比对放进 `components.*.versionStatus`。能力（`bridge_get_capabilities`）与 `/api/v1/status`、`bridge_diagnose` 均可查 | ⚠️ 桥接侧可查；**加载项上报自身指纹未做**（属 `wps-addon/**` 写区） |
| ISS-90 | `ws-server.ts`、`catalog.ts` | WS 关闭时按关闭码/终止原因分类：`replaced` / `clean-close` / `heartbeat-timeout` / `abnormal-close`；后两者标记 `suspectCrash` 并写入组件 `lastDisconnect` 与顶层 `suspectedHostCrash`，附 5 步恢复顺序指引；重连后清顶层信号但保留组件上的掉线历史；`bridge_diagnose` 增加 `host-crash` 检查项 | ✅ |
| ISS-16 | `gateway/excel.ts` | `save_workbook` 返回体新增 `scope:"workbook"` + `warning`（会把整个工作簿内存状态落盘，含其他会话/任务在途结果）；能力 `limitations` 同步 | ✅（说明落在返回体，工具描述在 definitions 写区外） |
| ISS-03 | `catalog.ts` | 校验错误列出**允许的参数集**（未知参数/枚举/类型错误都带期望值）；对 `excel_*`/`wps_*` 统一容忍 `host`/`workbookName`/`sheetName`：schema 不认时忽略并在返回体回报 `ignoredParams` + `ignoredParamsNote`，不再 422、也不静默吞掉 | ✅ |

## 改动文件清单

新增：`src/bridge/build-fingerprint.ts`
修改：`catalog.ts`、`ws-server.ts`、`mcp-server.ts`、`audit-store.ts`、`types.ts`、`service-client.ts`、`cli.ts`、`gateway.ts`、`gateway/{excel,audit,lock,types}.ts`、`office/{normalizer,adapter}.ts`、`contracts/host-methods.ts`、`tools/{audit,diagnostics}.ts`

未改（写区外）：`tools/definitions/**`、`tools/excel.ts`、`wps-addon/**`、`office-addon/**`、`skills/**`、`tests/**`。

## 验证（本轮实际执行）

| 检查 | 结果 |
|---|---|
| `npm run typecheck` | 0 错误 |
| `npm test` | **86/86 通过**（改动前后同为 86，未新增测试文件——`tests/**` 不在写区；构建 `dist/` 前后各跑一次均通过） |
| `npm run build:main` | 成功（`dist/bridge/cli.cjs` 1.91 MB） |
| `npm run check:agents` | 通过（新增 `build-fingerprint.ts` 已登记进 `src/bridge/AGENTS.md`，另补 3 条避坑条目） |
| `npm run snapshot:tools` 对比 `tools-snapshot.p6.json` | 工具**名称/顺序/数量(91)/schema 哈希逐项不变**；仅 3 处描述文本更新（`wps_get_audit_history`、`bridge_get_capabilities`、`bridge_diagnose`）+ `capabilities` 新增字段（`build`/`targetResolution`/`sessionLifecycle`/`audit.loggedNotRollbackable`/`serviceErrors`） |
| 临时脚本（`.scratch/`，验证后已删除） | ISS-46 复现用例被拒、ISS-48 留痕与拒绝回滚、ISS-49 两个客户端可区分、ISS-51 空闲回收与 429 文案、ISS-90 疑似崩溃信号、ISS-02 陈旧锁 + 真实打开列表、ISS-93 的 27 个 Microsoft 方法空载荷不抛错且宿主已支持的能力不再被阻断、ISS-97 白名单 28/30、ISS-99 宿主新字段全部透传 |

### 模拟验证读数（临时脚本，非真实宿主）

- **ISS-46**：patch → `wps_execute_script`（模拟"删表重建"）→ 旧 `auditId` 回滚 → **拒绝**："本记录…之后，工作簿 [fault.xlsx] 上发生过未经快照记录的操作：wps_execute_script…"；同区域更晚的 patch → 旧记录回滚**拒绝**（列出更新的记录 id）；最新记录回滚成功，返回 `identityChecks` 四项。
- **ISS-48**：`excel_format_cells` 返回 `auditId` + `rollbackable:false`；用该 id 回滚 → 拒绝（"只登记了操作事实，没有值/公式快照"）；历史 summary 里 `update_values:true`、`format:false`。
- **ISS-49**：两个 HTTP MCP 客户端（`client-one`/`client-two`）的审计 `clientName` 分别为 `client-one 9`、`client-two 9`，且 `sessionId` 不同。
- **ISS-51**：`WPS_BRIDGE_MCP_IDLE_MS=400` 时会话在 1.5s 内被回收（active 1→0，recycled 1）；超出上限返回 429 且含 `release` 数组。
- **ISS-90**：宿主 socket 被 `terminate()`（无关闭帧）→ `components.excel.lastDisconnect = {code:1006, kind:'abnormal-close', suspectCrash:true, guidance:"…5 步恢复顺序…"}`，顶层 `suspectedHostCrash` 同步。
- **ISS-02**：锁目标 `old.xlsx`（未打开）+ 实际打开 A/B 两个工作簿 → `hasOpenWorkbook:true`、`lockTarget.missing:true`、`openWorkbooks:[A.xlsx,B.xlsx]`、message 指路；`wps_get_locked_status.addonStatus.openWorkbooks` 返回同样列表。
- **ISS-93**：27 个 Microsoft 已实现方法在空载荷下 `normalizeOfficeRequest/Response` 均不抛错（保持既有契约测试成立）；`manage_sheet` 的 `tab_color` 正确改写为 `action:'color'+tabColor`，move/protect/unprotect 抛错；`manage_rows_and_columns` 的 `targetType:'column'` → `dimension:'columns'`，hide/set_size 抛错。
- **ISS-97**：`COM_EXCEL_METHODS.length === 28`，不含 `update_chart` / `save_workbook`。
- **ISS-59**：指纹可用，且**当场就测出不一致**：已部署的 WPS 加载项副本 `d4c01c7b…`（208,864 B，08:23）≠ 当前构建 `71e72af4…`（257,195 B，09:00），`deployedMatchesResource:false`。

## 未完成 / 未实测（不要当成已验证）

0. **与并行任务的接缝（重要）**：本任务进行中，Microsoft 侧（`office-addon/**`，写区外）另有一次提交 `68ab9c1`，
   把 `set_data_validation` / `manage_sheet` / `manage_rows_and_columns` / `get_charts` / `add_chart` /
   `find_and_replace` 的字段映射**在宿主侧补齐**，并新增台账 ISS-99（响应转换丢字段）交给桥接侧。
   本任务据此**回收了早先"映射不了就抛错"的部分实现**：凡宿主已实现并读回校验的（含 `manage_sheet` 的
   move、行列 hide/unhide/set_size、`add_chart` 的 dataRanges/yAxis/seriesSettings/smoothLine/hasDataLabels、
   `set_data_validation` 的提示/报错文案），桥接侧只做字段归一，不再阻断——否则会挡住宿主已经能做的能力。
   ISS-99 已按上表修好。**并发的 `git add -A` 提交（3591554）曾把本任务的部分中间改动混进无关提交**，
   本任务自身未执行任何 git 操作，最终状态以工作区为准（已 typecheck/test/build 全绿）。
1. **ISS-59 的加载项侧上报**：加载项目前只上报版本号（`2.1.0`），桥接侧无法证明"宿主进程里加载的到底是哪一份字节"。本任务只做到"磁盘/部署副本指纹 + 版本比对"。要真正回答"运行中的是哪一版"，需在 `wps-addon/**` / `office-addon/**`（写区外）的 register 包里加构建指纹。
2. **ISS-46 的残余缺口**：若有人**绕过桥接**（直接手工在宿主里）删表重建并写回完全相同的内容，桥接侧没有任何信号（身份信息只在宿主里）。彻底解决需要宿主提供工作表身份（创建时间/内部 ID）。已在错误文案与文档里说明判据来源。
3. **Microsoft 侧全部改动均未实机验收**（本候选版本 MS 通道本就零实测）：桥接侧改动只是参数归一与响应透传；`set_data_validation` 的规则构造、行列 hide/set_size、`add_chart` 的轴/标签等**由宿主侧实现**（对方提交自称真机 10/10，本任务未复验）。
4. **ISS-50 的 Python skill 侧**（`skills/**`，写区外）：`bridge_client.py` 的 `--status/--doctor` 仍会输出原始 `urlopen` 错误，未改。
5. 工具的**描述文本**（`tools/definitions/**`，写区外）：`wps_rollback` / `excel_save_workbook` / `wps_get_workspace_summary` / `excel_*`（Microsoft 通道能力）的说明未改，相关补充改由返回体字段与 `capabilities()` 承载（`audit.rollback`、`targetResolution`、`save_workbook.warning`）。
6. 未启动/重启任何后台服务（按要求）；上面全部为进程内模拟与静态验证，**没有真实 WPS/Excel 宿主验证**。

## 分层与契约影响

- 未修改 `ToolService` 接口（`list/capabilities/execute` 签名不变），`tests/session-isolation.test.ts` 通过。
- 协议层（`ws-server.ts` / `mcp-server.ts`）未新增对 `catalog` 的静态引用；`ws-server.ts` 新增依赖是 `build-fingerprint.ts`（只读文件指纹，不引用执行层）。契约层 `contracts/host-methods.ts` 仍为纯定义（新增两个常量）。
- 新增的**内部**参数：`UniversalGateway.executeTool(name, args, clientName, meta)` 第 4 个可选参数（catalog → gateway 传"被忽略的参数名"），对既有 3 参调用完全兼容；`GatewayContext` 新增 `targetSource`、`ignoredParams` 两个可选字段。
- `AuditRecord`：`beforeSnapshot` 由必填改为可选，新增 `sessionId`、`rollbackable`；`actionType` 联合类型扩展。持久化文件 `audit_history.json` 向后兼容（旧记录无新字段时按"可回滚"处理）。
