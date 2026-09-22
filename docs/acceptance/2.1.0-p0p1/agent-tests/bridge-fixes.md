# 桥接层四项修复报告（CAP-51 / CAP-50 / ISS-126 / CAP-42）

- 候选标识：`2.1.0-p0p1`
- 执行者：teammate `bridge-fixer`（共享任务 `task-5`）
- 改动范围：`src/bridge/office/**`、`src/bridge/errors.ts`、`src/bridge/gateway/lock.ts`，外加 Lead 明确授权的
  `src/bridge/contracts/host-methods.ts`（CAP-51）与 `src/bridge/gateway.ts`（ISS-126，仅锁相关行）
- **验证层级（先说清楚）**：本轮**只做源码改动 + 静态/单测验证**。按任务硬约束，**没有**执行
  `npm run build:*` / `setup`、**没有**重启服务、**没有**调用任何真机工具（WPS 独占资源由 Lead 串行处理）。
  因此下面所有"已修"都指**代码 + 不依赖产物的单测**这一层；**真机行为一律标"未验证"**。
- 未运行：`tests/session-isolation.test.ts`（拉起 `dist/` 旧产物，且任务说明其失败属预期）。
  `tests/ppt-layout.test.ts` 的「生成物与源码一致」当前失败，原因是**其他 teammate 改了 `wps-addon/src/**` 尚未重建**（见 §5），与本轮改动无关。

---

## 0. 结论摘要

| 项 | 判定 | 改了什么 | 验证到哪一层 |
|---|---|---|---|
| **CAP-51** COM 回退白名单 | **已修** | 白名单改为由「脚本实际分支」派生：**48 → 28**（路由表 50 项 ∩ `excel.ps1` 实际实现 28 项）；缺口清单不再手写 | typecheck + 新增同源守卫测试（解析 `excel.ps1` 真实分支比对）+ 全部非产物单测 |
| **CAP-50** 跨宿主字段错位 | **已修 5 处**（3 处破坏性、2 处静默降级）；另有 2 处**已转 task-6** | M1 数据有效性两处破坏路径 + M2 空搜索串"三通道"护栏 + 桥接层禁令（禁止把安全护栏拒绝转手原生通道）+ 边框色丢失 | typecheck + 新增 4 条回归单测（含"宿主调用次数为 0"断言）+ 源码逐条核对；**未真机** |
| **ISS-126** 残留目标锁 | **已修** | 锁目标在宿主里不存在且**由会话锁注入**时：自动解除 + 报错改写为可操作提示；状态查询补 `staleLockTargets` | typecheck + 新增回归单测（真跑 `executeTool`，桩宿主注入宿主原文错误）+ 非产物单测 |
| **CAP-42** `word_capture_preview` | **无需改：HEAD 已注册齐全** | 本轮只做核验，未改代码 | 静态核验 5 环（schema / 网关分支 / 分类 / 宿主 RPC / 已部署 `addon-core.js`）+ 台账测试通过；**未真机调用** |

---

## 1. CAP-51：COM 回退白名单按**实际覆盖**生成（正确性）

### 根因

`office/adapter.ts` 的 `fallbackToNative()` 用 `COM_EXCEL_METHODS`（`contracts/host-methods.ts`）作为
"Windows 原生 COM 能接管哪些方法"的白名单。该常量当时是：

```ts
COM_IMPLEMENTATION_GAPS = ['update_chart', 'save_workbook'];          // 手写
COM_EXCEL_METHODS = EXCEL_METHODS.filter(m => !GAPS.includes(m));     // 按路由表做减法
```

ISS-97 修它时路由表还是 **30** 项，减 2 = 28，与实际相符；此后路由表涨到 **50** 项
（CAP-15~21 的打印/导出/形状/超链接等），减法结果变成 **48 项** —— 而 `resources/office/excel.ps1`
的 `Invoke-ExcelTool` 只实现 **28** 个方法名。**20 个方法被声明为"可回退"，实际回退后只会抛
`Unsupported Excel method`**，白名单再次过度声明（同一缺陷的第二次发生，原因就是"手写清单 + 派生减法"）。

核对方式：逐个读 `excel.ps1` 的分支——前置 `if($method -eq ...)` 2 个（`get_workspace_summary`、`create_sheet`）
＋ `switch($method)` 23 个标签（其中 2 个是 `{$_ -in @(...)}` 多别名分支），共 28 个方法名。

### 改法

`src/bridge/contracts/host-methods.ts`：

- 新增 **`COM_IMPLEMENTED_METHODS`** = `excel.ps1` 实际实现集（28 项，唯一事实来源，逐项对应脚本分支）；
- `COM_EXCEL_METHODS = EXCEL_METHODS.filter(m => COM_IMPLEMENTED_METHODS.includes(m))` → **28**；
- `COM_IMPLEMENTATION_GAPS = EXCEL_METHODS.filter(m => !COM_IMPLEMENTED_METHODS.includes(m))` → **22**（派生，不再手写）。

`src/bridge/office/adapter.ts`：更新注释为 `28/50` 并写明派生关系。

`tests/contract-consistency.test.ts`：新增守卫「CAP-51：COM 回退白名单与 excel.ps1 实际分支同源」——
**解析 `resources/office/excel.ps1` 的真实分支**再与清单比对（不是同源清单互相比对），并断言
`COM_EXCEL_METHODS.length === 28`、`update_chart`/`save_workbook` 不在其中、缺口清单等于派生结果。
另加"解析到的方法数 < 20 即判解析器失效"的防空跑断言。

> 补充事实（与 Lead 的判断有出入，需记录）：`tests/contracts-boundary.test.ts` **并没有** COM 白名单断言
> （它只断言 `EXCEL_METHODS.length === 50` 与 `HOST_IMPLEMENTATION_GAPS`），所以本次修 CAP-51 没有触发它失败。
> 真正缺的是"白名单 vs 宿主脚本"的独立校验，已按上面的方式补进 `contract-consistency.test.ts`。

### 验证到什么程度

- `npm run typecheck` 通过；
- 守卫测试通过（解析到 28 个分支，与清单逐项一致）；
- `COM_EXCEL_METHODS=28`、`COM_IMPLEMENTATION_GAPS=22`（22 项 = `update_chart`、`save_workbook`、`get_style_token`、
  打印/导出系列、形状系列等，见 `COM_IMPLEMENTATION_GAPS` 输出）；
- **未验证**：Windows 真机上的回退行为（本机非 Windows，且不允许真机操作）。静态结论成立，不冒充实机通过。

### 遗留口径（Lead 已裁决，不再作为待办）

`routeOfficeFailure` 的拒绝文案让调用方"用 `bridge_get_capabilities` 查询能力清单"，但
`capabilities()` **没有**任何 COM 回退覆盖字段（`src/bridge/catalog.ts`），即：报错把使用者指向一个查不到答案的地方。
**Lead 裁决：不修**——加 `comFallback: { implemented, gaps }` 会动能力元数据契约（`check:claims` 与工具快照都会受影响），
记为**已知缺口**（下一个会话若要改，需同时处理 `catalog.ts` + `check-claims` + 快照三处）。
这是本轮唯一"文案指向查不到的地方"的残留，不影响 CAP-51 的白名单正确性本身。

---

## 2. CAP-50：M1–M8 跨宿主字段错位（正确性，改错会破坏真实数据）

### 2.1 逐条核对结果（M1–M8）

| 编号 | 方法 | 桥接层（本轮结论） | 宿主侧（源码核对） |
|---|---|---|---|
| M1 | `set_data_validation` | **本轮修 2 处**（见 2.2 / 2.3） | Office.js：在 `clear()` **之前**构造规则并抛错（安全）；COM：`excel.ps1` 在 `Delete()` 之前抛错（安全）；**WPS 宿主：先 `Delete()` 再抛错（破坏性，已转 task-6）** |
| M2 | `find_and_replace` / `search_cells` | **本轮修 3 处**（见 2.4 / 2.5） | Office.js 与 WPS 宿主都会拒绝 undefined/null，但**都不拦空串**（WPS `includes('')` 恒真、按空正则逐格替换；COM `IndexOf('')` 同理） |
| M3 | `create_pivot_table` | 已具备（`sourceAddress`/`destinationAddress` 归一 + 字段编排显式报错）。**未改** | Office.js 有 `create_pivot_table` 分支（`rpc.js`）；字段编排在 MS 侧不实现 → 桥接层直接拒绝（如实，不静默） |
| M4 | `set_filter_and_sort` | 已具备（`range`→`address`、`sortRules`→`sortColumn/sortOrder`、`enableAutoFilter` 无法下发时记入 `unsupportedFields`；多级排序与"只开筛选"显式报错）。**未改** | Office.js `filter-sort.js` 按新字段实现 |
| M5 | `manage_cell_comments` | 已具备（`clear_all`/无 id 的 `delete`/`author` 三种执行前拒绝，`read`→`list`）。**未改** | Office.js `comment.js` |
| M6 | `manage_sheet` | 已具备（`newSheetName`→`newName`、`color`→`tabColor` 别名补齐，保持两套语义都可落位）。**未改** | Office.js `sheets.js`：`move`/`tab_color` 已实现且**读回**；`protect`/`unprotect` 显式报错（MS 桌面不可靠）——比"静默 no-op"好，如实 |
| M7 | `get_charts` | 已具备（宿主已过滤时原样透传 `totalCount`；宿主未过滤时桥接侧按 `shapeName/chartIndex/chartTitle` 过滤，并标 `filteredFrom`）。**未改** | Office.js 自行过滤 |
| M8 | `add_chart` | 已具备（数据源/定位字段归一；`dataRanges` 多段由宿主如实 warnings）。**未改** | Office.js `chart.js`：`yAxis`/`seriesSettings`/`smoothLine`/`hasDataLabels`/`dataRanges` 逐项实现并读回，未生效写 warnings |

> 口径提醒：M1–M8 是"我方跨宿主不一致"，不是宿主能力缺陷。上表"未改"的六项是**本轮核对后确认已修且桥接层无需再动**，
> 不是"仍未修"。

### 2.2 【破坏性】数据有效性：`action:'read'` 在 MS 通道会变成"写"

- **改前会发生什么破坏**：schema 承诺 `action:'read'` 只读回、不修改。但 Office.js 的
  `handleSetDataValidation` 与原生 COM 的 `set_data_validation` **都没有 action 分支**：
  - Office.js：照 `validationType` 构造规则 → `clear()` → 写入。带 `minVal/maxVal` 的"读"请求会**真的写一条校验**；
  - 原生 COM：`Validation.Delete()` 之后 `Add(2,1,operator,$null,$null)` → **清掉既有校验**再写一条空规则。
  换言之：调用方只想读现状，结果改了文件（Windows 上还会把既有校验清掉）。
- **改后是什么行为**：`normalizeOfficeRequest` 在调用宿主前拒绝 MS 通道的 `set_data_validation(action='read')`，
  并标记为**参数安全护栏**（`unsafe`）→ `routeOfficeFailure` 拒绝换通道重放（否则 Windows 上正好撞上 COM 那条破坏路径）。
  错误里给出替代路径：`host=wps` 的 `wps_set_data_validation(action='read')`（WPS 侧已实现读回）或 `wps_execute_script` 读 `Range.Validation`。
  WPS 通道不受影响（`host=wps` 不进 normalizer，读回照常可用）。
- **验证**：`tests/office.test.ts` 新断言（拒绝 + `isUnsafeParamError === true` + 写入路径不受影响）；
  `tests/failure-routing.test.ts` 的 unsafe 决策表。**未真机**。

### 2.3 【破坏性】数据有效性：缺 `listItems` 时"先清空再报错"（桥接侧兜底）

- **改前会发生什么破坏**：`wps_set_data_validation` 只传 `address`（或只传 `minVal/maxVal`，网关会把
  `validationType` 默认成 `'list'`）→ WPS 宿主的实现**先** `Validation.Delete()`、**再**因缺候选项抛错
  （`wps-addon/src/excel.js:4073` 早于 `:4079`）→ **该区域既有的数据有效性被清空，调用方只看到一条报错**。
  这是 CAP-50 描述的那一例。
- **改后是什么行为**：`gateway/excel.ts` 的 `setDataValidation` 在调用宿主前校验
  「`validationType==='list'` 必须有非空 `listItems`」；不满足直接报错，并直接给出正确写法
  （`number_range` + `minVal/maxVal`，或 `action:'read'`）。`action:'read'` 不受此校验影响。
  **任何通道都不会再走到"先破坏再失败"**：Office.js/COM 本来就在删除前拦，现在 WPS 侧也被桥接层挡住。
- **验证**：`tests/contract-consistency.test.ts` 新增用例断言**宿主调用次数为 0**（用桩 `callWps` 计数）、
  合法请求（带 `listItems`）仍照常下发。**未真机**。
- **附加说明**：宿主侧的顺序问题（`Delete()` 早于参数校验）**根因在 `wps-addon/src/excel.js`**，
  不在我的独占范围，Lead 已转 task-6 做纵深修复；桥接层这层校验与宿主侧修复**不冲突**，
  宿主修好后本层变成纯粹的"更早、更可读的报错"。

### 2.4 【破坏性】空搜索串的查找替换：护栏被"跨通道回退"绕过（本轮最重要的一条）

- **改前会发生什么破坏**（完整链路，源码可证）：
  1. schema 允许 `searchQuery: ''`（JSON-schema 不拦空串；网关只拦 `undefined/null`）；
  2. MS 通道：桥接 `normalizer` 在调用宿主前判为破坏性并抛错 —— **但这次"执行前拒绝"被 adapter
     当作普通拒绝处理**：`normalizeOfficeRequest` 抛错 → `fallbackToNative()` → `routeOfficeFailure`
     见 `executed === 'no'` 且方法在 COM 白名单内 → **允许回退**；
  3. Windows 原生 COM 的 `excel.ps1` 对空查询用 `IndexOf('')`（恒 ≥ 0，整片区域都算命中）并按
     `regex('')` **逐格替换** → `"abc"` 变成 `"XaXbXcX"`。**护栏不仅没生效，还亲手把请求转交到了没有护栏的通道**；
  4. WPS 通道不走 normalizer：宿主 `findAndReplace` 只拦 `undefined/null`（`wps-addon/src/excel.js:2856`），
     空串同样 `includes('')` 恒真 + 空正则逐格替换 → **真机 WPS 上同样会改坏数据**。
- **改后是什么行为**：
  1. `gateway/excel.ts` 的 `findAndReplace` 增加空串前置校验（**三个通道共用这道闸**：wps / Office.js / COM 都在它之后才被调用）；
  2. `office/normalizer.ts` 的 `find_and_replace` 分组把该拒绝标记为 `UnsafeParamError`，
     `search_cells` 分组补上同样的空串护栏；
  3. `errors.ts` 新增 `unsafe` 语义：**参数安全护栏的拒绝不得换通道重放**（与"能力缺口"区分开——
     能力缺口仍保留 COM 接管，例如"Office.js 只开筛选按钮"仍由 COM 执行）；
  4. `adapter.ts` 在参数转换失败时把 `unsafe` 透传给 `BridgeError`。
- **验证**：`tests/failure-routing.test.ts` 新增「CAP-50：参数安全护栏（空搜索串）不得换通道重放，能力缺口仍可回退」：
  - 决策表：`unsafe` 的 `rejected` 即使方法在 COM 白名单内也 `reject`，且保留 `kind/executed`；
  - **walk through 真实调用路径**：win32 语义 + 桩通道下 `callOffice('find_and_replace', {searchQuery:'', replaceText:'X'})`
    断言 **Office.js 宿主调用 0 次、原生 COM 调用 0 次**；
  - 反向断言：能力缺口（`set_filter_and_sort` 只开筛选）仍回退 COM 1 次。
  `tests/contract-consistency.test.ts` 另断言 wps 通道下宿主调用 0 次。**未真机**。

### 2.5 【静默降级·非破坏】`borders` 颜色在 MS 通道被静默忽略

- **改前会发生什么**：schema 的 `borders` 既接受 `true` 也接受 `'#RRGGBB'`。WPS 宿主按该色上边框；
  Office.js 读的是 `params.borderColor`（缺失时用固定 `#D9D9D9`），而桥接层从不设置它 →
  **调用方指定品牌色画框，MS 侧实际是别的颜色**（`success:true`，无任何提示）。
- **改后是什么行为**：normalizer 在 `format_cells` 里把同一色值同时放到 `borderColor`
  （仅当 `borders` 是合法 `#RRGGBB`）。`borders:false/'none'` 保持原样——schema 已明确写明
  "传它们不会去除已有边框"，是**文档化的 no-op**，不属于本轮缺陷。
- **验证**：typecheck + 现有 `format_cells` 单测（未断言 `borderColor`，故不冲突）。**未真机**（MS 侧观感未验证）。

### 2.6 未改、已由 Lead 转出/裁决的两条（如实登记，不再是待办）

1. **WPS 宿主 `findAndReplace` 空串无护栏**：桥接层已挡住经由工具入口的调用，但**直接调用加载项 RPC**
   （或未来新增入口）仍可触发 → 建议在 `wps-addon/src/excel.js:2856` 把条件从 `undefined || null` 扩到"空串"，
   与 `searchCells`（已拦空）保持一致。**已转 task-6（excel-cap-fixer）**；本轮未改（非我独占文件）。
2. **`wps_search_cells` 的 `address` 到不了宿主**：加载项 `searchCells` 支持 `address` 限定范围
   （ISS-29 的修复），但 schema 没有该参数（`additionalProperties:false` 会直接拒绝），
   网关 `searchCells` 处理器也未转发 → **宿主能做、工具调不到**。修它需要同时改
   `tools/definitions/excel.ts` 与 `gateway/excel.ts`。**Lead 已转 task-6**；本轮未改。

### 2.7 超出原清单的一处改动（Lead 已批准保留）

按"CAP-50 里桥接层能做的部分继续做"，本轮**额外**改了 **`src/bridge/gateway/excel.ts`**（两处护栏：
`setDataValidation` 的 listItems 前置校验、`findAndReplace` 的空串前置校验）。
该文件不在我最初的独占清单内、也不在禁止名单内，且无人独占；两处改动都紧贴既有参数校验、
只做"调用宿主之前拒绝"，不改任何正常路径。
**Lead 裁决（本轮）：批准保留**——理由：① 紧贴既有参数校验，不是另起一套；② 关掉的是
"护栏把请求送到没有护栏的通道"这个逻辑洞；③ `unsafe` 语义把"参数安全拒绝"与"能力缺口"分开，
能力缺口仍允许 COM 接管。**不需要回退**。

---

## 3. ISS-126：残留目标锁 → 自动清理 + 可操作提示

### 根因

目标锁存在**桥接进程内存**里（`TargetLockStore`，按 `sessionId:host` 分桶）。宿主重启/文档关闭后，
锁里的名字在宿主里已不存在，但 `UniversalGateway.executeTool` 仍会把它注入到每次调用
（`args.documentName = TargetLockStore.resolve('word', ...)`）→ 本会话**后续所有** Word/PPT/Excel 工具
都带着这个不存在的名字去调宿主，宿主报
`未在 WPS 中找到目标 Word 文档 [agent-word.docx]。当前已打开: 测试文字文稿.docx`。
报错本身虽然列了"当前已打开"，但**没有一句指向"该重新锁目标"**，使用者会以为工具坏了；
而唯一的自救入口 `wps_get_locked_status` 对 Word/PPT 只给 `exists: null`（"未核对"）。

### 改法

1. `gateway/lock.ts` 新增 `isTargetNotFoundError()` / `reconcileStaleTargetLock()`：
   - 只在**两个条件同时成立**时清理：① 报错匹配"目标名字在宿主里不存在"（三个通道的措辞逐条登记：
     WPS 加载项 `未在 WPS 中找到目标 Word 文档/演示文稿/工作簿`、COM `Target must match exactly one open document`、
     macOS JXA `找不到指定目标文档`）；② 该目标是**由会话锁注入**的，不是调用方显式传入的；
   - 清理后用**可操作提示**改写错误：点明"目标锁已失效并自动解除"、保留宿主原始错误、
     给出三步（查 `wps_get_locked_status` → 重新 `wps_lock_target_document` 或显式传名 → 后续不再被带偏）；
   - 保留原失败的 `kind/channel/executed`（是 `BridgeError` 时按同分类重建），调用方仍能判断能否重试。
2. `gateway.ts` `executeTool`：只增加"记录注入来源"（每分支一行）与包住 handler 调用的 try/catch
   （`throw reconciled ?? error`）。其余逻辑与顺序完全不动（含"成功后才登记身份类操作"）。
3. `gateway/lock.ts` `getLockedStatus`：新增 `staleLockTargets`（表格锁可核对的部分）与
   `staleLockNote` / `wordLockNote` 提示，把"锁失效后怎么办"提前暴露在查询路径上。

### 验证到什么程度

- `tests/failure-routing.test.ts` 新增「ISS-126」用例，用桩 `bridgeServer.callWps` 跑通**真实 `executeTool` 路径**，三个场景：
  ① 会话锁注入 + 宿主报"未找到目标 Word 文档" → 错误含"目标锁已失效并自动解除"+ 指引，**且 `TargetLockStore.getLocks()` 变为 `{}`**（自动清理生效）；
  ② 调用方**显式传名**失败 → 不改写、**不动会话锁**（避免误删用户刚设的锁）；
  ③ 与目标无关的失败（"段落下标超出范围"）→ 不触发清理。
- 全部非产物单测通过。**未真机**：没有制造"宿主重启"场景（那需要真机/服务操作，属 Lead 串行范围）。
- **已知缺口（Lead 裁决：不修，本次范围外）**：`MsOfficeDriver` 自身的 `locks`（`office_lock_target` 那条链，
  用于 `office_execute_script`）是**另一套存储**，本轮未纳入"失效自动清理"。同源缺陷在这条链上仍然存在，
  下一个会话若要统一处置，改动点会在 `office/ms-office-driver.ts`。

---

## 4. CAP-42：`wps_word_capture_preview` 注册链路核验

**结论：HEAD 上已经接上了，本轮无需改代码。** 核验 5 环（全部静态/运行时读工具清单）：

| 环节 | 证据 | 结果 |
|---|---|---|
| 对外 schema | `src/bridge/tools/definitions/word.ts` 的 `wps_word_capture_preview` 定义（在 `wordToolDefinitions()` 返回数组内） | ✅ |
| 工具清单 | `getTools().find(name)` 命中，`annotations={readOnlyHint:false, openWorldHint:false}` | ✅ |
| 网关执行分支 | `registeredGatewayBranches()` 含 `wps_word_capture_preview` → `word.capturePreview` | ✅ |
| 分类 | `toolClassOf('wps_word_capture_preview') === 'word'`（能归类，不是孤儿） | ✅ |
| 宿主实现 | `wps-addon/src/dispatch.js` 有 `case "word_capture_preview"`；**已部署**的 `wps-addon/addon-core.js` 也含该方法 | ✅ |
| 死分支台账 | `tests/contract-consistency.test.ts` 的「实现了但未注册」断言通过（台账里只剩 `wps_eval_code`、`wps_ppt_add_chart`），说明它**不再是**死分支 | ✅ |

补充：`gateway/word.ts` 的 `capturePreview` 还带**落盘校验**（ISS-111：宿主回 `hasPdf:true` 但磁盘无文件时报错，
而不是返回假成功）——即"成功"这一层已被加固过。

**验证到什么程度**：静态 + 工具清单 + 台账测试；**未真机调用**（因此"导出 PDF 并落盘"这一层仍只有 ISS-111 时的历史真机结论，
本轮不重复声明）。`review` 该工具描述说的"返回路径不是图片"与实际实现一致（`previewPath('pdf')` + `verifiedOnDisk`）。

---

## 5. 验证清单（可复现）

| 命令 | 结果 |
|---|---|
| `npm run typecheck` | ✅ 通过（多轮执行，最后一次在全部改动之后） |
| `node --import tsx --test tests/office.test.ts tests/contract-consistency.test.ts tests/contracts-boundary.test.ts tests/failure-routing.test.ts tests/platform.test.ts tests/addon.test.ts tests/ppt-layout.test.ts tests/win-icon.test.ts` | 67 项中 **66 通过 / 1 失败** |
| `npm run check:params` | ✅ 扫描 82 个工具定义，"schema 有、处理器不读"为 0 |
| `npm run check:claims` | ✅ 29 个文件，无与元数据源矛盾的数量表述（口径：路由表 50 / 对外工具 138 / WPS 47 / MS 37） |
| 未运行 `tests/session-isolation.test.ts` | 按任务说明（拉起 `dist/` 旧产物）跳过 |
| 未运行 `npm run build*` / `setup` / 真机工具 | 按硬约束（服务与构建产物为独占资源） |

**唯一失败项**：`tests/ppt-layout.test.ts` 的「生成物与源码一致，且不含 import/export」——
`node scripts/build-wps-addon.mjs --check` 报"生成物与 `wps-addon/src/**` 不一致"。原因：本轮工作区里
`wps-addon/src/excel.js`、`wps-addon/src/word.js` 被其他 teammate 改动但尚未重建（`git status` 可证），
**与桥接层改动无关**；属任务说明中"改了源码导致产物检查失败是预期的"那一类。

---

## 6. 本轮新增/修改的测试（防回归）

| 文件 | 新增内容 |
|---|---|
| `tests/contract-consistency.test.ts` | CAP-51 同源守卫（解析 `excel.ps1` 真实分支）；CAP-50 破坏性前置条件拦截（宿主调用计数为 0） |
| `tests/failure-routing.test.ts` | CAP-50 `unsafe` 不换通道决策表 + 真实调用路径（Office.js/COM 均 0 次调用）；ISS-126 三场景（自动清理 / 不误删 / 不误清） |
| `tests/office.test.ts` | CAP-50：MS 通道 `set_data_validation(action='read')` 判为 unsafe + 空搜索串判为 unsafe |

---

## 7. 未验证 / 已裁决（如实汇总；**已无待办**）

1. **真机一律未验证**：CAP-51 的 Windows 回退行为、CAP-50 五处改动的宿主实际行为、ISS-126 的崩溃恢复场景、
   CAP-42 的预览落盘，全部只有源码 + 单测结论，**没有**任何真机/服务操作（硬约束）。
   上线需要 Lead 统一 build/部署后再做真机验收。
2. **CAP-51 残留**：`capabilities()` 未暴露 COM 覆盖清单，而拒绝文案却让调用方去那里查 → **Lead 裁决不修**，
   记为已知缺口（§1 遗留口径）。动它要同时处理 `catalog.ts` + `check:claims` + 快照。
3. **`MsOfficeDriver` 的另一套锁**（`office_lock_target`）未纳入自动清理 → **Lead 裁决不修**（本次范围外），
   记为已知缺口（§3）。
4. **`wps_search_cells` 的 `address` 未暴露** → **Lead 已转 task-6**（§2.6-2）。
5. **WPS 宿主 `findAndReplace` 空串无护栏** → **Lead 已转 task-6**（§2.6-1）。
6. **MS 通道 `set_data_validation(action='read')` 现在是"执行前拒绝"**：这是桥接层的如实处置，
   若产品希望 MS 侧也能读回，需要在 Office.js 加载项补 action 分支（他人范围，本轮未转出，登记备查）。
7. `gateway/excel.ts` 的两处改动超出我最初的独占清单 → **Lead 裁决：批准保留**（§2.7）。
8. 报告作者未同步 `issues.md` / `capability-backlog.md` 的状态列 —— 按 Lead 要求由其最终轮统一同步。

