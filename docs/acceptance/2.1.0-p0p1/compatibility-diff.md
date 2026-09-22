# P1 兼容性差异

候选标识：`2.1.0-p0p1`
比对对象：`tools-snapshot.baseline.json`（P1 改动前）→ `tools-snapshot.p1.json`（P1 完成后）
采集方式：`npm run snapshot:tools -- <目标文件>`，两次导出使用同一脚本、同一口径，已剥离时间戳等易变字段。

---

## 1. 对外工具契约：无变化

| 比对项 | 结果 |
|---|---|
| 工具总数 | 91 → **91** |
| 删除的工具 | **无** |
| 新增的工具 | **无** |
| `inputSchema` 变化 | **无**（逐工具 schemaHash 全等） |
| `description` 变化 | **无**（逐工具 descriptionHash 全等） |
| `excel_*` 统一入口、`wps_*` 兼容入口、别名 | 全部保留，未改名 |

结论：**P1 未改动任何工具名、参数 schema 或工具说明**。既有调用方、skill、MCP 客户端无需适配。

## 2. MCP `readOnlyHint` 注释：4 项由“写”更正为“读”

| 工具 | 基线 | P1 | 依据 |
|---|---|---|---|
| `wps_get_locked_status` | false | **true** | 纯本地返回目标锁状态，不触达宿主 |
| `wps_word_read_document` | false | **true** | Word 读取分支 |
| `wps_ppt_read_presentation` | false | **true** | PPT 读取分支 |
| `wps_ppt_get_slide_shapes` | false | **true** | PPT 形状读取分支 |

`readOnlyHint = true` 计数 17 → **21**。

这是**声明更正**，方向是从“偏保守（把只读标成写）”变成“与实现一致”，不存在把写入标成只读的风险；已核对全部 21 项均为真正只读。副作用是某些 MCP 客户端可能据此对上述 4 个工具降低确认强度——这 4 项均不修改文档，属预期行为。

### 2.1 `wps_inspect_api` 明确**不**纳入只读与重放（评审补正）

初版 P1 曾把 `wps_inspect_api` 同时加入只读集合与跨通道重放白名单，这是错的，已撤回：

- `wps_inspect_api` → `callOffice('inspect_api')`，执行的是**调用方给出的任意表达式**，不是受限读取。
- 表达式可以带副作用：取属性会触发宿主求值，也可以直接改文档，例如 `app.ActiveWorkbook.Worksheets.Add()`、`wb.Worksheets.Item(1).Delete()`、`pres.Slides.Add(1, 12)`、`doc.Content.InsertAfter("x")`。
- 因此静态无法保证它只读：既不能标 `readOnlyHint = true`，也不能在"结果不可判定"时改走另一通道重放（否则副作用表达式会在第二个通道再执行一次）。

最终状态：`wps_inspect_api` 的 `readOnlyHint` 保持 **false**（与基线一致），且不在 `REPLAY_SAFE_METHODS` 内。只有确认未执行（`executed === 'no'`）时才允许回退。回归测试见 `tests/failure-routing.test.ts` 的"带副作用的表达式探测既不算只读，也不允许跨通道重放"。

因此本次 `readOnlyHint` 变化为 **4 项**（不是 5 项）：`wps_inspect_api` 从"曾计划改"回到"与基线一致"。

注意：`MCP readOnlyHint` 与 `跨通道重放安全性` 现在是两个命名清晰、各自成对的判定（`errors.ts` 的 `isReadOnlyTool` / `isReplaySafeMethod`），不再共用一个集合。

## 3. `bridge_get_capabilities` 响应：字段保留、取值如实化

顶层键**完全保留**（`version` / `platform` / `preferredWorkflow` / `connection` / `hosts` / `audit` / `lifecycle` / `repairs` / `limitations`），`audit` 三项与基线逐字相同，因此消费方读取路径不破。

### 3.1 `hosts.microsoft.excel`

| 字段 | 基线 | P1 | 原因 |
|---|---|---|---|
| `implemented` | 30 项 | **27 项** | 原值直接回显 `EXCEL_METHODS`，其中 3 项没有可调用工具（见 D2/D3） |
| `implemented` 新增字段 `declaredNotCallable` | — | `["clear_range","insert_dimension","rollback_cells"]` | 如实列出“声明了但不可调用” |
| `implemented` 新增字段 `unimplementedOnHost` | — | `[]` | Microsoft 侧无实现缺口 |
| `validation` | `全量 42 项结构化能力支持 (macOS / Windows 统一)` | `代码路径已实现，但本候选版本尚未在真实 Microsoft Excel 上做实机验收；静态与模拟通道不代表实机通过` | 见下 |

`42 项` 与任何口径都对不上：`EXCEL_METHODS` 为 30 项，实际生成的 `excel_*` 工具为 27 项，全仓 grep “42 项”只有这一处且无对应清单。`(macOS / Windows 统一)` 更是一句**无证据的验收声明**——本机无 Windows 环境、Microsoft Excel 结构化通道未实机验收，故按“已实现 / 当前可用 / 实机已验证”三级口径改写。

### 3.2 `hosts.wps.excel`

| 字段 | 基线 | P1 |
|---|---|---|
| `implemented` | 30 项 | **26 项** |
| 新增 `declaredNotCallable` | — | `["clear_range","insert_dimension","rollback_cells"]` |
| 新增 `unimplementedOnHost` | — | `["update_chart"]` |
| `validation` | 保留原文 | 保留原文（原本就如实） |

`unimplementedOnHost: ["update_chart"]` 的依据（静态核对，**未实机复现**）：`wps-addon/addon-core.js` 的 RPC 分发 switch 只有 `add_chart` / `get_charts` / `delete_chart`，没有 `update_chart`，未识别方法会落到 `default:` 抛 `未知的 RPC 方法`；而 `update_chart` 被列在 `EXCEL_METHODS` 中，因此 `excel_update_chart`（`host=wps`）静态可判定必然失败。Microsoft 侧 `office-addon/public/taskpane.js` 有 `update_chart`，故不列入缺口。

**行为未改变**：P1 只改声明，**没有**在网关或适配器里对 `update_chart` 做预先拦截，`excel_update_chart`（`host=wps`）的路径与失败方式与基线一致。

### 3.3 `hosts.microsoft.word` / `hosts.microsoft.ppt`

| 字段 | 基线 | P1 |
|---|---|---|
| `word` | `Office.js 原生通道` | `未提供 Microsoft 结构化通道：Word 工具为 WPS 专用；Microsoft Word 需改用 office_execute_script 原生脚本（Windows 分支待实机验收）` |
| `ppt` | `Office.js 原生通道` | 同构表述（PPT） |

基线说法与实现冲突的证据：`office-addon/` 只有 `excel/manifest.xml` 一个清单，`office-addon/public/taskpane.js` 全文没有任何 `word_`/`ppt_` 分支；适配器对非 `EXCEL_METHODS` 方法会直接抛 `Microsoft Excel 尚不支持 <method>`。Word/PPT 的真实 Microsoft 通道是 `office_execute_script`（`MsOfficeDriver`），不是结构化 Office.js 工具。

## 4. 错误语义变更（对外可观察）

P1 之前，通道失败一律是普通 `Error`；P1 之后，通道边界与入参校验抛出 `BridgeError`，携带 `kind`（`unavailable` / `rejected` / `failed` / `unknown`）、`channel`、`method`、`executed`，且 `JSON.stringify(error)` 给出结构化描述。

| 场景 | 基线 | P1 |
|---|---|---|
| 加载项未连接 | `Error`，文本含“加载项未连接” | `BridgeError(kind='unavailable', executed='no')`，**文本未变** |
| 调用超时 | `Error`，文本含“超时” | `BridgeError(kind='unknown')`，**文本未变** |
| 加载项断开 | `Error('加载项断开，执行结果未知，请先读回确认')` | `BridgeError(kind='unknown')`，**文本未变** |
| `Bridge 已停止` | `Error('Bridge 已停止')` | `BridgeError(kind='unknown', 'Bridge 已停止，结果未知，请先读回确认')`，**文本有变化** |
| `ws.send` 回调报错 | 直接抛原始 `Error` | `BridgeError(kind='unknown', '…请求发送失败：<原因>；结果未知，先读回确认')`，**文本有变化** |
| 入参校验失败 | `Error`，文本如 `arguments.address: 必填` | `BridgeError(kind='rejected', executed='no')`，**文本未变** |
| 工具返回 `success:false` | `Error(result.error \|\| result.message)` | `BridgeError(kind='failed')`，**文本未变** |
| Microsoft 替代通道不支持该方法 | `Error('Microsoft Excel 尚不支持 <method>，请查询能力清单。')`——**丢弃了前一通道的错误与"是否已执行"** | `BridgeError`，**保留前一通道的 `kind` / `executed` / 原始文本**，并追加"未回退：<method> 不在 Windows 原生通道支持列表内，可用 bridge_get_capabilities 查询能力清单" |

### 4.1 替代通道不支持时的错误语义（评审补正）

初版 P1 在"方法不在替代通道白名单"这一分支上硬编码了 `kind='rejected'`、`executed='no'`，并替换掉原始错误文本。这有两个问题：

1. **把"结果未知"改写成"未执行"**。如果前一通道是超时或断连（`executed='unknown'`，宿主可能已执行），却被报成 `executed='no'`，调用方会据此认为可以安全重放——恰好把 P1.3 要防的重复写入又放回来了。
2. 丢失原始失败原因，排障时看不到前一通道到底发生了什么。

修正后该分支只**追加**说明，不改写种类与执行状态：

| 前一通道失败 | 替代通道不支持时返回 |
|---|---|
| `unavailable`（未连接，`executed='no'`） | `kind='unavailable'`，`executed='no'` |
| `failed`（宿主已报错，`executed='unknown'`） | `kind='failed'`，`executed='unknown'` |
| `unknown`（超时/断连，`executed='unknown'`） | `kind='unknown'`，`executed='unknown'` |

回归测试见 `tests/failure-routing.test.ts` 的"替代通道不支持该方法时，保留前一通道的原始错误与 executed 状态"。

变化仅限“错误对象的类型与结构化字段”，外加两处文本补充。依赖 `error.message` 做前缀匹配的调用方在这两处需要放宽匹配；`instanceof Error` 判定不受影响（`BridgeError extends Error`）。

## 5. Microsoft 通道失败处置变更（本次唯一的真实行为变化）

**变更点**：`src/bridge/office/adapter.ts` 不再对“宿主已执行后才发现失败”的情况回退到 Windows 原生 COM 通道。

| 场景 | 基线行为 | P1 行为 |
|---|---|---|
| 加载项未连接（确认未执行） | 回退 COM | **回退 COM（不变）** |
| 客户端参数转换失败（确认未执行） | 回退 COM | **回退 COM（不变）** |
| 读取方法超时/断连 | 回退 COM | **回退 COM（不变）** |
| **写入方法超时/断连** | **回退 COM，同一写入执行第二次** | **拒绝回退，报 `kind='unknown'`，要求先读回** |
| **写入方法宿主报错** | **回退 COM，同一写入执行第二次** | **拒绝回退，报 `kind='failed'`，要求先读回** |
| **宿主已执行但响应无法解析** | **回退 COM，同一写入执行第二次** | **拒绝回退，报 `kind='failed'`** |
| 非 Windows 平台 | 直接抛出 | **直接抛出（不变）** |
| 替代通道不支持该方法 | 抛“尚不支持” | **抛“尚不支持”（不变）** |

影响面：仅在 Windows + `host=microsoft` + 写入类方法 + 结果不可判定时生效。此前该分支会让一次写入**执行两次**并且向调用方返回“成功”（第二次执行的结果掩盖了第一次的失败）；现在改为明确报错。这是**故障场景**的行为修复，正常路径兼容性未变。故障前后的复现证据见 [evidence/p1.4-prefix-failure.txt](evidence/p1.4-prefix-failure.txt) 与 §6。

## 6. 证据索引

| 文件 | 内容 |
|---|---|
| `tools-snapshot.baseline.json` | P0 基线工具契约快照（`a016c010…`） |
| `tools-snapshot.p1.json` | P1 补正后快照（`2b1abe2d…`）；仅 `readOnlyHint` 4 处变化，工具数/schema/说明与基线全等 |
| `business-samples.json` | 模拟宿主真实往返样例 |
| `baseline/BASELINE.json` | 固定锚点提交（`cac6d38…`）与不可恢复清单；恢复与失败复现都读取它，不用会前进的 `HEAD` |
| `baseline/baseline-prior.patch` + `MANIFEST.sha256` + `verify-baseline.mjs` | 可恢复基线：按固定锚点恢复实测 22/22 哈希一致；5 个未跟踪文件不可恢复（见 `baseline/README.md`） |
| `baseline/verify-after-head-advance.sh` | 隔离验证：副本仓库 HEAD 前进后仍 22/22 通过 |
| `evidence/reproduce-prefix-failure.sh` | 在临时副本中复现修复前失败，不覆盖当前工作区 |
| `evidence/baseline-head-anchor-prefix-failure.txt` | "恢复依赖 HEAD"缺陷的修复前复现（`patch does not apply`，退出码 1） |
| `evidence/p1.4-prefix-failure.txt` | 故障注入测试在**修复前**代码上的失败输出（`tests 11 / pass 9 / fail 2`）；产生方式与断言含义见 [evidence/README.md](evidence/README.md) |
| `baseline.md` | P0 基线与检查结果 |
| `tool-inventory.md` | 91 个工具的宿主映射与 9 条能力声明差异 |
| `environment.md` | 环境、发布矩阵与签名要求 |
| `evidence/README.md` | 修复前日志的产生方式与断言含义 |
