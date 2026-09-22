# ISS-10 · 9 条未确认项的当前状态

收口时间：2026-09-22 · 对应问题：[issues.md](../issues.md) ISS-10 · 原始登记：[tool-inventory.md](../tool-inventory.md) §五（第 180–190 行）

## 这份文件解决什么

`tool-inventory.md` §五 如实登记了 **9 条"静态源码推断不出来"的未确认项**，并明确"不作推断"。那 9 条写在**第一轮盘点（15:22）**的时间点上；此后 P5 的真实宿主盘点（`mcp-sweep/`）、业务矩阵（`p5.4`）、总验收审计与修复批次（`issues.md` ISS-46 起）已经回答了其中一部分。

本文件**逐条给出当前状态**，只有两种落点：

- **已澄清** —— 有可复核的证据文件把它答了；写清结论 + 证据指向。
- **仍未确认** —— 保留登记，并写清**下一步怎么确认**（具体到"打开什么、调哪个工具、看哪个字段"）。

判定口径：只认**已落盘的证据文件、源码事实或实测记录**；"应该没问题"一律不算澄清。

## 汇总

| # | 未确认项（原文摘要） | 当前状态 | 结论一句话 |
|---|---|---|---|
| 1 | 实机可用性未确认 | ✅ **已澄清** | 第一轮的"纯静态"局限已被真实宿主盘点解除；结论是**能跑通成品，但大量工具"返回 success 却没生效"** |
| 2 | 部署版本与源码是否一致未确认 | ✅ **已澄清（结论：不一致）** | 实测查明"磁盘上是新构建、WPS 进程里跑的是旧构建"，已升级为 ISS-59（高） |
| 3 | "42 项"的来源未确认 | ✅ **已澄清** | 来源是两处手写文案：`catalog.ts` 的 validation 字段 + `taskpane.html`，无对应清单；两处均已改写 |
| 4 | 各工具的宿主超时阈值未确认 | ❌ **仍未确认** | 只有 1 个实测锚点（21 s 超时），逐工具阈值表仍未建立 |
| 5 | 无专用读回工具的写操作失败判定未确认 | ✅ **已澄清** | 缺口成立且已登记（ISS-44/94）、修法已决策（补读回工具）；判定路径已用脚本读回实测走通 |
| 6 | host=microsoft 下审计快照缺失的实际后果未确认 | ❌ **仍未确认** | Microsoft 通道自始未连接，无法运行时验证；只有静态结论 |
| 7 | Microsoft 侧 Word/PPT 能力边界未确认 | ✅ **已澄清** | `08-ms-vs-wps.md` 给出双通道能力对照；ISS-93 定性"21 个 Word/PPT 工具在该宿主是死路" |
| 8 | 8 行的"部分成功语义"未逐函数核对 | ✅ **已澄清 8/8** | 8 个函数全部有了实测或代码复核结论，且多数升级为独立 ISS 条目 |
| 9 | 审计记录条数上限的副作用未确认 | ❌ **仍未确认** | 代码事实已复核（只留最新 500 条），超过 500 条后的实际后果从未验证 |

**合计：9 条中 6 条已澄清、3 条仍未确认（#4 / #6 / #9）。**

> 状态口径说明：标记"已澄清"不等于"问题已解决"。#1、#2、#5、#8 澄清出来的都是**真实的缺陷**，已各自登记为独立 ISS 条目（见每条"派生条目"）。"澄清"指的是"这条未知项不再是未知"，不是"这条没问题"。

---

## 1. 实机可用性未确认 —— ✅ 已澄清

**原文**：本清单未连接任何宿主、未执行任何工具；所有"宿主路由/审计/部分成功"结论都是静态推导。

**为什么当时那样写**：`tool-inventory.md` 是只读盘点，依据是源码 + 工具清单快照，确实一个工具都没真跑。

**当前状态**：这条局限**已经被后续工作解除**。P5 做了真实宿主盘点：macOS 27.0 arm64 / WPS 12.1.28496 / 加载项 2.1.0，4 个并行子代理各自交出**真实成品**，并做了端到端链路核对。

**结论（实测读数，不是推断）**：

| 事实 | 数字 | 证据 |
|---|---|---|
| 派出的并行子代理 | 4 个（`inventory`/`tables`/`charts`/`shapes`），只给成品目标不给步骤 | [mcp-sweep/README.md](mcp-sweep/README.md) §方法 |
| 真实成品 | 3 个（134 个矢量元素的信息图 / 5 张原生图表的看板 / `A1:N34` 企业跟踪看板），均已由协调方拉渲染图核对 | 同上 §三个真实成品；`shapes-infographic-final.png`、`charts-dashboard-final.png`、`tables-dashboard-final.png` |
| 端到端链路 | 跨组件一致性 **18/18 项完全一致**，无丢行、无丢精度、无串表 | [05-e2e.md](mcp-sweep/05-e2e.md) §步骤 6 |
| 只读工具实测 | `inventory` 只读实测 11 个；**70 个写工具**的真实行为仍未逐条验证 | [01-inventory.md](mcp-sweep/01-inventory.md) §六 |
| 最典型的实测失败模式 | "返回 `success` 但什么也没发生"：ISS-38 / 39 / 17 / 29 / 11 / 42 六条 | [README.md](mcp-sweep/README.md) §本轮最典型的失败模式 |

**派生条目**（澄清出来的真实缺陷，已各自登记，不在本条重复）：ISS-11～ISS-45（第一轮盘点）、ISS-46～ISS-98（第二、三轮 + 审计）。

**残留（不构成本条"未确认"，是新的已知缺口，已另行登记）**：
- Microsoft Office 全通道未连接、未实机验收 → 见 ISS-93 / ISS-98、`mcp-sweep/README.md` §未验证；
- Word / PPT 的对象模型测绘未完成（`06-wps-api-map.md` 只完成了 Excel 的 48 个表达式，Word/PPT 未测绘，且因 ISS-89 崩溃而中断）；
- 70 个写工具里多数仍只有单点验证，不是覆盖性验证。

---

## 2. 部署版本与源码是否一致未确认 —— ✅ 已澄清（结论：**不一致**）

**原文**：`wps-addon/addon-core.js`、`office-addon/public/taskpane.js` 的分支结论只在"运行时加载项 = 当前源码构建"时成立；版本一致性仅在连接时由加载项上报版本与 `VERSION` 比对，本次未做连接验证。

**当前状态**：**已实测查明不一致，并且这是本轮最重要的方法学发现**，已升级为独立条目 **ISS-59（高）**。

**实测证据**（`mcp-sweep/README.md` 顶部版本说明，第一轮盘点结束后 15:57 查明）：

| 文件 | 含 `fitGeneratedPptShapes` | sha256 前 16 |
|---|---|---|
| WPS 已部署 `addon-core.js`（09-22 11:51 部署） | **有**（4 处） | `c30e548afaacbcfd` |
| 同名 `.backup-1789967115639`（09-21 13:05） | **无** | `3b34c6a5d7c74dde` |

- **成因**：WPS 进程在 09-22 11:51 那次部署**之前**就已启动，之后没有重载加载项 —— **磁盘上是新的，进程里跑的是旧的**。
- **行为指纹验证**（不靠版本号、靠行为）：干净实验里 `wps_ppt_generate_deck` 生成的形状**缩放系数恰为 1.0（完全没有换算）**，正是旧构建（无 `fitGeneratedPptShapes`）的行为 —— 见 [02-ppt.md](mcp-sweep/02-ppt.md) §1。
- **衍生结论**：依赖运行时行为的结论需重启后复测；已被**源码复核**确认的（ISS-01/38/39/40/42/54）在当前源码里同样存在，结论仍有效。
- **对照实验**：同一份文稿内出现"天然对照"——第一次生成跑旧构建、几分钟后同一次调用返回体已变成新格式（多出 `pageWidth/pageHeight/unit/layoutWarnings/visualVerificationRequired`），说明加载项在两次调用之间被重载过（[02-ppt.md](mcp-sweep/02-ppt.md) §1.6）。

**为什么"版本号比对"没能拦住它**：现有机制只在**连接时**把加载项上报版本与 `VERSION` 比对（`ws-server.ts`），而两个构建的版本号都是 `2.1.0` —— **版本号相同，构建不同**，比对不出差异；也没有任何指纹能判断"运行中的是哪一版"。

**派生条目**：ISS-59（部署流程 + 版本可观测性，高）、ISS-90（崩溃后宿主组件掉线、桥接无崩溃信号）。

---

## 3. "42 项"的来源未确认 —— ✅ 已澄清

**原文**：全仓仅 `catalog.ts:21` 一处出现该数字，无对应清单或计数逻辑可核对，无法判断它原本指哪一组能力。

**当前状态**：**来源已查明为两处独立的手写文案**，且都与任何真实计数对不上；两处均已改写。

| 出处 | 原文 | 与真实计数对照 | 处置 |
|---|---|---|---|
| `src/bridge/catalog.ts` 的 `microsoft.excel.validation` | `全量 42 项结构化能力支持 (macOS / Windows 统一)` | `EXCEL_METHODS` 实际 **30** 项；`excel_*` 工具实际 **27** 个 —— 42 与任何口径都对不上 | 按"已实现 / 当前可用 / 实机已验证"三级口径改写；删掉无证据的"(macOS / Windows 统一)" |
| `office-addon/public/taskpane.html:38`（**手写文件**，生成器不管它） | `42 项能力已就绪` | 该文件负责的是 Microsoft Excel 通道 | 改为 `Microsoft Excel · 27 项能力已就绪`；并纳入 `check:claims` 扫描范围 |

**证据指向**：
- [compatibility-diff.md](../compatibility-diff.md) 第 60–62 行（`42 项` 与任何口径都对不上的完整论证）；
- [issues.md](../issues.md) §ISS-04（交付包内残留）、§ISS-05（`check:claims` 扫描范围 + 模式双重漏网，已修已验证）；
- [final-review.md](../final-review.md) §10.1 / §11；
- 本目录复核：`office-addon/public/taskpane.html:38` 现为 `Microsoft Excel · 27 项能力已就绪`；`src/bridge/catalog.ts` 已无 `42`。

**★ 残留（本条已澄清，但残留一个未收口的事实）**：**已交付的三个包内仍是"42 项"**。复核命令与结果：

```
$ grep -on "4[0-9] 项[^<]*能力[^<]*" \
    "release/2.1.0/mac/mac-arm64/Office Agent Bridge.app/Contents/Resources/app.asar.unpacked/office-addon/public/taskpane.html"
38:42 项能力已就绪
```

处置已按使用者决策走**方案 B：不重打包，登记为已知差异**（见 [issues.md](../issues.md) §决策结论表 ISS-04 行），并入后续统一打包。**不是漏修**。

---

## 4. 各工具的宿主超时阈值未确认 —— ❌ 仍未确认

**原文**：多数网关分支不传 `timeout`（`office/adapter.ts:8`），实际超时由调用链默认值决定，未逐条核对。

**当前状态**：**仍未确认**。此后只积累了 **1 个实测锚点**，不足以反推阈值表。

**已有的唯一实测锚点**：

| 场景 | 观察 | 证据 |
|---|---|---|
| `Presentations.Add()` + `SaveAs`（原生脚本路径） | **21 s 后超时**；返回文案为 `WPS execute_script 超时，结果未知；先读取状态，不要自动重放写入`；但**读回确认文稿实际已建好**（`ppt3-blank.pptx, 960×540, 0 页`） | [02-ppt.md](mcp-sweep/02-ppt.md) §（第 250 行） |

这个锚点证明了两件**好事**：① 超时文案正确标了"结果未知、不要自动重放"（这正是 P1.3 修复要防的重复写入）；② 超时后**读回是可以判定真实结果的**。但它**没有**回答"阈值是多少、由谁决定、每个工具是否相同"。

**仍未确认的具体子问题**：
1. 超时阈值的确切取值（是不是 21 s？还是那次恰好 21 s 撞上某默认值？）；
2. 逐工具的阈值是否一致，还是按方法类别分档；
3. 网关分支不传 `timeout` 时，最终生效的是哪一层的默认值（网关 / HTTP 客户端 / 宿主加载项）；
4. 超时后 `executed` 判定（`unknown`）是否在**所有**写工具上都正确 —— 目前只有 1 例。

**下一步怎么确认**（可执行步骤）：
1. **静态先落表**：全仓枚举所有超时相关常量与默认值 —— `src/bridge/office/adapter.ts`、`wps-addon/src/connection.js`、`src/bridge/ws-server.ts`、HTTP 客户端调用点，产出一张"工具 → 是否显式传 `timeout` → 实际默认值来源"的对照表。这一步**不需要宿主**，可立即做。
2. **单点实测阈值**：连接 WPS 后，用一个可控的长耗时脚本（例如 `for` 循环空转到指定毫秒）二分逼近，实测出至少 3 个不同类别工具（读 / 写 / 脚本）的真实阈值与超时文案。
3. **验证 `executed` 判定**：对**每一类**写工具各造一次超时（宿主侧故意阻塞），读回确认实际是否已写入，核对返回的 `executed` 字段是否正确 —— 这是 ISS-59 之外最容易造成重复写入的位置。
4. 把结果补进 `tool-inventory.md` 的超时列，并把本条从"未确认"移出。

---

## 5. 无专用读回工具的写操作失败判定未确认 —— ✅ 已澄清

**原文**：条件格式、数据验证、透视表、冻结窗格、Word 页眉页脚水印等没有对应的读取工具，"失败时能否读回判定"在表中标为推断，需实机确认。

**当前状态**：**已澄清，且结论比当时更严重**——不只是"能不能读回判定"未确认，而是**这类工具确实存在"写不进去也返回 success"的静默失败**。缺口已登记，修法已决策。

**澄清出来的事实**：

| 事实 | 证据 |
|---|---|
| **缺读回工具是一类系统性缺口**：C 类"只写不读"共 **7 项**（条件格式、冻结窗格、数据有效性、筛选状态、工作表保护/标签色、Word 页眉页脚/水印、透视表） | [issues.md](../issues.md) ISS-94；[07-tool-gap.md](mcp-sweep/07-tool-gap.md) §19、§287 |
| 条件格式既无图标集/公式规则入口，也**无读取/清除工具** | ISS-64 |
| 冻结窗格只能用 `wps_execute_script` 探测（本次就是这么实测的），且实测出了真实语义问题 | ISS-41 |
| `excel_set_data_validation` 在 microsoft 通道上：先 `dataValidation.clear()` 再无条件不设置 → **清除原有校验且不报错**，返回 `success:true` | [07-tool-gap.md](mcp-sweep/07-tool-gap.md) §240 |
| **判定路径本身已走通**：脚本读回可行 —— ISS-41 就是靠脚本读回 `SplitColumn` 才得出结论 | ISS-41 §结论存疑 |
| 修法已决策：**补读回工具**（`get_conditional_formatting` / `get_data_validation` / `get_freeze_panes`），`clear_range` 按 DP3 再评估 | [issues.md](../issues.md) §决策结论表 ISS-44/45 行 |

**证据指向**：`issues.md` ISS-44 / ISS-45 / ISS-64 / ISS-94 / ISS-41；`mcp-sweep/07-tool-gap.md`。

**★ 残留（已转为独立待办，不再是"未确认"）**：读回工具**尚未实现**。在实现之前，任何格式/结构类写操作的验收都只能靠 `wps_execute_script` 探测——路径已知可行，但不是产品能力，AI 调用方**不会自己想到这条路**。

---

## 6. host=microsoft 下审计快照缺失的实际后果未确认 —— ❌ 仍未确认

**原文**：源码显示 Office.js 通道不返回 before/after 快照（`normalizer.ts:243-258`、`office-addon/public/taskpane.js:750-756`），据此 `gateway.ts:891` 会拒绝回滚；本次未实机验证该记录的实际落库形态。

**当前状态**：**仍未确认**，且**确认的前提条件始终没有具备**——Microsoft Office 通道自始未连接。此后所有涉及 MS 通道的条目都被迫停在静态层面（ISS-93 标"待修"、ISS-98 明确写"**MS 通道未连，无法运行时验证**"）。

**已经静态确认的部分**（这些不算未确认）：
- Office.js 通道不返回 before/after 快照；
- 因此 `excel_patch_cells`（host=microsoft）写入后，审计记录**没有可回滚的快照**；
- `gateway` 的 `wps_rollback` 会因此**拒绝回滚**并给出理由（不是静默成功）。

**仍未确认的具体子问题**：
1. 该记录在 **audit-store 里的实际落库形态**：是否真的被写入一条 `rollbackable:false` 的记录，还是根本不写记录？
2. 被拒绝时**返回给调用方的文案**是否可操作（是否说清"该通道不支持回滚"并给替代路径）？
3. 用户会不会因此**误以为"已写过、但回滚列表是空的"** —— 即审计历史里的可见性缺口。

**下一步怎么确认**（可执行步骤）：
1. **前置**：在 Microsoft Excel（macOS 或 Windows）里加载 Office.js 加载项并完成授权，使 `msExcel` 组件在 `/api/v1/status` 里变为 `connected`。这是本条的唯一阻塞点，**没有替代方案**。
2. 调一次 `excel_patch_cells`（host=microsoft），区域任取（如 `A1:B2`）。
3. 立刻调 `wps_get_audit_history`，检查：新记录是否存在、`host` 字段是否为 `microsoft`、是否带 `beforeSnapshot`/`afterSnapshot`/`rollbackable` 字段。
4. 用上一步的 `auditId` 调 `wps_rollback`，**逐字记录**返回的 `error` / `message` / `restoredScope`。
5. 判定标准：若 rollback 被拒绝且文案说清原因 → 本条澄清（结论=行为可预期）；若返回 `success` 但什么都没恢复 → **升级为新的高危条目**（假成功）。
6. 顺带核对 ISS-98（`excel_find_and_replace` 字段名错位）与 ISS-93（12 个 `excel_*` 参数错位），它们同样卡在这一个前置条件上。

---

## 7. Microsoft 侧 Word/PPT 能力边界未确认 —— ✅ 已澄清

**原文**：`office_execute_script` 的能力等于宿主 JXA/COM 脚本能力，没有静态清单可枚举；"Office.js 原生通道"的说法是否另有未暴露实现，未确认。

**当前状态**：**已澄清**——做了专门的双通道能力对照，并且"Office.js 原生通道"这个含糊说法已被判为不成立。

**澄清出来的结论**：

| 结论 | 证据 |
|---|---|
| 双通道能力逐条对照完成（**552 行**对照报告 + 补强建议） | [08-ms-vs-wps.md](mcp-sweep/08-ms-vs-wps.md) |
| **21 个 Word/PPT 工具在 host=microsoft 上是死路** | ISS-93 |
| `normalizer.ts` 只适配 **14** 个方法，**12 个 `excel_*`** 在 host=microsoft 参数错位 | ISS-93 |
| `capabilities()` 中 microsoft word/ppt 的"Office.js 原生通道"描述**已被判为不实**并改写 | `docs/refactoring-plan.md` P2.x；[compatibility-diff.md](../compatibility-diff.md) §3 |
| adapter 把整张 `EXCEL_METHODS`(30) 当 COM 回退白名单，而 COM 实际只覆盖 28/30 → **白名单过度声明** | ISS-97 |

**残留（极其有限，且已如实登记）**：`office_execute_script` 的**真实**能力上界 = 宿主 JXA/COM 脚本能力，这一点在静态层面不可枚举（脚本能调什么由宿主对象模型决定，而对象模型随 Office 版本变化）。但这不是"未确认"，而是**设计上的开放边界**：`office_execute_script` 本就是逃生舱（escape hatch），其能力面按定义等于宿主脚本面。已知的"宿主能做但没暴露"的部分已按死分支清单登记在 ISS-95。

---

## 8. 8 行的"部分成功语义"未逐函数核对 —— ✅ 已澄清 8/8

**原文**：`excel_freeze_panes`、`excel_set_data_validation`、`wps_freeze_panes`、`wps_set_data_validation`、`wps_word_write_content`、`wps_word_page_layout_and_watermark`、`wps_ppt_add_business_cards`、`wps_ppt_insert_native_chart` 在表中标为 `（推断）`，仅由 schema 与网关分支推导，宿主实现未逐函数阅读。

**当前状态**：**8 个函数全部拿到了实测或代码复核结论**，且多数升级为独立 ISS 条目。逐条如下：

| # | 函数 | 核对结论 | 证据 |
|---|---|---|---|
| 1 | `excel_freeze_panes` | 行列**对称**使用同一套 `-1` 约定（`SplitRow = freezeRowIndex - 1`、`SplitColumn = freezeColumnIndex - 1`）；子代理"必须传 4 才冻结两列"的结论**不成立**，判定**不修** | ISS-41 §结论存疑（代码复核） |
| 2 | `wps_freeze_panes` | 同上，同一实现路径 | ISS-41 |
| 3 | `excel_set_data_validation` | **清除原有校验且不报错**，返回 `success:true`（先 `dataValidation.clear()` 再无条件不设置） | [07-tool-gap.md](mcp-sweep/07-tool-gap.md) §240 |
| 4 | `wps_set_data_validation` | 无读取入口，属 C 类"只写不读"7 项之一；下拉项/区间/提示语均无法读回 | ISS-94；[07-tool-gap.md](mcp-sweep/07-tool-gap.md) §287、§442 |
| 5 | `wps_word_write_content` | **实测吞掉所有小写字母 `a`**（现象已确认，根因待隔离） | ISS-67 |
| 6 | `wps_word_page_layout_and_watermark` | 不接 `pageNumberFormat`（422）；水印**只落正文层第 1 页**；header/footer/watermark **只作用于第 1 节**且不提示 | ISS-57、ISS-58、ISS-72 |
| 7 | `wps_ppt_add_business_cards` | **可用**：缩放正确，超栏时有明确报错（`卡片数量超过分栏数，请拆分为多页，避免内容被截断`）；独立调用正常 | [02-ppt.md](mcp-sweep/02-ppt.md) §142、§241 |
| 8 | `wps_ppt_insert_native_chart` | **100% 失败**且文案误导：`AddChart/AddChart2/AddOLEObject` 都是 function 但返回 `null`、不建形状，报的却是"数据配置未完成"；`generate_deck` 的 chart 布局同样失败且**已插入的页不回滚** | ISS-80、ISS-81；[02-ppt.md](mcp-sweep/02-ppt.md) §2.5（第 217 行：读回该页 `shapeCount: 0`） |

**结论**：`tool-inventory.md` 表中这 8 行的 `（推断）` 标记**可以去掉**，改为指向上述 ISS 条目。7 号是**澄清为"没问题"**，其余 7 个是**澄清为"有确定的缺陷"**。

**残留**：`wps_word_write_content` 吞字母的**根因**仍在隔离中（现象与影响面已确认，属于"根因待定"，不是"未确认能不能用"）。

---

## 9. 审计记录条数上限的副作用未确认 —— ❌ 仍未确认

**原文**：`audit-store.ts:47-49` 只保留最新 500 条，超过后旧记录被丢弃；这对"回滚可用性"的实际影响未验证。

**当前状态**：**仍未确认**。此后没有任何一轮测试触及 500 条上限——P5 期间审计库只有 **34 条** `patch_cells` 留痕（含 1 次真实回滚），距离上限极远；此后所有回滚相关的条目（ISS-46 / ISS-47 / ISS-48）讨论的都是**判定逻辑**，不是**容量淘汰**。

**已复核的代码事实**（这部分不算未确认）：

```ts
// src/bridge/audit-store.ts:47-48
// records 使用最新记录在前的顺序，只保留最新 500 条。
this.records = this.records.slice(0, 500);
```

即：**淘汰是静默的**——没有告警、没有"被淘汰条数"的计数暴露、`wps_get_audit_history` 也看不出"曾经有更早的记录"。

**仍未确认的具体子问题**：
1. 被淘汰的 `auditId` 再调 `wps_rollback` 时，报的是"记录不存在"还是别的措辞，**文案是否可操作**（是否告诉调用方"可能因容量淘汰"）；
2. 淘汰是否**只影响回滚**，还是也会让"后续修改判定"（ISS-46 的内容相等判据）失去历史依据；
3. 500 这个数字**够不够**：一次批量盘点（如本次 3 张成品表的构建）会产生多少条记录；
4. `wps_clear_audit_history`（整库清空、不可回滚、无二次确认，见 `tool-inventory.md:138`）与容量淘汰叠加后的**不可逆风险**。

**下一步怎么确认**（可执行步骤）：
1. **不需要真跑 500 次工具**：`audit-store.ts` 是纯本地实现，可以在隔离的 Node 脚本里直接构造 501 条记录（或把上限临时改小到 3 条，在**临时副本**里验证，禁止改工作区源码），断言第 501 条写入后第 1 条消失、且 `getHistory` 无任何"已淘汰"信号。
2. 在真实宿主上验证**用户可见的那一半**：用一个小上限的构建写 3 条 `patch_cells`，再用最早那条的 `auditId` 调 `wps_rollback`，**逐字记录**返回文案。
3. 判定标准：若文案说清"记录不存在/可能已被淘汰"并给出替代路径 → 澄清（结论=行为可预期但需在文档说明）；若返回 `success` 却什么都没恢复 → **升级为新的高危条目**。
4. 顺带结论：容量策略是否需要在 `wps_rollback` 的说明里写明（与 ISS-28 的"恢复范围"说明是同一条说明的补充）。

---

## 附：本轮判定用到的证据文件

| 文件 | 用在哪条 |
|---|---|
| [tool-inventory.md](../tool-inventory.md) §五 | 9 条原文（本文件的输入） |
| [mcp-sweep/README.md](mcp-sweep/README.md) | #1（方法/成品/失败模式）、#2（版本说明） |
| [mcp-sweep/01-inventory.md](mcp-sweep/01-inventory.md) | #1（只读实测 11 个、70 个写工具未验）、#8 |
| [mcp-sweep/02-ppt.md](mcp-sweep/02-ppt.md) | #2（版本行为指纹）、#4（21 s 超时锚点）、#8（PPT 两行） |
| [mcp-sweep/05-e2e.md](mcp-sweep/05-e2e.md) | #1（18/18 跨组件一致性） |
| [mcp-sweep/07-tool-gap.md](mcp-sweep/07-tool-gap.md) | #5、#8（数据验证） |
| [mcp-sweep/08-ms-vs-wps.md](mcp-sweep/08-ms-vs-wps.md) | #7（双通道对照） |
| [compatibility-diff.md](../compatibility-diff.md) | #3（42 项论证）、#7 |
| [issues.md](../issues.md) | 全部（派生条目 ISS-04/05/41/44/57/58/59/64/67/72/80/81/93/94/97/98） |
| [final-review.md](../final-review.md) | #3 |
| 源码复核 | `src/bridge/audit-store.ts:47-48`（#9）、`wps-addon/src/excel.js` freeze/save（#8、#2） |
