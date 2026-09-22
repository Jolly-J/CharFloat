# Office Agent Bridge 项目改造计划与执行台账

状态：**P0 通过（环境部分阻塞）、P1 通过、P2.1–P2.4 + DP4 完成（P2.5/P2.6 部分完成）、P3 代码拆分与静态/模拟验证完成、P4 完成；P5.1/P5.3 部分完成、**P5.2 完成**——macOS 与 Windows 双平台发布包均已产出，Windows 包经使用者实机运行，mac 与 win 构建物一并通过使用者验收；P5.4–P5.7 真实宿主业务矩阵、加载项安装升级与最终汇总待做。**
日期：2026-09-22（P0/P1 及评审补正于同日更新）。
范围：保留现有技术栈、对外接口和正常业务行为，改善模块边界、契约一致性、验证与 AI 协作。

当前候选标识：`2.1.0-p0p1`，证据目录 [docs/acceptance/2.1.0-p0p1/](acceptance/2.1.0-p0p1/)（基线、可恢复补丁、工具清点、环境、兼容性差异与故障注入日志）。

## 1. 最终目标与完成定义

目标不是按行数拆文件，而是让每个模块承担清晰职责、依赖方向可理解、关键逻辑可独立验证。AI 应能沿任务导航找到少量相关文件，完成局部修改，并通过自动检查发现漏改。

最终必须同时满足：

- [ ] 核心模块按职责拆分，外部工具名、兼容入口、文档操作语义保持兼容。
- [ ] 工具契约、运行时能力和说明一致，区分“已实现”“当前可用”“实机已验证”。
- [ ] macOS 与 Windows 的目标发布包均构建成功，并核验包内资源。
- [ ] 发布包在对应系统实际启动、连接、执行、保存和重启复验通过。
- [ ] 既有业务回归通过，正常业务无回退；故障场景明确报错且无重复写入、误改目标或数据丢失。
- [ ] 导航、避坑指南、测试和发布证据齐备；所有必验项有证据，不能用“未运行”代替通过。

“运行无异常”指本计划规定的平台、宿主版本和场景内，无未处理异常、崩溃、挂起或错误业务结果，不承诺覆盖所有未来版本与任意输入。

## 2. 基线与边界

### 已知起点（执行前重新确认）

- 已存在根导航及模块 AGENTS.md、任务型 skill，以及 release 按版本/平台分类规则。
- 当前工作区有前序未提交修改；必须保留并纳入基线，不使用 git reset/clean 清理。
- gateway、WPS 核心、Office.js 任务窗格存在职责集中，类型与 schema 分散维护。
- catalog → gateway → adapter → catalog 存在依赖回环。
- 上次检查发现 ws-server 中 appendServiceLog 未定义；先确认再修复，不将此历史记录作为当前测试结果。
- Office.js 任意异常后 Windows 尝试 COM，需验证并限制不安全重试。
- 既有回归以模拟为主；历史“测试通过”不能直接算此次发布验收。

### 不纳入本次改造

不引入微服务或大型 monorepo 框架，不重做 UI，不新增办公业务，不批量升级依赖，不改变已有文件格式和保存习惯。不自动提交、推送、部署或签名发布。

修复重复写入、错误能力声明等缺陷时，保持正常场景兼容，记录故障场景前后行为；若需要改变正常产品行为、参数或删掉能力，单列决策点，不能混入结构重构。

## 3. 目标结构与依赖方向

以下为目标职责示意，最终名称可以在职责不变的前提下调整。源码模块化与宿主部署形式分开：加载项仍可构建为现有入口所需的单文件。

```text
src/bridge/
  contracts/          工具契约、类型、错误分类、能力元数据；不依赖执行层
  tools/              excel、word、ppt、native、audit 等处理器
  office/             宿主路由与适配、Office.js 参数转换、原生驱动
  transport/          按实际需要抽离 HTTP/MCP/WebSocket 协议逻辑
  catalog.ts          注册、校验、调度入口，保留兼容导出
  gateway.ts          迁移期兼容门面，逐步移出业务与声明
wps-addon/
  src/                连接、分发、excel、word、ppt、shared
  addon-core.js       构建生成的部署入口；迁移前仍为源码
office-addon/
  src/                Office 生命周期、连接、分发、Excel 操作
  public/taskpane.js  构建生成的部署入口；迁移前仍为源码
tests/
  contracts/          schema、处理器、宿主映射一致性
  unit/               布局、转换等可独立运行的逻辑
  integration/        协议、审计、服务生命周期
  acceptance/         真实宿主验收与测试数据
docs/
  refactoring-plan.md 本计划及执行台账
  validation.md       可复用验证流程
  acceptance/         每次候选版本验收记录与证据索引
```

- 执行层依赖契约层，契约层不引用 gateway、adapter 或服务实例。
- 工具注册关联 schema、处理器和宿主支持信息；类型尽量从同一份 schema 推导。优先评估项目现有校验/依赖，不同时维护两套互相独立的权威 schema。
- 技术能力元数据和连接状态共同决定当前可用性；实机验证记录单独维护，不从“实现方法名存在”推导通过。
- 宿主差异显式留在适配器，不为了统一而掩盖不支持、部分失败和不同返回语义。
- 文件拆分按职责和可测试性，不规定机械行数上限，不制造大量仅转发一行的模块。

## 4. 分阶段完成清单

每阶段独立验收。拆分先保持行为，再做已确认的修复，避免将搬迁和业务变化混在同一批次。

### P0：冻结基线与验收环境

- [x] P0.1 记录源码 HEAD、工作区差异清单和文件指纹，区分前序修改与本次新增；不得只记录 HEAD 代表脏工作区。
      → `docs/acceptance/2.1.0-p0p1/baseline.md` §1（HEAD `cac6d38`、tracked 11 + untracked 16 指纹表、前序/本次分离；`package.json` 基线值经还原重算并已注明）。**评审补正后不再只有指纹**：可恢复内容另存为 `baseline/baseline-prior.patch` + `MANIFEST.sha256`，`verify-baseline.mjs` 实测从**锚点提交**恢复 22/22 一致；5 个未跟踪且已改动文件如实登记为不可恢复
- [x] P0.2 导出实际 tools/list、入参 schema、关键响应和业务样例，建立兼容快照。
      → `scripts/snapshot-tools.ts`（只读，从真实 `getTools()/capabilities()` 导出）；`tools-snapshot.baseline.json`、`tools-snapshot.p1.json`、`business-samples.json`（模拟宿主真实 RPC 往返）
- [x] P0.3 清点全部对外工具及宿主映射，标明读取、写入、部分成功与审计覆盖。
      → `docs/acceptance/2.1.0-p0p1/tool-inventory.md`：91 行主表 + 差异 D1–D9 + 同源发现 A1–A4 + 未确认项 9 条
- [x] P0.4 重跑类型、构建、测试，记录既有失败和环境限制。
      → `baseline.md` §3：typecheck **失败退出码 2**（`ws-server.ts` 4×TS2304）、test 25/25 通过、build 与 build:website、check:agents 通过；环境限制含 Node 25 vs CI 22
- [ ] P0.5 准备独立测试文件、模拟宿主与对应系统的桌面环境；记录 Office/WPS 版本。
      → **部分完成**。测试文件（9 个 .ts / 2 个 .ps1）、模拟宿主（`tests/addon.test.ts` 夹具）齐备；WPS 12.1.28496、Microsoft Excel 16.113 已记录；**Windows 桌面环境缺失**（无任何虚拟化与 PowerShell），Microsoft Word/PowerPoint 未安装。详见 `environment.md` §5
- [x] P0.6 确认发布架构矩阵和签名要求。默认保留现有已出现的 macOS arm64、Windows x64、Windows arm64；macOS x64 如属于实际支持范围则补入，不能默默丢弃已有目标。
      → `environment.md`：矩阵确认 macOS arm64 + Windows x64 + Windows arm64（Windows 架构经 **PE 头实测**而非文件名：`-win.zip` = x86-64，`-arm64-win.zip` = Aarch64）。**macOS x64 仓库内查无依据**，已登记决策点并由 **DP1 决议**：本轮沿用这三个目标，Intel Mac 暂不扩大，Rosetta 转译不能代替 Intel 原生验收。签名/公证按 **DP2** 与内部运行验收分开记录，需要凭据时再处理

退出条件：基线可复现，真实验收环境有明确可用性；缺失环境登记为阻塞，不把未来承诺标成完成。
→ **基线可复现已达成**（补丁 + 清单 + 实测恢复）；**环境准备部分阻塞**（Windows 桌面环境缺失、5 个未跟踪基线文件不可恢复），均已登记。

### P1：恢复检查基线并消除明确风险

- [x] P1.1 定位并修复当前类型检查错误，保持最小修改。
      → `src/bridge/ws-server.ts` 既有 `./runtime.js` 导入行补 `appendServiceLog`（该函数由 `runtime.ts:103` 导出，属漏导入而非缺实现）；`npm run typecheck` 退出码 0
- [x] P1.2 建立错误分类：执行前不可用、明确拒绝、执行失败、结果未知。
      → 新增叶子模块 `src/bridge/errors.ts`：`BridgeError`（`kind` / `channel` / `method` / `executed` + `toJSON`）；`unavailable`/`rejected` 默认 `executed='no'`，`failed`/`unknown` 保守默认 `'unknown'`；已接入 `ws-server.ts` 通道边界、`catalog.ts` 入参校验与宿主失败、`office/adapter.ts`
- [x] P1.3 限制自动回退：只有确认前一通道未执行且替代通道支持时才允许写操作回退；超时或结果未知先读回，不跨通道重放。
      → `errors.routeOfficeFailure()` 为唯一判定入口；`mayReplayOnAnotherChannel()` = 方法不改文档 或 `executed==='no'`；`adapter.ts` 拆分“客户端转换 / 宿主调用 / 响应转换”三阶段，响应转换失败不再回退
- [x] P1.4 增加“写入已发生但响应失败”的故障注入，证明不会再次执行。
      → `tests/failure-routing.test.ts`（当前 11 项）。桩替 `bridgeServer.callOfficeAddon` / `MsOfficeDriver.windows` 统计“宿主执行次数”与“原生回退次数”。**修复前复现**：同文件在锚点版 adapter 上 2 项失败（写入被回退执行第二次且返回成功），日志存 `evidence/p1.4-prefix-failure.txt`（该文件"结果签名"为 `tests=11 pass=9 fail=2`）；修复后 11/11 通过
- [x] P1.5 修正能力声明与实际路由、实现和验证状态的冲突。
      → `capabilities()` 属实化（见 `compatibility-diff.md` §3）：删除无依据的“全量 42 项 … (macOS/Windows 统一)”（实际 30 项宿主方法 / 27 项可调用工具）；新增 `declaredNotCallable`（`clear_range`/`insert_dimension`/`rollback_cells`）与 `unimplementedOnHost`（wps: `update_chart`）；microsoft word/ppt 由“Office.js 原生通道”改为如实说明；4 个只读工具的 `readOnlyHint` 由 false 更正为 true。对外工具名/schema/说明零变化
- [x] P1.6 CI 接入导航检查、官网构建及对应依赖安装；保留现有双系统检查。
      → `.github/workflows/verify.yml`：双系统作业新增 `npm run check:agents`；新增 `website` 作业。补正后官网作业改为 `npm ci --prefix website`（DP5 补齐 lockfile 后成立）。各命令本机实测通过；**未在远端 CI 实跑**

### P1 评审补正（2026-09-22）

| # | 补正项 | 结果 | 证据 |
|---|---|---|---|
| 1 | `inspect_api` 移出只读集合与重放集合，补副作用表达式回归测试 | 该工具执行任意表达式、可能改文档，`readOnlyHint` 保持 false 且不可重放；只读计数由初版 22 回正为 **21**（实际更正 4 个工具）。新增测试覆盖“结果未知不得重放 / 确认未执行仍可回退” | `errors.ts`、`tests/failure-routing.test.ts`、`compatibility-diff.md` §2.1 |
| 2 | `routeOfficeFailure` 保留前一通道原始错误与 `executed` | 移除硬编码 `kind='rejected'`/`executed='no'`，改为只追加“未回退：不在原生通道支持列表内”；`unavailable→no`、`failed→unknown`、`unknown→unknown` 三种情形均保留 | `errors.ts`、`compatibility-diff.md` §4.1、新增测试 |
| 3 | 补齐可恢复基线 | 新增 `baseline/baseline-prior.patch`（22 文件）、`MANIFEST.sha256`、`verify-baseline.mjs`；实测恢复后 22/22 哈希一致。**5 个未跟踪且已改动文件如实登记为不可恢复，未重造** | `baseline/README.md`、验证器退出码 0 |
| 4 | 修复前日志纳入可跟踪证据 | `.log` → `.txt`（`*.log` 被 gitignore）；`git check-ignore` 确认不再被忽略；重跑复现，`tests 11 / pass 9 / fail 2`（与证据文件"结果签名"一致） | `evidence/p1.4-prefix-failure.txt`、`evidence/README.md` |
| 5 | 统一阶段状态、更新台账、记录 DP 决议 | 本文件状态行、阶段台账、执行记录与决策点表已同步 | 本文件 |
| 6 | **恢复材料与 HEAD 解耦**（二轮评审） | 新增 `BASELINE.json` 固定锚点提交；恢复与失败复现均改读锚点而非 `HEAD`；README 恢复命令改用从仓库根目录解析的绝对路径（无个人路径）；失败复现在临时副本执行，不再覆盖工作区；新增隔离验证 `verify-after-head-advance.sh` 证明"副本 HEAD 前进后仍 22/22" | `baseline/BASELINE.json`、`baseline/verify-baseline.mjs`、`baseline/verify-after-head-advance.sh`、`evidence/reproduce-prefix-failure.sh`、`evidence/baseline-head-anchor-prefix-failure.txt` |



退出条件：基础检查通过；明确风险有失败前复现或故障注入证据、修复后回归；仍待实机项目明确列出。

### P2：统一契约并解除回环

- [x] P2.1 提取独立工具契约与宿主方法标识，解除 adapter 对 catalog 的反向依赖。
      → 新增契约层 `src/bridge/contracts/host-methods.ts`（宿主方法路由表、宿主实现缺口、只读与重放属性、`bareName` 归一），**不依赖执行层**；`office/adapter.ts` 改从契约层引用，`catalog.ts` 保留同名兼容导出（`EXCEL_METHODS`/`READ_TOOLS`/`HOST_IMPLEMENTATION_GAPS`/`isReadOnlyTool` 与契约层**同一对象**）。新增 `tests/contracts-boundary.test.ts` 做结构守卫：adapter ↛ catalog 已达成、契约层无执行层依赖、回环台账只减不增。
      → **未完成部分**：经 `ws-server.ts` 静态引用 `catalog.ts` 形成的 2 条回环仍在（见下方 P2.2），已列入守卫测试的 `KNOWN_CYCLES_PENDING_P2_2` 台账，消除前不得宣称"回环已清除"。失败前证据见 `evidence/p2.1-prefix-adapter-catalog-cycle.txt`。

- [x] P2.2 建立工具注册机制，关联 schema、执行器、宿主支持与读写属性。
      → 分层落地：**契约层** `contracts/tool-service.ts`（纯接口，无执行器）；**运行时注册层** `tool-registry.ts`（把 catalog 的定义与执行器绑成 `ToolService`）；**组装入口** `compose.ts`（唯一知道谁实现谁，注入服务实例）；**协议层** `ws-server.ts` / `mcp-server.ts` 改为只认注入的 `ToolService`，不再静态引用 catalog。
      → **回环消除**：守卫改为 **Tarjan 强连通分量（SCC）** 完整口径后，测出协议层与执行层原属**一个含 5 个模块、8 条内部边的强连通分量**（不是先前 DFS 口径声称的两条路径）；P2.2 后 SCC 列表为**空**，`KNOWN_CYCLIC_COMPONENTS` 已清空。失败前证据见 `evidence/p2.2-prefix-scc.txt`。
      → 入口点 `cli.ts --serve`、`scripts/inspect.ts`、`tests/service.test.ts` 均调用 `composeBridgeServer()` 装配；未装配即使用 HTTP/MCP 路由会得到明确错误。

- [x] P2.2 补正：**恢复 stdio 会话语义**（评审发现的 P1 级业务回归）。
      → 回归成因：把 proxy 换成 `ToolService` 时，`execute` 只收到 `clientName`，**丢掉了 `sessionId`**；后台按 `body.sessionId` 缺省落到 `http-local`，于是所有 stdio 客户端共用一个会话，文档锁互相覆盖（A 锁定 A.xlsx、B 锁定 B.xlsx 后 A 查到 B.xlsx，A 后续操作会误指向 B 的文件）。同时 `clientName` 被误改为 `MCP Agent`（原为 `stdio MCP`），审计归属也不对。
      → 修法：`sessionId` 与 `clientName` 一起放进契约层 `ToolCallContext`，**随每次 execute 显式传递**；`cli.ts` 远程实现在 `/api/v1/tool/call` 中原样转发；本地实现在 `tool-registry.ts` 用 `requestContext.run` 落地会话与宿主。协议层不再自己判断宿主。
      → 新增 `tests/session-isolation.test.ts`：**两个独立 stdio 客户端进程 → 同一后台**，覆盖 WPS 文档锁与 Microsoft 目标锁的"分别锁定、互查、单方解锁不影响另一方"。修复前 2 项失败（`actual: 'B.xlsx'`），修复后通过。原有生命周期测试保留通过。
      → 证据：`evidence/p2.2-session-isolation-regression.txt`。

- [x] DP4（随 P2 契约整理落实）：`excel_update_chart`（host=wps）提前明确拒绝。
      → `catalog.executeValidatedTool` 在**调用宿主之前**用契约层 `isUnimplementedOn(host, method)` 判定，抛出 `BridgeError(kind='rejected', executed='no')`，并给出替代路径（`wps_execute_script` 或 `host=microsoft`）。Microsoft 侧有实现，不被误伤（未连接时仍为 `unavailable`）。测试同时断言能力声明与行为一致。

- [x] P2.3 逐类迁移，保留 excel_*、wps_*、office_* 的已发布入口和必要别名。
      → 工具**定义**按类拆到 `src/bridge/tools/`：`diagnostics.ts`、`excel.ts`（由宿主方法路由表派生统一 `excel_*`，原 `wps_*` 兼容名并保留）、`audit.ts`，装配点在 `tools/index.ts`；`catalog.getTools()` 退化为装配调用并保留原导出名。新增 `toolClassOf()` 显式分类，无法归类的工具直接抛错。
      → **零兼容差异**：契约快照逐字节不变、91 个工具名与顺序完全一致。搬迁中一度把两个诊断工具的顺序写反，被快照比对当场拦下并修正——装配顺序是有契约意义的。
- [x] P2.3 剩余部分（评审指出后补做）：工具描述符逐类迁移。
      → `getOpenAiTools()` 的 59 条定义迁到 `src/bridge/tools/definitions/`（`excel.ts` 864 / `word.ts` 337 / `ppt.ts` 299 / `microsoft.ts` 93 / `lock.ts` 59 / `script.ts` 53 / `audit.ts` 30 / `index.ts` 51 / `shared.ts` 41 = 1827 行）；`gateway.ts` 1913 → **280 行**（实测；早期台账写 285 为陈旧数字），只留 `executeTool` + `HANDLERS` + `registeredGatewayBranches()` + `TargetLockStore` + `getOpenApiSchema`；`getOpenAiTools(activeWbHint?)` 签名不变、体改为委托。
      → `activeWbHint` 分支完整保留（只影响 28 个 excel 工具的 `workbookName` 描述与 `wps_get_workspace_summary` 的描述三元分支）：带/不带 hint 的描述哈希链迁移前后完全相同。
      → **零兼容差异**：契约快照逐字节一致；独立条数核对 59→59、名称顺序哈希、逐条 schema 哈希链、description 哈希链（plain 与 hinted）全部不变。
      → **归类与已发布顺序的冲突（保持顺序优先）**：`office_lock_target`/`office_unlock_target` 的 `toolClassOf` 是 `lock` 但按前缀聚合在 `definitions/microsoft.ts`；`wps_rollback` 标签是 `audit` 却位于 Excel 区块内部，故 `excel.ts` 导出两个函数、由 `index.ts` 插回原位。标签与定义文件归属**有意不一一对应**。
      → 新增唯一边 `gateway.ts → tools/definitions`；`tools/index.ts` 不再 import gateway，SCC 仍为空。
- [x] P2.4 统一请求/响应边界，缩小边界 any；不以大规模类型重写扩大范围。
      → 新增契约层 `contracts/boundary.ts`：`ToolArgs`（工具入参）、`ChannelParams`（跨通道载荷，要求可 JSON 序列化）、`ToolResult`（返回值），并提供 `isToolArgs`/`asToolArgs` 收敛。
      → 应用到**边界签名**：`adapter.callOffice`、`ws-server.callWps`/`callOfficeAddon`、`normalizer.normalizeOfficeRequest`/`NormalizedRequest`。
      → **刻意保留的宽松点**（并写明理由，避免为了消 any 而做大规模类型重写）：`normalizeOfficeResponse(raw: any)` 是宿主差异的落点，收窄会掩盖不同宿主的返回语义；其返回类型也保持宽松，因为各分支产出的是 `WorkspaceSummary`/`PatchResult` 等具名业务类型，强套统一 `ToolResult` 会要求它们都有索引签名。`normalizeOfficeRequest` 内部的规范化工作副本保留 `any`。
      → `any` 计数 72 → 64（只动边界，不动内部实现）。**注意：本次总验收复验未能复现该口径**——`grep -rn ": any\|<any>\|as any" src/bridge` 得 **62**，台账未记录原始统计方式，引用时以"只动边界"的定性结论为准；typecheck 0；契约快照逐字节不变。
      → 循环依赖部分已在 P2.2 达成（SCC 台账为空）。
      → 本项结束时 `tests/contracts-boundary.test.ts` 的 `KNOWN_CYCLES_PENDING_P2_2` 必须为空数组。
- [~] P2.5 **部分完成**：校验入参转换、返回值及错误，验证“声明了但未实现”和“实现了但未注册”均会被检查发现。
      → 新增 `tests/contract-consistency.test.ts`（初版 **6 项**，后续追加到当前 **10 项**）：①「注册了但未实现」必须为空（每个对外工具都要有执行分支，`excel_*` 按重写规则核对）；②「实现了但未注册」与债务台账一致（当前 6 项死分支：`wps_clear_range`/`wps_eval_code`/`wps_get_style_token`/`wps_ppt_add_chart`/`wps_reload_addon`/`wps_word_capture_preview`，处置见 DP3）；③能力声明与宿主实现缺口同源一致；④`readOnlyHint` 与契约层同源。
      → **审计归属自动化断言已补齐**（评审建议纳入）：①本地路径断言 `clientName`/`host` 写入审计记录；②**真实 stdio 链路**断言（评审指出前一版绕过了 CLI/HTTP/`ToolService`）：两个 stdio 客户端 → CLI 代理 → HTTP → 后台 → 模拟宿主写入 → 经同一 stdio 客户端读回审计记录，断言 `clientName === 'stdio MCP'`。把 `clientName` 改回 `MCP Agent` 可稳定复现失败（`actual: 'MCP Agent'`）。
      → **宿主分支存在性已独立校验**（评审指出后补上）：新增测试直接解析**加载项源码**的 RPC 分支（`wps-addon/src/dispatch.js`、`office-addon/src/rpc.js`），断言 `EXCEL_METHODS` 的每个方法要么在加载项里有分支、要么已在 `HOST_IMPLEMENTATION_GAPS` 声明；同时反向断言"已声明的缺口必须真的缺失"（防声明过期）。实测：唯一缺口 `update_chart` 与声明一致；注入过期声明（把 `read_range` 标为缺口）即失败。**这不再是同源清单比对。**
      → **仍未覆盖**：返回值形状与错误语义未按宿主方法逐一校验。
- [~] P2.6 **部分完成**：MCP 能力资源与工具说明读取共同元数据；skill 和官网中的人工说明核对该来源。
      → **共同元数据已成立**：`catalog.getTools()` / `catalog.capabilities()` 是唯一权威源，MCP 的 `bridge://capabilities` 资源、`bridge_get_capabilities` 工具、HTTP `/api/v1/capabilities` 三处都读它（P2.2 后经注入的 `ToolService`）。
      → **人工说明核对**：新增只读检查 `npm run check:claims`（`scripts/check-capability-claims.ts`），扫描 `skills/`、`website/src/` 与**加载项页面**（`office-addon/public`、`wps-addon`，扫描范围在 2026-09-22 总验收审计后扩展）中"能力数量"表述并与元数据源比对。**当前 27 个文件无矛盾**；注入「全量 42 项」可稳定报错，即 P1.5 修过的那类漂移。
      → **口径已按评审修正**：区分**路由数 30**（`EXCEL_METHODS`，不代表宿主实现）、**对外工具数 91**、**WPS 可调用 26**、**Microsoft 可调用 27**。含糊表述（"全量 N 项"/"N 项结构化能力"）必须落在某个宿主的可调用数上；落在路由数上会被判错并明确提示"30 是路由表长度，不代表宿主实现"。实测「全量 30 项」报错、「全量 26 项」按 WPS 口径通过并给出提示。
      → **已修正的定性矛盾（评审指出，无硬编码数量也有说明漂移）**：`skills/office-agent-bridge/references/capabilities.md` 原写"macOS Microsoft Excel 当前未实现结构化工具"、`src/bridge/mcp-server.ts` 原写"Microsoft Excel 结构化工具目前面向 Windows COM"，均与当前 Office.js 优先 / Windows COM 回退的代码路径不符，已改为如实描述（macOS 与 Windows 同一 Office.js 通道，COM 仅 Windows 回退；本候选版本未实机验收）。
      → 已接入 CI 双系统作业；`skills/AGENTS.md` 写明元数据源。
      → **未完成**：定性表述（宿主支持范围、验证程度）仍**无自动检查**，只能人工核对；本轮只修了已发现的两处。

退出条件：契约快照无未经批准的不兼容差异；核心依赖回环消除；自动契约测试通过。

### P3：拆分业务实现与建立加载项构建

- [x] P3.1 按 Excel、Word、PPT、脚本和审计拆分 gateway，兼容门面保持稳定。
      → `src/bridge/gateway.ts` 2763 → 2037 行：`executeTool` 只做"锁定注入 → 查 `HANDLERS` 注册表（68 条）→ 组 `GatewayContext` → 调用"；处理器按类拆到 `src/bridge/gateway/{excel(481),word(176),ppt(156),audit(97),microsoft(69),script(68),lock(50),locks(46),types(55)}.ts`（合计 1198 行）。`getOpenAiTools`/`getOpenApiSchema` 逐字节未变；`UniversalGateway` 与 `TargetLockStore` 的公开签名逐字未变。
      → **测试被削弱的问题由我修正**：P3.1 因 `case` 是保留字而把 68 条分支名以**注释**保留在门面里，而 `tests/contract-consistency.test.ts` 原本正则抓 `case "..."`——那样测试只是在"校验注释"。已改为导出真实注册表键 `registeredGatewayBranches()` 并让测试直接读取；注释清单另由模块加载时的 `verifyBranchManifest()` 与注册表双向自检。
      → 验证：typecheck 0、test 56/56、契约快照逐字节不变、SCC 仍为空；独立保真校验（69 项按源行号逐条比对）仅 3 处已声明的改写。
      → 未验证：真实 WPS/Microsoft 宿主（仅模拟通道）；处理器体为等量左移缩进，非逐字节同缩进。
- [x] P3.2 WPS 拆分连接/分发、目标定位及三类文档操作；布局计算与宿主写入分开。
      → `wps-addon/addon-core.js`（5170 行）改为构建生成物（5247 行）；源码拆为 `wps-addon/src/` 9 个模块：`shared`/`connection`/`ribbon`/`dispatch`/`excel(1731)`/`word(935)`/`ppt-layout(51)`/`ppt(1218)`/`bootstrap`。
      → **布局计算与宿主写入已分离**：`ppt-layout.js` 零宿主 API、可 Node 直接 import（`@build-strip` 导出块构建时移除）；宿主写入在 `ppt.js`。
      → 验证：`node --check` 0、构建幂等（3 次同哈希）、`tests/addon.test.ts` 2/2、`service.test.ts` 5/5、干净重建通过；acorn 顶层语句多重集 120/121 逐字节相同，唯一差异是 `fitGeneratedPptShapes` 改为调用 4 个新抽出的纯函数。**未真机验证。**
- [x] P3.3 Office.js 拆分生命周期、RPC 与 Excel 操作，保留 load/sync 语义。
      → `office-addon/public/taskpane.js`（1809 行）改为构建生成物（1871 行）；源码拆为 `office-addon/src/` 19 个片段（`state`/`ui`/`lifecycle`/`connection`/`rpc`/`bootstrap` + `excel/` 下 13 个域文件）。
      → 生成物是原文件的**严格超集**：`diff` 证明 0 行删除/修改、仅新增 62 行且全为注释或空行；`dispatchExcelTool` 77 个 case、58 个函数定义前后集合完全一致；`context.sync()` 61 处、`.load(` 40 处、`Excel.run(` 46 处计数一致。
      → 验证：`node --check` 0、构建幂等、`tests/office.test.ts` 4/4、`platform.test.ts` 6/6、干净重建通过、桩测试 20 条 RPC 序列快照逐行一致。**未实机验证；其余 57 个分支仅有"原样搬迁 + 语法检查"级保证。**
- [x] P3.4 构建链生成原部署入口，确认 HTML、manifest、安装器和打包配置引用一致。
      → 新增 `scripts/build-wps-addon.mjs`、`scripts/build-office-addon.mjs`（纯拼接、幂等、`--check` 只校验不写文件），接入 `npm run build:wps-addon` / `build:office-addon` / `build:addons`，并把 `build:addons` 并入 `npm run build`。
      → 引用一致性已核对：`wps-addon/index.html` 仍加载 `addon-core.js`；`taskpane.html` 的 `./taskpane.js?v=2.0.1`、`manifest.xml` 的加载项地址、`addon-installer.ts` 的 `office-addon/excel/manifest.xml` 路径均未变。
      → **打包决策**：`build.files` 增加 `!wps-addon/src/**` 与 `!office-addon/src/**`，发布包只带部署入口，不带构建源码。
- [x] P3.5 明确源码与生成物策略：迁移到生成入口后禁止手改；从干净目录可以重建，不能依赖旧产物。
      → 两个生成物顶部均有"请勿手改"声明；根 `AGENTS.md`、`wps-addon/AGENTS.md`、`office-addon/AGENTS.md`、`scripts/AGENTS.md` 均写明"改 `*/src/**`，构建刷新入口"；旁路工作区一致校验：删除生成物后 `npm run build:addons` 可重建（两个子任务各自实测）。
      → 构建脚本强制 `src` 目录与模块清单**双向**一致：多一个（未登记）、少一个（缺失）、重复登记三种情况均中止构建。**其中 WPS 侧的「未登记」检测是评审补上的**——原实现只做单向检查，新加 `src/*.js` 而忘记登记会静默漏拼、`--check` 仍返回 0。三种漂移场景已实测 exit 1。
- [x] P3.6 PPT 测试改为直接导入布局模块，移除按字符串截取源码的测试方式。
      → `tests/ppt-layout.test.ts` 重写：纯布局计算从 `wps-addon/src/ppt-layout.js` **直接 import**；宿主适配器 `fitGeneratedPptShapes` 从 `wps-addon/src/ppt.js` 源码片段加载到"已注入布局函数"的 vm 上下文执行（加载源码模块，不再切割生成物）；新增"生成物与源码一致且不含 import/export"断言。9/9 通过。
- [x] P3.7 测试若移动到子目录，同步根测试命令，证明所有原有用例仍被发现和执行。
      → 本轮**未移动**测试文件（仍在 `tests/` 平铺），根命令 `tests/*.test.ts` 覆盖全部；实测用例数 25（P0 基线）→ 56（当前），逐次只增不减、`skipped 0`。新增的两个 stdio 类测试必须在 `npm run build` 之后运行，已写入 `tests/AGENTS.md`。
- [x] P3.8 每迁移一类业务，执行对应契约和宿主回归，核对基线行为。
      → 每次迁移都以**契约快照逐字节不变**为硬门槛（P2.3、P3.1 均通过）；加载项迁移以既有测试为准入门槛（`addon.test.ts`、`service.test.ts`、`office.test.ts`、`platform.test.ts`、`ppt-layout.test.ts` 全部通过），并由子任务各自补充等价性证据（acorn 顶层语句多重集、diff 行级对照、桩测试序列快照）。
      → **未做真实宿主回归**：无 WPS/Excel 实机加载、授权、sync、读回与视觉预览。

退出条件：模块可独立定位和测试，加载项打包与加载通过；行为差异清单中无未解释项。

### P4：协作导航与经验维护闭环

> 完成记录（P4.1–P4.5 全部落实）
> - **P4.1**：根导航表的三行入口同步到拆分后的结构——工具定义 `tools/definitions/` → 装配 `catalog.ts` → 处理器 `gateway/`；WPS 与 Office.js 行改为指向 `*/src/**` 并标注部署入口为**生成物**。各模块 `AGENTS.md` 的文件地图在每次迁移中同批更新（`src/bridge/`、`wps-addon/`、`office-addon/`、`scripts/`、`tests/`）。
> - **P4.2**：新增避坑条目均附测试或证据链接（回退策略→`failure-routing.test.ts`、SCC→`contracts-boundary.test.ts`、会话隔离→`session-isolation.test.ts` + 复现证据、构建漂移→生成物 `--check`、伪 `case` 清单→`contract-consistency.test.ts`）；已失效的旧限制（`KNOWN_CYCLES_PENDING_P2_2` 台账、生成物"仍在源码"等表述）已删除或更新，未保留永久禁令。
> - **P4.3**：核对 skill 与 MCP 描述——`SKILL.md` 已写明批量执行（"优先使用可批量处理且满足需求的结构化工具"）、原生脚本（"专用工具未覆盖不等于不支持…只读探测再批量执行"）、结果验证与**按需保存**（"用户要求落盘或任务交付需要保存时调用对应保存接口并核对结果"）；`mcp-server.ts` 的 INSTRUCTIONS 写明"保存成功、内存修改和回滚覆盖范围是不同状态；只有单元格 patch 的值与公式有审计回滚"。本轮修正了两处与代码路径不符的宿主通道描述（见 P2.6）。
> - **P4.4**：三类定位任务的入口实测可达——①新增工具参数：`src/bridge/AGENTS.md` → `tools/definitions/`（schema）→ `gateway/`（映射）→ `office/normalizer.ts`（转换）→ `tests/contract-consistency.test.ts`（回归）；②修复 PPT 布局：`wps-addon/AGENTS.md` → `src/ppt-layout.js`（纯计算）→ `src/ppt.js`（宿主写入）→ `tests/ppt-layout.test.ts`；③诊断连接：`src/bridge/AGENTS.md` → `service-client.ts` → `ws-server.ts` → `compose.ts`。三组路径均在首读目录内闭合，**无需全项目扫描**。
> - **P4.5**：`check:agents` 已接入 CI 双系统作业；`build.files` 含 `!**/AGENTS.md`（开发规范不进用户运行包）；`skills/**` 保留在包内，`skills/office-agent-bridge/` 的 `SKILL.md` + `references/{capabilities,operations,native-scripting}.md` + `scripts/` 完整。


- [x] P4.14.1 同步根任务导航、局部文件地图、调用链和验证命令，删除失效入口。
- [x] P4.24.2 避坑条目链接到测试或验收证据；已失效的临时限制删除或更新。
- [x] P4.34.3 核对 skill 与 MCP 描述，明确批量执行、原生脚本、部分完成和保存语义。
- [x] P4.44.4 用三类真实开发定位任务检查导航：新增工具参数、修复 PPT 布局、诊断连接。记录首读文件和扩展搜索原因，不能只凭路径检查宣称导航高效。
- [x] P4.54.5 导航检查纳入 CI；开发规范不进入用户运行包，用户 skill 资源完整保留。

退出条件：导航校验通过，代表性定位任务可从明确入口完成，说明与实现核对完成。

### P5：双系统构建、运行及业务验收

- [~] P5.15.1 锁定候选源码和依赖版本，在 macOS/Windows 环境分别执行完整检查与打包。
- [x] P5.25.2 按发布架构矩阵生成 ZIP 与解包目录，检查命名、架构、资源、入口及校验值。
- [~] P5.3 从实际发布包运行，验证桌面、CLI、stdio MCP、HTTP MCP，而非仅从源码启动。（**部分完成**）
      → 已从**包内** `app.asar.unpacked/dist/bridge/cli.cjs` 运行：`--help` 正常、`--serve` 后 `/health` 返回 `{"service":"wps-bridge","protocol":2,"version":"2.1.0"}`、`/api/v1/mcp-tools` 返回 **91 个工具**、提示词资源 HTTP 200、包内 `office-addon/excel/manifest.xml` 就位。
      → **2026-09-22 使用者实机（Windows）**：运行本轮发布包，反馈"win运行正常"，并验收 mac 与 win 构建物 → 打包后的 Windows 桌面应用启动这一项由使用者完成；证据见 [p5.2-windows-cross-build.md](acceptance/2.1.0-p0p1/p5/p5.2-windows-cross-build.md) 的"使用者实机验收"一节。
      → **仍未做**：macOS 侧打包后 GUI 启动、包内 stdio MCP 客户端握手、双客户端共存验收。
- [~] P5.4 完成下方业务矩阵及故障矩阵，保留结构化读回、日志和必要截图。（**macOS + WPS 部分完成**）
      → 2026-09-22 在使用者指定的测试工作簿上执行：**只读路径全通过**（工作区摘要、6 表结构、读区域含公式与数字格式、图表清点、高保真截图）；**写入路径**通过建表、格式化、列宽、冻结、数据验证、插图与读回、保存；**审计留痕与 `wps_rollback` 实测可用**；**P5.6 落盘核验首次拿到真实证据**（解包 xlsx 命中写入标记）。
      → 写入全部限定在新建测试表内，原有 6 表按「值+公式+数字格式」逐单元格校验和前后一致（见证据 §4）。
      → **发现真实缺陷**：`excel_patch_cells` 混用 `values` + `formulas` 时静默清空数据 → 登记为 **DP7**，未擅自修改。
      → **未做**：条件格式/排序筛选/透视表/行列操作/查找替换/批注、Word 与 PPT 矩阵、回滚负向用例、Microsoft Excel（Office.js）通道、Windows 侧。
      → 证据：[p5.4-macos-wps-business-matrix.md](acceptance/2.1.0-p0p1/p5/p5.4-macos-wps-business-matrix.md)
- [ ] P5.5 测试加载项安装/升级、保留其他配置、后台重启与双客户端共存。
- [ ] P5.6 保存后关闭测试文件并重新打开，核验实际落盘内容；不只检查内存数据。
- [ ] P5.7 汇总所有目标平台结果、遗留项和证据，满足最终门槛才标记完成。

退出条件：所有必验平台/架构/业务通过，无未解决的破坏业务、重复执行或误写问题。缺少真实 Windows 或某架构设备时，该项待验收，不能用交叉打包代替运行通过。

## 5. 平台与业务验收矩阵

每个候选版本在 docs/acceptance/<候选标识>/ 中维护实际矩阵，不能只修改本表的通用标准。

| 层级 | 必验范围 | 通过标准 |
|---|---|---|
| 静态与单元 | 类型、契约、布局、转换、导航 | 无新增/未解决检查失败，原有用例未被跳过 |
| 构建 | 桌面、后台 CLI、加载项、官网 | 干净构建成功，产物不依赖旧 dist 文件 |
| 发布包 | P0 确认的 macOS/Windows 架构 | 解压完整、CPU 架构正确、入口/资源可定位、SHA-256 可核对 |
| 启动 | 对应系统实际发布包 | 主界面、托盘、日志、CLI 正常，无未处理异常 |
| 服务 | stdio/HTTP、双客户端、GUI 退出、后台重启 | 会话和凭据有效；退出 GUI 不误停后台；重连和失败明确 |
| WPS Excel | macOS/Windows | 范围读写、零值与公式、样式、结构、图表、矢量形状、截图、保存重开正确 |
| Microsoft Excel | Office.js；Windows 原生分支 | 参数/返回转换正确，实际路线有证据；回退不会重复写入 |
| WPS Word | macOS/Windows | 新建、段落/表格读写、格式、批量查找替换、保存重开；非目标内容保留 |
| WPS PPT | macOS/Windows | 页面/形状/表格/原生图表、脚本、预览与保存；不同页面尺寸布局正确 |
| Microsoft Word/PPT | 实际声明支持的原生脚本及预览路线 | 精确目标、执行读回、保存/预览按实际工具能力核验；不宣称与 WPS 全功能等同 |
| 安装与升级 | 两系统及已支持客户端 | 保留其他配置、损坏文件不覆盖、路径含中文/空格可用 |
| 审计 | patch 与回滚 | 快照一致时恢复；有后续修改则拒绝覆盖；不夸大覆盖范围 |

业务清单必须从 P0 全量工具盘点补全：每个对外工具至少有契约用例，每个已声明支持的宿主操作有执行验证记录。同类操作可批量验收，但结果按工具留存。未支持项应明确拒绝，不能把缺失实现标为回归通过。

### PPT 与文件业务的具体断言

- [ ] 720×405、960×540、1440×810、720×540 页面：几何按比例正确，字体不因默认坐标集中左上角。
- [ ] 短文本、长中文、混合语言、换行及已有对象：无静默截断、无无关对象移动；溢出警告可定位。
- [ ] 每个测试页有真实预览；表格和图表类型及数据正确，截图不能代替可编辑对象。
- [ ] 多文档同开且切换焦点时，操作仍命中指定文件；只读/保护状态明确失败。
- [ ] 修改前后比对目标范围及关键非目标区域，保存重开后公式、文字、对象和格式保持。

### 故障与恢复断言

- [ ] 未连接、错误目标、参数非法：执行前拒绝，无附带修改。
- [ ] 执行中断、超时、响应丢失：报告未知或部分完成，不自动跨通道重放。
- [ ] 图表/形状创建后配置失败：定位已有对象续作，无重复对象。
- [ ] 服务重启、加载项重连、客户端退出：状态可恢复，无凭据泄露或误停其他客户端。
- [ ] 异常路径日志可定位原因，同时不包含 token 或用户文档全文。

## 6. 构建与发布证据要求

现有命令以 package.json 为准：`npm run check:agents`、`npm run typecheck`、`npm test`、`npm run build`、`npm run build:website`。执行改造时补齐加载项构建及发布矩阵命令，不能在尚未实现前声称可用。

打包可通过现有 dist 脚本传 electron-builder 参数，例如 macOS 主机 `npm run dist -- --mac --arm64`，Windows 主机 `npm run dist -- --win --x64`；其他目标架构独立构建并运行验收。同版本构建会覆盖同名产物，测试前保留候选标识与校验值，避免混用旧包。

每项证据记录：候选源码指纹、依赖锁文件指纹、系统/架构、宿主版本、工具清单快照、命令、退出码、开始结束时间、产物路径与 SHA-256、运行步骤、读回结果、截图或日志链接。测试文件仅用独立样例，不上传用户数据或凭据。

默认验收沿用当前 ZIP + 解包形式。签名、公证、企业安全策略和首次运行拦截属于发布条件，P0 明确目标；未提供所需凭据时单独标记，不通过关闭系统保护伪造验收。

## 7. AI 执行与备忘规则

### 每次开始

1. 读根 AGENTS.md、本计划的进度与当前阶段，检查 git status/diff。
2. 根据当前任务 ID 进入目标模块，不重新全项目扫描。
3. 核对上次证据与当前代码是否仍一致；有新修改则只重跑受影响验证。

### 每批完成

1. 更新对应任务勾选框和下方台账，附文件与证据位置。
2. 代码已完成但验证未完成时保持未勾选，在台账注明“待验证”。
3. 记录失败原因、已排除假设、不要重试的操作，以及下一步入口。
4. 更新受影响 AGENTS.md/skill；不在文档中仅写“已完成”而无验证依据。
5. 发现新增重大决策，仅讨论该变化；不要重开已确认路线。

### 阶段台账

| 阶段 | 状态 | 代码/说明变更 | 验证证据 | 阻塞与下一步 |
|---|---|---|---|---|
| P0 | **基线盘点通过**；环境准备部分阻塞 | 新增只读快照导出脚本、可恢复基线补丁与验收证据目录；未改业务行为 | `acceptance/2.1.0-p0p1/`：baseline.md、baseline/、tool-inventory.md、environment.md、两份快照、业务样例 | Windows 桌面环境缺失（阻塞 P5）；5 个未跟踪基线文件不可恢复；待评审确认 |
| P1 | **评审通过**（首轮 5 项 + 二轮 1 项补正已落实） | 新增 `errors.ts`；`ws-server.ts` 补导入+分类错误；`adapter.ts` 三阶段拆分+回退策略；`catalog.ts` 声明属实化；CI 加导航与官网 | typecheck 0、test 35/35、build 0、build:website 0、check:agents 0、npm ci --prefix website 0、基线恢复 22/22、HEAD 前进后恢复 22/22；`evidence/p1.4-prefix-failure.txt`（修复前 2 项失败） | 待评审确认后方可启动 P2；CI 需在远端跑一次 |
| P2 | P2.1–P2.4、DP4 完成；**P2.5、P2.6 部分完成** | 契约层 `contracts/`、注册层 `tool-registry.ts`、组装入口 `compose.ts`；协议层注入式；SCC 守卫；DP4 前置拒绝；会话上下文；契约一致性测试；能力声明检查 | typecheck 0、test 全部通过、build 0、build:website 0、check:agents 0、check:claims 0；SCC 为空；契约快照逐字节不变 | **P2.5/P2.6 剩余**与 **P3 宿主验收**待做；DP3 仍暂缓 |
| P3 | **代码拆分与静态/模拟验证完成；宿主验收待完成** | gateway 按类拆分 + 注册表；两个加载项拆为 src + 构建生成入口；构建脚本双向校验；PPT 测试改直接导入 | typecheck 0、test 全部通过、build 0、契约快照逐字节不变、SCC 为空、生成物 --check 一致、构建漂移三场景 exit 1 | 实机与发布包验收仍缺（P5） |
| P4 | **P4.1–P4.5 完成** | 根导航与模块地图同步到拆分后结构；避坑条目挂测试/证据；skill 与 MCP 语义核对；三类定位任务实测可达；导航检查在 CI、开发规范不入包 | `check:agents` 0；三组入口路径均可达；`!**/AGENTS.md` 与 `skills/**` 打包配置核对 | — |
| P5 | P5.1/P5.3 部分完成、**P5.2 完成**：双平台包均已产出，Windows 包经使用者实机运行并连同 mac/win 构建物一并验收；**P5.4–P5.7 未做** | 打包命令与包内核验；候选与依赖锁记录；Windows exe 图标条目格式修复（<256 用 DIB、256 用 PNG、64KB 硬上限） | mac：`npm run dist -- --mac --arm64` 退出码 0、架构 arm64；win：x64/arm64 两个 zip，PE 架构与版本资源实测，图标 7 条目全为 DIB(<256)+PNG(256) 且组目录长度无截断；包内 91 工具、资源完整、无开发规范与 addon 源码 | 使用者实机：Windows 包运行正常、构建物验收通过；**真实宿主业务矩阵、加载项安装/升级、保存后重开核验未做**；Windows 侧未跑完整检查（typecheck/test） |

> 本表记录**各阶段收尾当时**的状态与计数，属于历史快照；数字随时段漂移过（测试数 53→56→73 等），**当前权威数字以 [final-review.md](acceptance/2.1.0-p0p1/final-review.md) §1「本次复验结果」为准**。2026-09-22 总验收前已把与证据文件冲突的数字逐条更正。

### 当前交接备忘

- 当前任务：P0 通过（环境部分阻塞）、P1 通过；**P2.1–P2.4、DP4 完成**，P2.5/P2.6 部分完成；**P3.1–P3.8 代码拆分与静态/模拟验证完成，宿主验收待完成**；**P5.1/P5.2/P5.3 部分完成——双平台包已产出，Windows 包经使用者实机运行并连同 mac/win 构建物一并验收**。
- 下一个任务（按优先级，均为未完成项）：
  1. **P2.5 剩余**：描述符已就位（`tools/definitions/` + `toolClassOf` + `HANDLERS`），可基于它补"宿主分支是否存在 / 返回值形状 / 错误语义"的检查。已覆盖：注册表分支存在、能力声明同源、只读同源、Microsoft 27 个方法的参数/响应转换边界可跑通。
  2. **P2.6 剩余**：定性表述只有「已证实错误表述」黑名单（3 条），其余仍需人工核对。
  3. **P3 验收缺口**：加载项"加载通过"与宿主回归**无实机证据**——WPS/Excel 未真机加载、授权、sync、读回、预览；Office 侧 77 个分支中桩测试仅覆盖 20 个；WPS 侧 `fitGeneratedPptShapes` 改为调用 4 个新抽出的纯函数（非逐字节等价）。
  4. **实机阶段测试**：`npm run dev` 由使用者执行；前置构建已校验（`build:main` / `build:addons` / `build:renderer` 均退出码 0，dist 产物就位，vite 与 Electron 可用）。改 `*/src/**` 后需重跑 dev 或 `npm run build:addons` 再重载加载项。
  5. **P5 剩余**（P4 已完成，不再是"下一个任务"）：按 **P5.4 业务矩阵 → P5.6 保存后重开 → P5.5 加载项安装升级 → P5.7 汇总** 推进；需要使用者提供**隔离的测试文件**与真实 WPS/Excel 环境。
- 已确认方向：现有技术栈内渐进模块化、契约统一、双平台构建与真实业务验收。
- 已确认事实：基线 typecheck 失败而 build 通过（tsup 不做类型检查）；91 个对外工具；`EXCEL_METHODS` 30 项 / 可调用 27 项；`inspect_api` 执行任意表达式，既非只读也不可重放。
- 待核实：CI 远端首跑结果；**每份发布包与源码指纹的对应关系**（三份包在不同时间产出，无逐包对应记录，见 P5.1）。**本机（macOS）仍无 Windows 桌面环境，P0.5 的阻塞判定不变**；Windows 实机运行由**使用者侧**完成，见 P5.2 的"使用者实机验收"。凭据/签名类问题按 DP2 与 P5 一并处理。
- 禁止误判：**代码改造侧**的验证仍是静态检查、构建与模拟宿主层级，**没有真实宿主业务验收**；**发布包侧**已有使用者的 Windows 实机"可运行 + 构建物"验收，但**业务矩阵（P5.4–P5.7）没有任何实机证据**。73 项测试通过不代表业务可用；"构建通过"不等于"业务验收通过"。

### 决策点与决议（DP1–DP6）

| 编号 | 决策点 | 决议 | 落实位置 |
|---|---|---|---|
| DP1 | macOS x64 是否纳入发布矩阵 | **本轮沿用 macOS arm64、Windows x64、Windows arm64；Intel Mac 暂不扩大。Rosetta 转译不能代替 Intel 原生验收。** 不擅自补入也不擅自丢弃 | `environment.md` 发布架构矩阵；本文件 P0.6 |
| DP2 | 签名与公证 | **分开记录内部运行验收与正式分发要求**；需要凭据时再处理，内部验收不以签名/公证为门槛 | `environment.md` 签名与分发要求一节 |
| DP3 | 为 `clear_range`/`insert_dimension`/`rollback_cells` 新增对外工具 | **保持暂缓**；P2 先整理已有契约，再评估是否补工具 | P2.2/P2.3 |
| DP4 | `excel_update_chart`（host=wps）提前明确拒绝 | **属于明确错误处理，可在契约整理中落实并测试，无需设计多套方案** | P2 契约整理（P1 只改声明，未改行为） |
| DP5 | `website/` 缺 lockfile | **已补齐**：生成 `website/package-lock.json`（193 包），核对后与根 `node_modules` 现有版本**零差异**；`npm ci --prefix website` 与官网构建均实测通过；CI 改用 `npm ci` | `website/package-lock.json`、`.github/workflows/verify.yml` |
| DP6 | 发布包命名 | **已确定使用包含版本、平台、架构的新命名；旧包保留原名，无需重新选择** | `package.json` 的 `artifactName`（现值即符合）；P5.2 按此核验 |
| **DP7** | `excel_patch_cells` 同时传 `values` 与 `formulas` 时，`formulas` 里的 `''` 会把该单元格清空（第二步覆盖第一步），而返回文案是"已成功修改 N 个单元格" → **对调用方是静默数据丢失**（详见 [p5.4](acceptance/2.1.0-p0p1/p5/p5.4-macos-wps-business-matrix.md) §3） | **待决策**。三个方向：**A（推荐）** `formulas` 中 `''`/`null` 视为"不动该单元格"，只对非空项赋公式；**B** 矩阵含空项时直接报错，要求调用方二选一；**C** 保持现状，只在工具说明与 skill 里写明"不要混用"。A 最贴合直觉且不改正常路径；B 最保守但会让既有调用方式变严格；C 零改动但把风险留给调用方 | 决策后改 `wps-addon/src/excel.js` 的 `patchCells` + `tests/addon.test.ts` 回归 + 工具说明 |

> **总验收期间发现的问题统一登记在 [问题台账](acceptance/2.1.0-p0p1/issues.md)**（ISS-01 起，逐条修复并更新状态）。DP7 对应其中的 ISS-01。


### 执行记录

| 日期/任务 ID | 实际变更 | 检查与结果 | 证据 | 未完成与下一步 |
|---|---|---|---|---|
| 2026-09-22 / P0.1 | 记录 HEAD、工作区差异与 27 个文件指纹，分离前序与本次修改 | `git rev-parse`、`git diff --name-only`、`git ls-files --others`、逐文件 SHA-256 | `acceptance/2.1.0-p0p1/baseline.md` §1 | `package.json` 基线值经还原重算，已在文中注明 |
| 2026-09-22 / P0.2 | 新增只读快照脚本 `scripts/snapshot-tools.ts` + `snapshot:tools` 命令；导出两份工具快照与业务样例 | 两次导出 91 工具、只读计数一致；脚本在重名时抛错 | `tools-snapshot.baseline.json`、`tools-snapshot.p1.json`、`business-samples.json` | 契约测试自动化留在 P2.5 |
| 2026-09-22 / P0.3 | 91 个工具逐行清点宿主映射、读写、部分成功与审计覆盖 | 与快照逐名比对：无缺失/重复/多余 | `tool-inventory.md` | 9 条差异中 D1/D2/D3/D6/D7 已在 P1.5 处理；其余为声明层待议 |
| 2026-09-22 / P0.4 | 重跑类型、构建、测试、官网构建、导航检查 | typecheck **失败退出码 2**（4×TS2304）；test 25/25；build 0；build:website 0；check:agents 0 | `baseline.md` §3 | 关键事实：build 通过不能证明类型正确 |
| 2026-09-22 / P0.5 | 清点测试资产、模拟宿主、宿主版本与桌面环境 | WPS 12.1.28496、Excel 16.113；Word/PPT 未安装；无任何 Windows 环境 | `environment.md` | 部分完成，Windows 侧登记阻塞 |
| 2026-09-22 / P0.6 | 确认发布矩阵与签名要求 | Windows 架构经 PE 头实测（`-win.zip`=x86-64、`-arm64-win.zip`=Aarch64）；macOS x64 无依据 | `environment.md` | DP1/DP2 待决策 |
| 2026-09-22 / P1.1 | `ws-server.ts` 补 `appendServiceLog` 导入（最小改动） | `npm run typecheck` 退出码 0 | `git diff src/bridge/ws-server.ts` | — |
| 2026-09-22 / P1.2 | 新增 `src/bridge/errors.ts`（叶子模块，无依赖）；通道边界与入参校验接入分类 | typecheck 0；既有 25 项全通过，消息文本基本未变 | `compatibility-diff.md` §4 | 其余错误出口在 P2.4 统一边界时复核 |
| 2026-09-22 / P1.3 | `adapter.ts` 拆为三阶段；`routeOfficeFailure` 为唯一回退判定 | 决策表测试 + 三组故障注入通过 | `tests/failure-routing.test.ts` | Windows 实机需在 P5 复验 |
| 2026-09-22 / P1.4 | 故障注入“写入已发生但响应失败” | 修复前：**11 项中 2 项失败**（回退执行第二次并返回成功）；修复后 **11/11** 通过 | `evidence/p1.4-prefix-failure.txt`（签名 `tests=11 pass=9 fail=2`） | — |
| 2026-09-22 / P1.5 | 能力声明属实化；4 个只读工具 `readOnlyHint` 更正 | 工具数 91→91、schema/说明零变化、仅 4 处 `readOnlyHint` 变化 | `compatibility-diff.md` §1–§3 | DP3/DP4 待决策，未删除任何能力 |
| 2026-09-22 / P1.6 | CI 双系统作业加 `check:agents`；新增官网构建作业 | YAML 解析校验通过；各命令本机实测通过 | `.github/workflows/verify.yml` | 未在远端 CI 实跑 |
| 2026-09-22 / 评审补正 1 | `inspect_api` 移出只读集合与重放集合；新增副作用表达式回归测试 | `wps_inspect_api` 保持 `readOnlyHint=false`；只读计数回正为 17→**21**（4 项） | `errors.ts`、`tests/failure-routing.test.ts`、`compatibility-diff.md` §2.1 | D6 描述同步收窄 |
| 2026-09-22 / 评审补正 2 | `routeOfficeFailure` 替代通道不支持时保留原始错误与 `executed` | 不再把 `unknown` 改写成 `no`；三种 kind 均保留 | `errors.ts`、`compatibility-diff.md` §4.1、新增测试 | — |
| 2026-09-22 / 评审补正 3 | 补齐可恢复基线 | `baseline-prior.patch`（22 文件）+ `MANIFEST.sha256` + 恢复验证器；实测 22/22 哈希一致 | `baseline/README.md` | 5 个未跟踪文件不可恢复，已如实登记 |
| 2026-09-22 / 评审补正 4 | 修复前日志改为可跟踪证据 | `.log` → `.txt`，`git check-ignore` 确认不再被忽略；重跑复现 `tests 11 / pass 9 / fail 2`（与证据文件"结果签名"一致） | `evidence/p1.4-prefix-failure.txt`、`evidence/README.md` | — |
| 2026-09-22 / DP5 | 补齐官网 lockfile | 生成 `website/package-lock.json`（193 包）；与根 `node_modules` 现有版本**零差异**；`npm ci --prefix website` + `build:website` 均退出码 0；CI 改用 `npm ci` | `website/package-lock.json`、`.github/workflows/verify.yml` | — |
| 2026-09-22 / 评审补正 5 | 统一阶段状态、记录 DP1–DP6 决议、更新台账 | 状态行/阶段台账/本记录已同步；复验：typecheck 0、test 35/35、build 0、build:website 0、check:agents 0、基线恢复 22/22 | 本文件、`baseline.md` §5 | 待评审确认 |
| 2026-09-22 / P2.2 | 分层 + 依赖注入解除协议层与执行层循环依赖；DP4 落实 | 新增 `contracts/tool-service.ts`、`tool-registry.ts`、`compose.ts`；`ws-server`/`mcp-server` 去除 catalog 静态引用；`cli`/`inspect`/`service.test` 装配；DP4 前置拒绝 | typecheck 0、test 42/42、check:agents 0、build 0、build:website 0；SCC 为空；契约快照零变化 | 证据 `p2.2-prefix-scc.txt`；待阶段评审 |
| 2026-09-22 / P2.5（部分） | 契约一致性校验 + 审计归属断言；按评审收窄断言口径 | `tests/contract-consistency.test.ts`：两类漂移 + 声明同源 + `readOnlyHint` 同源 + 审计归属；`session-isolation.test.ts` 增真实 stdio 审计归属 | 51/51 通过；「缺少网关分支」为 0；死分支台账 6 项；把 `clientName` 改错可稳定复现失败 | 宿主分支/参数转换/返回与错误语义**未覆盖**，待 P2.3 |
| 2026-09-22 / P2.6（部分） | 能力元数据同源 + 文档声明检查；按评审区分口径并修定性矛盾 | `check-capability-claims.ts` 区分路由/工具/各宿主可调用数；修正 skill 与 MCP 说明中两处与代码路径不符的描述 | 「全量 30 项」报错并指出 30 是路由数；「全量 26 项」按 WPS 口径通过 | 定性表述仍无自动检查 |
| 2026-09-22 / 评审补正 8 | 构建链双向校验、dev 启动构建加载项、移除重复分支清单、台账纠偏 | `build-wps-addon.mjs` 补未登记/缺失/重复三项检测（三种场景实测 exit 1）；`dev.ts` 增加 `build:addons` 步骤与重建说明；`gateway.ts` 移除为旧正则服务的 `case` 注释清单、`BRANCH_NAMES` 与 `verifyBranchManifest`（2763→1913 行）；P2.3 改标部分完成并写明剩余范围（**该标注是当时快照；P2.3 已在本轮后续补做完成，见上方 P2.3 两条**）；状态行与台账统一 | 新增/缺失/重复三种漂移均 exit 1；typecheck 0、test 56/56、契约快照逐字节不变 | P2.3/P2.5/P2.6 未完成项已登记 |
| 2026-09-22 / P2.3 | 工具定义逐类迁移到 `tools/`，新增显式分类 | 新增 `tools/{index,diagnostics,excel,audit}.ts`；`catalog.getTools()` 改为装配调用；新增 `toolClassOf`；补边界与分类守卫 | 契约快照逐字节不变（91 工具、顺序一致）；typecheck 0；test 53/53；SCC 仍为空 | 执行器拆分留 P3.1 |
| 2026-09-22 / P2.6 尾项 | 数量检查改为宿主感知 | `check-capability-claims.ts` 逐行识别宿主；含糊表述无宿主即报错，不再仅凭数字反推 | 「Microsoft 全量 26 项」报错（应为 27）；无宿主「全量 26 项」报错要求补充宿主；check:claims 0 | 定性表述仍无自动检查 |
| 2026-09-22 / P2.2 补正 | 修复 stdio 会话语义丢失（跨客户端文档锁串扰） | `sessionId`+`clientName` 收进 `ToolCallContext` 显式传递；`cli.ts` 转发、`tool-registry.ts` 落地；协议层去掉宿主路由知识；新增 `tests/session-isolation.test.ts` | 修复前 2 项失败（A 查到 B.xlsx）；修复后 44/44 通过；契约快照零变化；生命周期测试保留通过 | 证据 `p2.2-session-isolation-regression.txt`；待阶段评审 |
| 2026-09-22 / 评审补正 7 | 修复复现脚本变量插值缺陷并改为签名判定 | `$status（` → `${status}`、`exit "$status"`；不再只看退出码，改为核对 `tests/pass/fail` 与两个失败用例名，退出码 0=复现成功 | `evidence/reproduce-prefix-failure.sh`、`p1.4-prefix-failure.txt` | — |
| 2026-09-22 / P2.1 | 提取契约层，解除 adapter 对 catalog 的反向依赖 | 新增 `contracts/host-methods.ts`；`adapter` 改引用契约层；`catalog` 兼容导出与契约层同一对象；新增 `tests/contracts-boundary.test.ts` | typecheck 0、test 40/40、check:agents 0、契约快照逐字节不变；新守卫在旧 import 下报出 `catalog → gateway → adapter → catalog` 回环 | **剩余 2 条经 ws-server 的回环未消除**，已入台账，P2.2 处理 |
| 2026-09-22 / 评审补正 6* | 恢复材料与 HEAD 解耦 | 新增 `BASELINE.json` 锚点；`verify-baseline.mjs` 改读锚点并支持 `--repo`；README 绝对路径化；失败复现移入临时副本；新增 HEAD 前进隔离验证 | `baseline/`、`evidence/`；实测：锚点恢复 22/22（退出码 0）、HEAD 前进后 22/22（退出码 0）、失败复现 1（预期）、工作区 adapter 指纹未变 | 修复前 HEAD 依赖已留证；待评审确认 |
| 2026-09-22 / P5.2 图标 | Windows exe 文件图标空白：两轮定位到条目**编码格式**与 **PE 长度字段截断**两个缺陷 | `scripts/build-win-icon.mjs` 重写为纯 Node：<256 写 32bpp DIB、256 写 PNG、尺寸集合改 16/24/32/48/64/96/256、加 65535 字节硬上限、写完自检不通过不落盘；`npm run dist` 串联 `build:win-icon`；新增 `tests/win-icon.test.ts`（10 项）与只读核验脚本 `inspect-pe-icons.mjs` | 对照原版 `electron.exe`（小尺寸 DIB + 256 PNG）；两个架构 exe 均 7 条目、`<256` 全 DIB、组目录长度与资源实际长度逐条一致（修掉 67624→2088 的截断）；typecheck 0、test 73/73、check:agents 0、check:claims 0、生成器幂等 | 证据 `p5.2-windows-cross-build.md` §图标问题第三轮；修完收到使用者实机确认 |
| 2026-09-22 / P5.2 验收 | 使用者实机运行 Windows 发布包并验收构建物 | 无代码变更（仅记录验收） | 使用者原话："可以，win运行正常，mac和win构建物我都验收" | 覆盖范围仅"包可运行 + 构建物验收"；**P5.4–P5.7 业务矩阵仍未做**，不得据此判定整体验收通过 |

## 8. 回退与最终放行

每阶段保持可独立检查的变更边界。需要回退时只撤回本次对应修改，保留前序和用户工作；真实测试用独立文件，不依赖未覆盖的统一撤销。接口兼容门面在调用方与发布包验证完成前保留。

最终放行清单：

- [ ] P0–P5 必验任务全部完成，未完成项没有以跳过冒充通过。
- [ ] 业务契约差异已逐项解释并接受，无意外工具删除或字段变化。
- [ ] 双系统及约定架构的发布包构建和运行证据完整。
- [ ] 全业务/故障矩阵通过，非目标内容和保存后数据完整。
- [ ] 无未解决的数据破坏、错误目标、重复写入、崩溃或阻断业务问题。
- [ ] 根导航、模块规范、skill、能力说明与实际版本一致。
- [ ] 最终报告说明修改范围、验证环境、产物位置及真实限制。

计划编写完成不代表改造完成。环境不足时可以推进不依赖该环境的阶段，但不能提前宣告整体验收通过。
