# MCP 与宿主路由

## 文件地图

- [gateway.ts](gateway.ts)：**兼容门面**。`executeTool` 只做锁定注入 → 查 `HANDLERS` 注册表 → 组上下文 → 调用；`registeredGatewayBranches()` 导出真实注册表键供测试读取；工具定义已按类迁出（P2.3），此处 `getOpenAiTools(activeWbHint?)` 只委托 `tools/definitions/`；`getOpenApiSchema` 仍在此。
- [gateway/](gateway/)：**按类拆分的执行处理器**（P3.1）。`types.ts` 上下文与处理器类型、`locks.ts` 目标锁、`excel.ts`/`word.ts`/`ppt.ts` 三类文档操作、`script.ts` 原生脚本与反射、`audit.ts` 回滚与审计、`lock.ts`/`microsoft.ts`。处理器不各自 import 执行器，全部经 `GatewayContext` 注入。
- [contracts/](contracts/)：**契约层**，纯定义，不引用执行器与服务实例。
  - [contracts/host-methods.ts](contracts/host-methods.ts)：宿主方法标识、路由表、宿主实现缺口、只读与重放属性。
  - [contracts/tool-service.ts](contracts/tool-service.ts)：协议层需要的 `ToolService` 接口与工具描述结构。
- [tools/](tools/)：**工具定义与分类注册**（P2.3 逐类迁移）。
  - [tools/index.ts](tools/index.ts)：装配点，顺序为 诊断 → 统一 excel 入口 → 网关原生工具 → 审计；含 `toolClassOf` 分类。
  - [tools/definitions/](tools/definitions/index.ts)：**按类拆分的工具定义**（P2.3）：`shared.ts` 定义类型与 `activeWbHint` 上下文、`excel.ts`(864) / `word.ts`(337) / `ppt.ts`(299) / `microsoft.ts` / `lock.ts` / `script.ts` / `audit.ts`，`index.ts` 按**已发布顺序**拼装。
  - [tools/diagnostics.ts](tools/diagnostics.ts)：`bridge_*` 诊断工具。
  - [tools/excel.ts](tools/excel.ts)：由宿主方法路由表派生统一 `excel_*` 入口（原 `wps_*` 兼容名并保留）。
  - [tools/audit.ts](tools/audit.ts)：审计查询与清理工具。
- [tool-registry.ts](tool-registry.ts)：**运行时注册层**，把 catalog 的定义与执行器绑定成 `ToolService`。
- [compose.ts](compose.ts)：**组装入口**，唯一知道"谁实现谁"，向服务实例注入工具服务。
- [catalog.ts](catalog.ts)：共享工具清单、入参校验与能力描述。
- [mcp-server.ts](mcp-server.ts)：MCP 协议、说明及响应；经注入的 `ToolService` 取工具，不引用 catalog。
- [ws-server.ts](ws-server.ts)：HTTP/WebSocket、认证及加载项连接；经注入的 `ToolService` 取工具，不引用 catalog。
- [errors.ts](errors.ts)：失败分类（不可用/拒绝/失败/未知）与跨通道回退策略；只依赖契约层。
- [context.ts](context.ts)：请求会话与宿主上下文。
- [office/adapter.ts](office/adapter.ts)：宿主队列和 Office.js/COM 路由。
- [office/normalizer.ts](office/normalizer.ts)：Office.js 请求响应转换。
- [office/ms-office-driver.ts](office/ms-office-driver.ts)：Windows 原生驱动及 macOS JXA。
- [service-client.ts](service-client.ts)：发现或启动独立后台。
- [cli.ts](cli.ts)：CLI 和 stdio 入口。
- [audit-store.ts](audit-store.ts)：审计持久化；含"未审计的身份类操作"台账（回滚身份判据用）。
- [build-fingerprint.ts](build-fingerprint.ts)：磁盘/部署产物的构建指纹（桥接入口、WPS 加载项随包与已部署副本、Office.js 任务窗格），供判断"运行中的是哪一版"。
- [process-runner.ts](process-runner.ts)：原生进程输入输出。
- [runtime.ts](runtime.ts)：路径与运行配置。

## 定位与联动

参数问题从 gateway 对应工具分支追到 catalog 和 adapter。连接问题从 service-client/ws-server 开始。Microsoft 优先 Office.js，Windows 存在 COM 回退；office_execute_script 另走原生驱动，不能混为一条链。失败种类与回退判定集中在 `errors.ts`，不要在 adapter 或 gateway 里另写一套“能不能重试”的条件。分层固定为：**契约层** `contracts/`（类型/schema/能力元数据，不引用执行器）→ **执行层**（gateway/adapter）→ **协议层** `ws-server`/`mcp-server`（只认注入的 `ToolService`）← **组装入口** `compose.ts` 注入。协议层与执行层不互相静态引用。

## 修改边界

保留会话与宿主隔离；办公调用按宿主串行。认证数据不进入日志。新增字段同时检查 schema、映射、normalizer 和宿主读写，不能只修改描述。错误不能折叠为成功；回滚必须校验修改后快照。

## 验证

`node --import tsx --test tests/service.test.ts tests/office.test.ts tests/platform.test.ts tests/lifecycle.test.ts`（项目根运行）；改动失败分类或回退策略时追加 `tests/failure-routing.test.ts`；工具的**定义**放 `tools/` 对应类文件，`catalog.ts` 只做装配调用与兼容导出——**装配顺序影响 tools/list 与契约快照**，改动后用 `npm run snapshot:tools` 比对必须逐字节不变；改动模块边界或契约层时追加 `tests/contracts-boundary.test.ts`；**新增入口点（如在别处调用 `bridgeServer.start()`）必须同时调用 `composeBridgeServer()` 装配工具服务**，否则 HTTP/MCP 路由会报"工具服务尚未装配"；改动 stdio 代理或 `ToolService` 签名后必须 `npm run build` 再跑 `tests/session-isolation.test.ts`（它拉起 `dist/` 产物）；再按变更执行 typecheck/build。

## 避坑

工具声称绑定目标但 schema 缺少目标名称 → AI 无法显式传参 → 检查 schema 与执行分支是否都接受名称 → 用 [参数契约测试](../../tests/ppt-layout.test.ts) 核验。写入超时或 Office.js 异常 → 不直接假设未执行或安全重试 → 先读回；尤其修改自动回退策略时检查重复写入风险。

Office.js 写入在“宿主已执行之后”才失败（超时、断连、宿主报错、响应转换失败）→ 旧实现对任何异常都回退 Windows 原生 COM → 同一次写入被执行两次，且第二次的成功掩盖第一次的失败，调用方拿到“成功” → 只有 `executed === 'no'`（通道未连接、参数在执行前被拒）或不改文档的方法才允许回退，其余报错并要求先读回 → [故障注入测试](../../tests/failure-routing.test.ts) 在 HEAD 版 adapter 上 2 项失败、修复后 10/10 通过（日志见 [evidence](../../docs/acceptance/2.1.0-p0p1/evidence/p1.4-prefix-failure.txt)）。

用 `npm run build` 退出码判断类型是否正确 → tsup 不做类型检查，基线曾出现 build 通过而 `typecheck` 4 处 TS2304 → 类型与构建是两层检查，必须分别执行 → 基线记录见 [baseline.md](../../docs/acceptance/2.1.0-p0p1/baseline.md) §3。

能力声明直接回显路由表（如 `implemented: [...EXCEL_METHODS]`、写死的“全量 N 项”）→ 声明与可调用工具数、宿主实现不符 → 路由表含没有 schema 的宿主方法，宿主加载项也可能缺 RPC 分支 → 声明按“可调用工具 ∩ 宿主实现”生成，缺口单列 `declaredNotCallable` / `unimplementedOnHost` → 差异清单见 [compatibility-diff.md](../../docs/acceptance/2.1.0-p0p1/compatibility-diff.md) §3。

把“反射探测”当成只读工具 → `wps_inspect_api` 被同时标为只读和可安全重放 → 它执行调用方给出的**任意表达式**，表达式可取属性触发宿主求值、也可直接改文档（如 `app.ActiveWorkbook.Worksheets.Add()`），静态无法保证无副作用 → 既不得标 `readOnlyHint`，也不得进入跨通道重放白名单；只有确认未执行时才可回退 → [回归测试](../../tests/failure-routing.test.ts) 的“带副作用的表达式探测既不算只读，也不允许跨通道重放”。

执行层需要契约数据时顺手 `import` 注册入口（catalog）→ 形成 `catalog → gateway → adapter → catalog` 回环 → 契约与执行层互相依赖，依赖方向无法单独理解，改动任一侧都可能影响初始化顺序 → 宿主方法标识等契约数据放 `contracts/`（`catalog` 保留同名兼容导出）；协议层改为通过**注入的** `ToolService` 取工具，由 `compose.ts` 装配 → [边界守卫测试](../../tests/contracts-boundary.test.ts)（证据见 [P2.1](../../docs/acceptance/2.1.0-p0p1/evidence/p2.1-prefix-adapter-catalog-cycle.txt)、[P2.2](../../docs/acceptance/2.1.0-p0p1/evidence/p2.2-prefix-scc.txt)）。

用"已访问即跳过"的 DFS 判断依赖回环 → 只能发现回环、不能完整枚举 → 早到一个路径会把中间节点标记完成，通往同一节点的其它路径被跳过（例如先经 adapter 到达 ws-server，就看不到 gateway 直达 ws-server 这条边），漏报后容易误判"只剩两条回环" → 用**强连通分量（Tarjan）**做完整口径，台账为空才算真正无环 → [边界守卫测试](../../tests/contracts-boundary.test.ts) 的 SCC 断言。

跨进程转发工具调用时把 `sessionId` 交给 `AsyncLocalStorage` → 跨 HTTP/进程边界不会自动携带 → 后台落回默认会话 `http-local`，不同 stdio 客户端的文档锁互相覆盖（A 锁定 A.xlsx、B 锁定 B.xlsx 后，A 查询到 B.xlsx），A 后续操作会误指向 B 的文件；改 `ToolService` 签名时最易漏 → `sessionId` 与 `clientName` 放进 `ToolCallContext` **随每次 execute 显式传递**，远程实现原样转发到后台 → [session-isolation.test.ts](../../tests/session-isolation.test.ts)（复现证据见 [evidence](../../docs/acceptance/2.1.0-p0p1/evidence/p2.2-session-isolation-regression.txt)）。

契约层里放执行器或可变注册表 → 只是把回环换个位置，并引入依赖初始化顺序问题 → 契约层保持纯定义（类型/schema/能力元数据），实现绑定放 `tool-registry.ts`，装配放 `compose.ts` → 契约层依赖校验见 [边界守卫测试](../../tests/contracts-boundary.test.ts) 的"契约层不依赖执行层、协议层与服务实例"。

新增工具只加执行分支不补定义 → 出现"实现了但未注册"死分支 → AI 永远调不到该能力，且没人会被提醒 → 定义加到 `tools/` 对应类文件并让 `toolClassOf` 能归类，死分支台账由 `tests/contract-consistency.test.ts` 守住（只减不增） → 该测试的「实现了但未注册」断言。

按类拆分工具定义时凭印象重排顺序 → `tools/list` 与契约快照变化，客户端看到的工具顺序被打乱（诊断工具曾是 `bridge_get_capabilities` 在前，重排后变成 `bridge_diagnose` 在前） → 装配顺序是有契约意义的，必须与迁移前逐项一致 → `npm run snapshot:tools` 比对必须逐字节不变。

测试从源码里正则抓 `case "..."` 统计分支 → 分支改成注册表后，测试实际只在**校验注释文本**，注释与真实执行映射脱节也不会失败（被削弱的检查） → 分支清单曾以注释保留在门面里，抓文本就会退化成"看注释通过" → 测试改读真实注册表（`registeredGatewayBranches()` 导出 `HANDLERS` 键），并在 `executeTool` 未知工具时抛错 → [契约一致性测试](../../tests/contract-consistency.test.ts) 的「注册工具缺少网关分支」与「实现了但未注册」两项；注册表 68 条。**不要为迁就测试而在源码里保留伪 `case` 注释清单**——那会让测试退化成校验注释，该清单已移除。

回滚的“是否有后续修改”只比区域内容 → 别人删表重建并写回**完全相同**的内容时，旧 `auditId` 被放行、把这批新数据清空（真实数据破坏） → 内容相等不等于同一次修改 → 回滚前先判**身份**（同区域是否有更晚的未回滚记录、记录之后是否发生过未经快照记录的身份类操作、目标工作表是否仍存在），最后才比内容；任一不满足即拒绝并给出“读 `wps_get_audit_record` 的旧快照 + `patch_cells` 显式写回”的替代路径 → [修复记录](../../docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/12-bridge-fixes.md)；判据来源与残余缺口（绕过桥接的手工删表重建无信号）同页说明。

跨宿主参数映射在桥接侧与宿主侧**各写一份** → 宿主侧补齐后桥接侧仍在抛错，把宿主已经能做的能力挡掉（本任务实测：`manage_sheet` move、行列 hide/set_size、`add_chart` 的 dataRanges/yAxis 都曾被桥接侧误拒） → 两层各自演进必然漂移 → 宿主已按 schema 字段名实现并读回校验的，桥接侧只做字段归一/别名，**不重复实现、不重复阻断**；真正映射不了的才显式报错或记入响应 `unsupportedFields` → [修复记录](../../docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/12-bridge-fixes.md) 的 ISS-93/ISS-99 两行。

响应转换用白名单重建对象 → 宿主新增的 `warnings` / `dataRangesApplied` / `seriesCount` / `yAxisApplied` / `kind` / `renderedBy` 等“如实提示”被静默丢掉，调用方以为参数都生效了 → 白名单只对新实现友好，宿主演进即失效 → 响应分支先铺开宿主原始响应再覆盖统一字段（`{...raw, ...}`），并在宿主已自行过滤/截断时不要重复处理 → 同上 ISS-99。

把"执行前拒绝"一律当作可换通道重放 → 参数安全护栏被绕过：空搜索串的查找替换在桥接层被判为破坏性并抛错，但这次 `rejected`（`executed === 'no'`）仍被 `fallbackToNative` 转交 Windows 原生 COM，而 `excel.ps1` 对空查询用 `IndexOf('')`（恒真）＋空正则逐格替换，把整片内容改坏（护栏亲手把请求送到没有护栏的通道） → 回退判定只区分了"是否已执行"，没区分**拒绝的性质**（能力缺口 vs 参数安全） → 安全护栏用独立类型（`normalizer.ts` 的 `UnsafeParamError` / `errors.ts` 的 `BridgeErrorOptions.unsafe`）标记，`routeOfficeFailure` 见 `unsafe` 直接拒绝回退；能力缺口仍保留 COM 接管（如"Office.js 只开筛选按钮"） → [故障注入测试](../../tests/failure-routing.test.ts) 的「CAP-50：参数安全护栏（空搜索串）不得换通道重放」断言 win32 下 Office.js 与 COM **各 0 次调用**；决策表用例见同文件。

COM 回退白名单用手写缺口清单 + 路由表做减法 → 路由表从 30 涨到 50 后，减法结果 48 项里有 20 项在 `resources/office/excel.ps1` 里根本没有分支，白名单再次过度声明（回退后只抛 `Unsupported Excel method`） → 手写清单一过期就没人提醒，而"路由表 ≠ 宿主实际覆盖"是常态 → 白名单与缺口清单**都由实际实现集派生**（`contracts/host-methods.ts` 的 `COM_IMPLEMENTED_METHODS`，28/50） → [契约一致性测试](../../tests/contract-consistency.test.ts) 的「CAP-51：COM 回退白名单与 excel.ps1 实际分支同源」**解析脚本真实分支**再比对（含防解析器失效的空跑断言）。

残留目标锁（宿主重启后锁指向已不存在的文档）直接报"未找到目标文档" → 本会话后续所有带锁工具持续失败，报错不指向"该重新锁目标"，使用者以为工具坏了（ISS-126） → 锁存在桥接进程内存里，宿主重启只换了文档集合，没有失效信号 → 失败时按两条同时成立的判据自动清理并改写提示：①报错匹配"目标名字在宿主里不存在"（三通道措辞见 `gateway/lock.ts` 的 `TARGET_NOT_FOUND_PATTERNS`）②该目标是**由会话锁注入**而非调用方显式传入（显式传错名字时不动用户的锁）；同时保留原失败的 `kind/executed` → [故障注入测试](../../tests/failure-routing.test.ts) 的「ISS-126」三场景（自动清理 / 不误删 / 不误清）。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
