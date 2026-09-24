# 项目协作导航

## 从任务找入口

先读本文件，再读目标目录 AGENTS.md，随后只搜索该目录和下表关联文件。局部证据不足、入口不存在或问题跨模块时再扩大搜索，并说明新的调查范围。不要以全项目扫描、读取全部规范作为默认起点。局部规范补充本文件，不重复共性规则。

| 任务 | 首读目录规范 | 进一步定位 |
|---|---|---|
| MCP 参数、工具清单、路由、审计 | [Bridge](src/bridge/AGENTS.md) | tools/definitions/（工具定义按类）→ catalog.ts（装配与校验）→ gateway/（处理器）→ office/adapter.ts |
| WPS 执行、Word/PPT、矢量绘图 | [WPS](wps-addon/AGENTS.md) | src/dispatch.js（RPC 分支）→ src/{excel,word,ppt}.js → 部署入口 addon-core.js（**生成物**） |
| Microsoft 表格请求及宿主差异 | [Bridge](src/bridge/AGENTS.md)、[Office.js](office-addon/AGENTS.md) | office/adapter.ts → normalizer.ts → src/rpc.js → 部署入口 public/taskpane.js（**生成物**）；Windows 回退见原生驱动 |
| Windows 原生 Office 操作 | [原生 Office](resources/office/AGENTS.md) | runner.ps1 → excel.ps1/common.ps1 |
| 安装、客户端配置、桌面生命周期 | [主进程](src/main/AGENTS.md) | installer-engine.ts、addon-installer.ts、index.ts |
| 桌面交互与状态显示 | [界面](src/renderer/AGENTS.md)、[IPC](src/preload/AGENTS.md) | App.tsx → 相关组件 → 对应 IPC |
| AI 不会调用已有能力 | [技能](skills/AGENTS.md) | 对应 SKILL.md → gateway 工具说明 → 宿主实现 |
| 测试失败与验收 | [测试](tests/AGENTS.md) | 对应测试文件，不先运行真实宿主脚本 |
| 开发、检查、构建脚本 | [脚本](scripts/AGENTS.md) | 对应脚本和 package.json |
| 官网内容和交互 | [官网](website/AGENTS.md) | website/src/App.tsx → components |
| 宣传片剧本、录制环境与道具 | [宣传片](promo/AGENTS.md) | README.md（拍摄台本）→ 01/02 剧本 → [施工/plan.json](promo/施工/plan.json) |

其他目录：[docs](docs/) 放架构和验证说明；[prompts](prompts/) 放运行时提示词；[build](build/) 和 resources 的图片为打包素材。它们沿用根规范。发布包的代码保护（sourcemap 排除、V8 字节码、验证脚本）见 [code-protection.md](docs/code-protection.md)；发布专用副本由 [build-release-assets.mjs](scripts/build-release-assets.mjs) 生成到 `dist/release-app/`，只在该副本做混淆，业务源码和技能原稿保留。当前架构参考 [architecture.md](docs/architecture.md)，主入口配置为 [package.json](package.json) 和 [tsup.config.ts](tsup.config.ts)。

阶段验收证据按候选标识存放在 `docs/acceptance/<候选标识>/`，至少包含基线指纹、工具契约快照、环境清单与兼容性差异；只写"已完成"而没有对应证据的条目视为未完成。

证据材料要能被**下一个会话**直接使用：恢复类材料把锚点（提交 SHA、版本号）写进元数据文件，不要绑定会变化的 `HEAD`；脚本路径从仓库根目录解析，不写个人目录；修复前复现用临时副本，不要覆盖当前工作区。

**加载项部署入口为生成物**（改造计划 P3.4/P3.5）：`wps-addon/addon-core.js` 由 `npm run build:wps-addon` 从 `wps-addon/src/**` 生成，`office-addon/public/taskpane.js` 由 `npm run build:office-addon` 从 `office-addon/src/**` 生成；两者顶部都有"请勿手改"注释。**改源码后重新构建，禁止手改生成物掩盖源码问题**；从干净目录（删除生成物后）执行 `npm run build:addons` 必须能重建并通过测试。

生成物 dist/、release/、website/dist/、上述两个加载项入口和依赖 node_modules/ 不作为修改入口，也不维护 AGENTS.md。开发编译在 dist/，发布包按 release/<版本>/<平台>/ 存放；改源文件后通过构建更新，禁止手修产物掩盖源码问题。

## 改造任务入口

涉及模块化、契约统一及双平台验收的改造，先读 [改造计划与执行台账](docs/refactoring-plan.md)，按当前阶段与任务 ID 继续；计划内保留完成清单、验收标准和交接备忘。

## 协作与变更边界

使用简体中文。明确的小需求直接实施；架构、产品行为或明显扩大的范围先讨论。既定方向内的工程细节自主完成。先看 git status/diff，保留用户及其他任务的修改；不自动提交、推送、部署或清理数据。

界面、CLI、MCP 与宿主加载项职责分离。专用工具缺失不代表宿主不能做，先检查原生脚本路径。操作真实文件须在任务授权范围内，不能把测试脚本当成无副作用检查。

| 变更 | 同步检查 |
|---|---|
| 工具或参数 | schema、gateway 映射、宿主实现、读回结果、相关测试、skill |
| 宿主能力 | 实现、平台验证状态、MCP 描述、skill；涉及宣传时核对官网 |
| IPC | 主进程处理器、preload 暴露、renderer 调用 |
| 安装与服务 | 配置合并、凭据处理、生命周期、相关测试 |
| PPT 布局 | 真实页面尺寸、坐标字号、文本容量、读回和真实预览 |

## 验证与交付

按修改范围选择检查：[package.json](package.json) 中的 `npm run typecheck`、`npm test`、`npm run build`；官网使用 `npm run build:website`；导航使用 `npm run check:agents`。局部测试见模块规范，不为纯文档改动启动办公软件。

模拟测试、构建、真实宿主验证分别报告，不能互相替代。失败时区分本次回归、既有缺陷与环境限制，保留原始证据；不为通过检查顺手改无关模块。最终说明完成内容、主要文件、验证及遗留问题。

## 文档随代码维护

新增/移动/删除关键文件，改变职责、调用链、接口、命令或验证路径时，在同一次修改中更新根导航或最近的 AGENTS.md。普通函数内部变化无需制造文档更新。收尾检查：导航是否仍能定位、联动是否完整、避坑条目是否仍成立，再运行导航检查。

避坑条目格式为“触发条件 → 错误做法 → 已证实原因 → 正确做法 → 验证/证据”。高频关键条目就近写在模块规范，长案例才放 docs 并链接。只记录已证实且可复用的经验；猜测留在调查结果里。能回归的问题优先补测试，文档解释原因。修复使旧限制失效时更新或删除条目，不保留永久禁令。

不要记录凭据、用户文档内容、临时版本状态、个人路径或易失效行号。检查脚本只能确认导航结构与路径，职责和经验是否正确仍需修改者核对。跨模块并行工作只有获准后开展，先分配文件归属与接口边界。
