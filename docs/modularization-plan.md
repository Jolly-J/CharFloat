# 超大单体文件模块化拆分方案

状态：M0、M1 已完成静态、构建与模拟验证，M1-H 本机 macOS WPS 回归验收完成，Windows 待验；Word、PPT 和 Bridge 领域拆分后续批次尚未执行。
范围：WPS Excel 为第一批；Word、PPT 分批推进；Bridge 定义与处理器为可选后续批次。

本方案延续 [改造计划与执行台账](refactoring-plan.md) 的渐进模块化方向，不替代历史记录或真实宿主验收台账。第一批证据见 [Excel 模块化验收记录](acceptance/2.2.0-excel-modularization/README.md)。

## 1. 现状与可行性

以下为迁移前的源码规模；Excel 已迁至 18 个领域片段，最大片段 703 行，其余文件保持原状：

| 文件 | 行数 | 主要职责 |
|---|---:|---|
| `wps-addon/src/excel.js` | 5,930 | 表格读写、样式、验证、图表、形状、工作簿与事件 |
| `wps-addon/src/word.js` | 2,819 | 文档读写、版式、表格、审阅、域与内容控件 |
| `wps-addon/src/ppt.js` | 1,673 | 整套生成、页面布局、形状、表格、图表、保存与预览 |
| `src/bridge/tools/definitions/excel.ts` | 1,405 | 表格工具 schema、说明与上下文描述 |
| `src/bridge/gateway/excel.ts` | 1,069 | 表格处理器、参数转发与审计 |

### 可沿用的基础

- Office 加载项已有 `office-addon/src/excel/` 领域分片，共 13 个源码文件，可参考职责组织及递归枚举方式，不照搬 Office.js 宿主实现。
- WPS 构建器将源码片段顺序拼入一个 IIFE，再由现有 esbuild 压缩，目标为 `es2018`；保留这条完整链路。
- `ppt-layout.js` 和 `sheet-sort.js` 已使用 `@build-strip` 包裹测试导出，可沿用其 Node 直接导入方式。
- 现有 RPC、生命周期、布局、契约、指纹与发布保护测试可作为回归基础。

### 必须补足的条件

1. [WPS 构建器](../scripts/build-wps-addon.mjs) 原先只枚举 `src` 第一层。M0 已补齐递归清单校验及未登记、缺失、重复反向测试，再完成 Excel 迁移。
2. 原拟议格式模块的相关函数体约 1,100 行，`table-pivot` 的主要函数体约 1,211 行，尚未包括全部辅助逻辑；原行数预算不能作为切分依据。
3. 文件拆开后仍共享 IIFE 作用域，改善定位和编辑冲突，但不会自动获得模块作用域隔离或 Node 导入能力。
4. [参数转发检查器](../scripts/check-param-forwarding.ts) 同样只扫描定义和处理器目录的第一层。Bridge 拆分前必须适配，防止扫描覆盖减少后仍报通过。

## 2. 目标与变更边界

- 仅治理本方案列出的目标文件，建立清晰的领域入口、函数归属和依赖说明，不扩展为全项目千行文件治理。
- 大部分业务文件以 300～800 行为参考。完整职责可以小于 300 行；超出参考区间时说明原因，不为了行数切断原函数或制造转发文件。
- 抽出的纯函数可直接在 Node 中导入，不依赖 `window`、`wps` 或宿主对象，不在导入时创建连接、事件和计时器。
- 工具名、参数、说明、顺序、返回、错误、审计、目标定位及宿主写入语义保持一致。
- 不增加运行时模块加载或网络请求；对代表性路径记录同环境耗时，明显变化需调查，不承诺无法验证的“零性能损耗”。

不新增办公能力、不重写算法、不调整产品行为、不升级依赖、不引入新打包器、不改变发布保护策略。不自动安装、部署、重载真实宿主、提交或发布。发现既有业务缺陷时另行登记，结构搬迁与业务修复分开处理。

### 源码与运行时约束

- 保留 `wps-addon/addon-core.js` 部署路径、外层闭包、现有压缩配置及指纹协议；不手改生成物。
- WPS 业务函数保持原名称与普通 `function` 声明，不改成箭头函数，不增加模块私有 IIFE。
- 根 `wps-addon/src/shared.js` 已有 `getWorksheet`、`hexToExcelColor` 等跨组件 helper，保留原位置，不在 `excel/` 内重复定义。
- 领域专用常量和状态留在所属文件，不强制集中到 `shared.js`；核对初始化依赖，确保第一次使用前完成初始化。
- 根 `shared.js` 保持最前，`bootstrap.js` 保持最后；保留 `sheet-sort.js`、`ppt-layout.js` 的职责和测试导出。
- 拆分会改变源码拼接体与构建指纹。验收要求新生成物头部、注册和心跳上报一致，不要求新旧指纹相等。

## 3. 分阶段实施路线

### M0：冻结基线并补足构建能力

在移动业务代码之前完成：

1. 记录固定提交 SHA、版本、工作区差异及目标源码、锁文件、生成物的 SHA-256。保存目标文件原始内容与可恢复差异，保留已有修改并标明归属。
2. 在新的 `docs/acceptance/<候选标识>/modularization/` 目录保存基线、契约、函数归属、初始化顺序及检查结果，不覆盖旧候选材料。
3. 构建器递归枚举 `src/**/*.js`，以一致的相对路径与显式 `MODULES` 双向比较；保留重复、缺失、未登记检测，不用目录自动排序代替拼接顺序。
4. 在临时夹具中证明三类漂移均失败，尤其是“子目录新增文件未登记”；合法嵌套清单成功，检查模式不写文件。
5. 业务代码未移动时，构建器改造不得改变生成物内容。

建立原文件全部顶层函数、常量、状态和执行语句的唯一归属清单。迁移后核对无遗漏、无重复定义；清单匹配只证明文件集合完整，不证明依赖顺序正确。

### M1：WPS Excel 拆分

目标文件均位于 `wps-addon/src/excel/`。以下为职责与主要入口，执行前补齐所有辅助函数，不使用未经统计的逐文件行数预算：

| 文件 | 职责与主要入口 |
|---|---|
| `pure.js` | 无宿主依赖的矩阵与列名计算，如 `normalize2DArray`、`excelColumnName`；测试导出 |
| `design.js` | 设计语言读取、样式普查、层级、调色板和令牌、`getStyleToken`；相关 `STYLE_*` 常量 |
| `styles.js` | 样式、对齐与边框读取、`getRangeStyles`；调用验证读取模块 |
| `format.js` | `formatCells`、`formatTextSegment`、`addConditionalFormatting` |
| `validation-read.js` | 验证读取、列表与边界解析、值评估、`collectValidationViolations`；相关映射与扫描限制 |
| `validation-write.js` | `buildDataValidationPlan`、规则恢复、写后校验、`setDataValidation` |
| `range.js` | 读写、搜索、清空、复制、查找替换、回滚及变更检测；`CHANGE_BASELINES` 与 `captureOrDiffChanges` 一起迁移 |
| `sheets.js` | 工作表增删复制、`manageSheet`、大纲、工作表状态、`setSheetView` |
| `rows-columns.js` | 行列调整、插入、自动适配、冻结窗格 |
| `chart-read.js` | 图表清点、序号定位、导图：`getCharts`、`getChartOrdinal`、`exportChartImage` |
| `chart-write.js` | `addChart`、`updateChart`、`deleteChart` |
| `shape.js` | 形状读取、增改、编组、层级、图片管理及相关类型/字体常量 |
| `table-pivot.js` | 仅保留结构化表与透视表，不再承接验证、批注、筛选排序 |
| `filter-sort.js` | `setFilterAndSort`，保留根 `sheet-sort.js` 的排序读回校验 |
| `comments-links.js` | 批注、超链接与命名区域 |
| `workbook.js` | 新建、保存、工作簿自定义视图与文档属性 |
| `export.js` | 打印布局、PDF 导出、工作表预览 |
| `events.js` | `hookWpsEvents`、`onSelectionChange`；共享连接与事件状态的原位置和注册时机不变 |

最终名称可按函数归属微调，不漏掉未公开函数、内部校验或事件路径。小模块保留完整职责，不要求填满行数。

#### 构建清单目标

以下清单在递归校验完成、初始化依赖核对后生效：

```javascript
const MODULES = [
  "shared.js", "connection.js", "ribbon.js", "dispatch.js", "sheet-sort.js",
  "excel/pure.js", "excel/design.js", "excel/validation-read.js",
  "excel/styles.js", "excel/format.js", "excel/validation-write.js",
  "excel/shape.js", "excel/range.js", "excel/sheets.js", "excel/rows-columns.js",
  "excel/chart-read.js", "excel/chart-write.js",
  "excel/table-pivot.js", "excel/filter-sort.js", "excel/comments-links.js",
  "excel/workbook.js", "excel/export.js", "excel/events.js",
  "word.js", "ppt-layout.js", "ppt.js", "panel.js", "bootstrap.js",
];
```

实际迁移将 `shape.js` 放在 `range.js` 前，保留原文件全部 15 个初始化声明的相对顺序（STYLE → VALIDATION → SHAPE → CHANGE_BASELINES）；函数声明仍提升，调用发生在启动前初始化完成之后。

纯函数采用现有的受保护测试导出，拼接前整块移除：

```javascript
  // @build-strip:start — 仅用于 Node 单元测试
  export { normalize2DArray, excelColumnName };
  // @build-strip:end
```

约束为**宿主生成物不得残留模块语法**。源码只允许上述测试导出，不新增运行时 `import`、`export` 或 `require`。纯函数所需常量必须同属可导入模块，不能依赖闭包里未导出的宿主状态。

### M2：Word 与 PPT 各自分批拆分

Excel 静态与模拟检查通过后再进入下一批。真实宿主验收未完成时保持“待宿主验收”，不得写成全部完成。Word、PPT 各自保存基线与证据，不混成一批迁移。

| Word 目标文件（`wps-addon/src/word/`） | 职责 |
|---|---|
| `document.js` | 新建、保存、关闭与属性读取 |
| `read.js` | 文档摘要、段落、书签、样式等读取，调用域与内容控件 helper |
| `content.js` | 内容管理、段落定位与插入、写入、格式排版、查找替换 |
| `layout.js` | 页面设置、页码与水印，保留危险参数拒绝逻辑 |
| `table.js` | 表格管理 |
| `review.js` | 修订、批注、审阅 |
| `fields.js` | 域类型、快照/差异、目录、交叉引用、故事区域更新与锚点安全判断 |
| `content-controls.js` | 控件类型、定位、列表、创建/更新与宿主限制 |
| `preview.js` | 文档预览 |

`WORD_CC_TYPES`、`WORD_CC_TYPE_NAMES` 与填充映射的顶层 `for` 循环一起迁移，先于首次调用完成初始化；`WORD_FIELD_TYPE_NAMES` 留在域模块。不得只搬常量而漏掉执行语句。纯计算 helper 确认无宿主依赖后才抽为可导入叶子文件。

| PPT 目标文件（`wps-addon/src/ppt/`） | 职责 |
|---|---|
| `shared.js` | 颜色转换、页面索引、形状定位等共同 helper |
| `presentation.js` | 演示文稿读取、保存、预览 |
| `slides.js` | 页面增删复制重排、尺寸配置与 `SLIDE_PRESETS` |
| `render.js` | 页眉、卡片、矢量图表、整套生成与 `fitGeneratedPptShapes` |
| `shapes.js` | 形状清点、编辑、媒体管理 |
| `table-chart.js` | 表格管理与原生图表插入 |
| `business-card.js` | `pptAddBusinessCards` 与场景排版 |

根 `ppt-layout.js` 继续承载布局纯函数。同步修改 [PPT 布局测试](../tests/ppt-layout.test.ts) 中读取 `src/ppt.js` 的路径，加载适配器所在的新片段并注入必要依赖，不保留仅供旧测试读取的重复实现。

### M3：Bridge 定义与处理器拆分（可选后续）

保留 `definitions/excel.ts`、`gateway/excel.ts` 兼容门面，从子目录显式聚合/重导出，尽量保持调用方导入路径及 `excel.<handler>` 注册形式不变；门面不再承载业务正文。

- 领域子文件按实际职责组织，如 cells、formatting、validation、charts、shapes、sheets、workbook。
- 保留 `excelToolDefinitions(ctx)` 和 `excelToolDefinitionsAfterAudit(ctx)` 两段输出，[总装配](../src/bridge/tools/definitions/index.ts) 中审计工具的插入位置不变。不得按领域重排已发布的工具顺序。
- 原样透传 `ctx.wbDesc`、`ctx.activeWbHint`，分别比对无上下文和带上下文的定义输出。
- 处理器继续经 `GatewayContext` 注入执行器，保留目标锁、会话、审计、错误分类和响应透传。
- 移动前适配参数检查器：递归发现文件、按相对路径识别模块、正确解析门面重导出，防止同名处理器互相覆盖。
- 反向夹具覆盖嵌套模块丢参、处理器无法解析、同名处理器归属错误；比对迁移前后扫描与解析集合，不接受检查覆盖减少后仍报通过。既有未解析例外单独说明。
- 使用显式文件导入，核对 [依赖边界测试](../tests/contracts-boundary.test.ts) 对新路径的解析仍有效，无环约束不放宽。

## 4. 等价性证据与风险控制

### 代码与初始化

- 原文件的顶层声明、执行语句和跨文件依赖全部有唯一归属，迁移后无遗漏或新增重名。
- 用去除位置/注释信息的 AST 或函数级内容对照证明原函数体不变；文件头、测试导出及必要构建/装配差异逐项说明。
- 单独核对非函数语句的相对顺序、初始化依赖和副作用时机。仅比较声明数量或语句多重集不足以证明等价。

### RPC、契约与压缩

- 迁移前后运行同一组模拟宿主输入，比对返回、错误、宿主调用顺序/次数、写入值、审计与目标信息。仅排除请求 ID、时间戳、构建指纹等明确易变字段，不忽略业务差异。
- 覆盖读取、普通写入、混合值/公式、执行前拒绝、宿主失败和未知方法；每个迁移领域有代表用例，现有单元格用例不能替代全领域回归。
- 契约快照逐字节一致，以本次实际工作区为基线，不套用历史工具数或快照。带 `activeWbHint` 的动态描述另行对比。
- 测试实际压缩生成物的 RPC 与生命周期；核验头部前 4096 字节可读指纹、注册和心跳一致性；发布保护副本另跑现有回归。
- 指纹、语法和模拟测试不能替代 WPS JSA 的真实运行证据。

### 回退与导航

每批独立复核和回退。只撤回本批迁移与构建清单调整，保留前序和用户修改；从固定基线恢复后重新构建，不手修生成物，不对当前工作区执行 `git reset/clean`。

入口或测试路径变化时，同批更新最近的导航：[WPS](../wps-addon/AGENTS.md)、[脚本](../scripts/AGENTS.md)、[测试](../tests/AGENTS.md)，Bridge 阶段再更新 [Bridge](../src/bridge/AGENTS.md)。根入口改变时同步根导航，不提前将未创建目录作为有效链接。

## 5. 验证步骤与完成判定

### 静态、构建与模拟检查

以下是代码实施的检查命令；M0/M1 的实际结果另记录在本批证据中：

```bash
npm run build:wps-addon
node scripts/build-wps-addon.mjs --check
node --check wps-addon/addon-core.js
node --import tsx --test tests/addon.test.ts tests/addon-lifecycle.test.ts tests/ppt-layout.test.ts tests/sheet-sort.test.ts tests/addon-freshness.test.ts tests/release-protection.test.ts tests/wps-addon-modules.test.ts tests/wps-excel-pure.test.ts tests/wps-excel-rpc.test.ts
npm run typecheck
npm run build
npm test
npm run check:params
npm run check:agents
```

- 另跑本轮新增的递归清单、纯函数和领域 RPC 用例。测试保留在根 `tests/`，遵循 `*.test.ts` 命名，避免根命令漏发现。
- `--check` 仅比对源码重新生成的内容与磁盘生成物，不输出逐文件哈希匹配，也不证明新旧业务行为等价。
- 同输入重复构建必须逐字节一致。干净重建在临时副本中移除生成物后执行 `build:addons`，不删除当前工作区产物。
- 完整 `npm test` 前必须成功执行 `npm run build`，因为 stdio 测试执行 `dist/bridge/cli.cjs`，不能测试旧产物。
- 参数检查如仅因 tsx CLI 本地管道权限失败，可用 `node --import tsx scripts/check-param-forwarding.ts` 执行同一脚本；保留原错误，不能以此忽略业务检查失败。
- 不固定“56+ 项”历史数量，记录实际发现、通过、失败、跳过统计。必验项不得跳过，既有失败与环境限制分别登记。
- 拆分前后各运行 `snapshot:tools`，保存到本次证据目录并比对。M1 确认导出含 `capabilities.build` 和 `capabilities.connection.build` 的当前时间、磁盘更新时间及重建后 SHA；明确排除这两块构建诊断，其他契约/能力逐字节比对，产物指纹单独留证，不修改快照导出器。Bridge 阶段另比对带上下文的定义输出、注册映射及参数检查覆盖集合。

### 真实宿主验收

使用明确授权的独立测试文件，记录实际加载指纹、宿主版本与组件。Excel 至少覆盖读写、样式/验证、工作表与工作簿、图表、形状、结构化表/透视表/排序、批注链接、导出和事件；记录既有不支持能力及预期错误。

Word、PPT 按本批职责覆盖创建/读取/编辑/保存、主要对象、错误路径及预览。写后读回，保存后关闭重开核验落盘；几何检查使用真实页面尺寸、坐标字号和真实预览。保留危险调用禁用，不触发已知崩溃路径。

执行前确定本批必验的系统和宿主版本。缺少环境则登记待验，不用 macOS 模拟或交叉构建替代 Windows 宿主验收；Bridge 涉及跨宿主链路时追加 Microsoft Excel 与原生回退相关场景。

### 分层完成标准

| 状态 | 必要证据 |
|---|---|
| 代码拆分完成，待宿主验收 | 清单完整、函数/初始化无未解释差异、契约不变、构建幂等、干净重建及相关静态/模拟/完整回归通过、导航同步 |
| 本批宿主验收完成 | 上述证据齐全，本批必验系统/宿主的加载、业务、保存重开及必要预览有真实记录 |
| 发布验收完成 | 另按现有发布流程完成包内容、保护和真实运行检查；代码拆分不自动授予此状态 |

证据包含固定源码锚点、工作区差异、输入及产物指纹、归属与依赖、契约、命令和退出码、模拟前后差异、实际测试统计，以及宿主读回/预览或待验原因；不得包含凭据或用户真实文档内容。

## 6. 待执行清单

- [x] M0：固定提交与文件指纹基线、函数/初始化归属清单、递归构建校验及 4 项反向/合法夹具完成；移动前生成物一致。
- [x] M1：88 个函数、15 个初始化声明迁至 18 个片段，内容及初始化顺序不变；3 项纯函数与 21 项领域 RPC 测试、契约比对、完整构建、干净重建与重复构建通过；完整回归 152/152、无跳过。证据见本批验收记录。
- [ ] M1-H：**本机项完成，跨平台项待验**。macOS WPS 12.1.28496 已经独立 CLI 验证部署/重载、运行指纹、业务、保存关闭重开、图片和 PDF；Windows 未验。既有问题与全部证据见 [本机宿主报告](acceptance/2.2.0-excel-modularization/host/README.md)。
- [ ] M2-W：Word 独立迁移及静态/模拟回归完成。
- [ ] M2-W-H：Word 本批真实宿主验收完成。
- [ ] M2-P：PPT 独立迁移、测试路径同步及静态/模拟回归完成。
- [ ] M2-P-H：PPT 本批真实宿主验收完成。
- [ ] M3：如进入 Bridge 批次，先适配参数检查再拆定义/处理器，契约顺序和检查覆盖不退化。
- [ ] M3-H：Bridge 本批相关宿主链路验收完成。

每批记录实际变更、验证证据、未完成事项和下一步。第一批代码执行范围为 M0、M1，真实宿主操作另按可用环境与测试文件授权执行。后续批次不因文档列出目标而视为已经实施或验收。

第一批交接：后续已在本机部署/重载 WPS 候选，真实操作仅限新建测试文档，未提交或发布。M1-H 本机项完成，Windows 待验；Word/PPT/Bridge 领域拆分后续批次独立推进。真实 CLI 验收发现字节码运行时的 fs 动态导入失败，用户批准后改为两处静态依赖并完成实际字节码回归和真机导出复验；修复后完整回归 153/153、契约不变。首轮沙箱监听错误及宿主探针修正保留原始记录，不混作产品回归。
