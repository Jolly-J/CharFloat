# M1-H 本机真实 WPS 验收

日期：2026-10-12。结论：**macOS WPS 的本次模块化回归验收完成，保留既有问题说明**；Windows 真实宿主未验，不据此作全平台放行。

## 环境与范围

- macOS 27.0.1 / arm64，WPS 12.1.28496（应用标识 `com.kingsoft.wpsoffice.mac`）。
- 独立 Electron Node 运行时加载 `dist/bridge/cli.cjs` / `cli.jsc`，通过真实 stdio MCP → 独立后台 → WPS RPC 执行，不依赖桌面窗口承担后台服务。
- 服务使用独立临时运行目录与审计库，沿用现有凭据及本机 19890/19891 端口，未改默认安装记录或其他 AI 客户端配置。
- 只写本任务新建的 `CharFloat_M1H_20261012.xlsx` 和无结构化表的视图对照工作簿；原有“工作簿1”保留打开，未向其发送写入请求。
- 用户授予本次工具宿主的磁盘权限后，备份现有两处 WPS 插件目录，使用项目安装器部署候选并重载 Excel 组件。其他插件索引与阻断配置保持不变，备份含凭据，留在隔离运行目录，不进入仓库。

环境和候选标识见 [session.json](session.json)，部署读回见 [02-addon-deployment.json](02-addon-deployment.json)。运行中的加载项上报指纹与源码生成物及部署副本一致，见 [06-runtime-build-confirmed.json](06-runtime-build-confirmed.json)。WPS 插件与独立 CLI 的实际运行证据分别记录，不用磁盘存在代替宿主加载成功。

证据中的个人目录替换为 `<repo>` / `<user-home>` / `<acceptance-runtime>`；脚本按当前仓库与主目录解析必要路径。文本证据与产物的哈希见 [evidence-sha256.json](evidence-sha256.json)，最终结论汇总见 [host-summary.json](host-summary.json)。

## 功能矩阵

| 验收项 | 最终结果与断言 |
|---|---|
| 区域读写、零值、混合值/公式 | 正常；值与公式读回一致，空公式项保留同批值 |
| 非法维度拒绝 | 正常；写前拒绝，前后值不变 |
| 审计与回滚 | 正常；取得真实 auditId，独立测试单元格恢复原值 |
| 格式与样式 | 正常；14pt、粗体、颜色、顶对齐、28pt 行高读回 |
| 富文本片段 | 正常；“保留”两字的粗体和字色由 Characters.Font 读回 |
| 数据有效性 | 列表和数值范围生效；缺候选项在调用宿主前拒绝，既有规则保持 |
| 条件格式与设计语言 | 规则读回含阈值 100，真实设计令牌与可用性信息返回 |
| 列宽与行隐藏 | 自动适配生效；第 30/31 行隐藏读回后再解除 |
| 冻结窗格 | 修正探针参数后，实际 SplitRow=1、SplitColumn=0；不只检查 frozen=true |
| 复制、命名区域、超链接 | 数据复制读回一致；命名区域与超链接存在并落盘 |
| 批注 | 正文与地址读回正确，保存后保留 |
| 筛选排序 | 越界列写前拒绝；金额升序为 50/80/120/150/200，筛选状态可读 |
| 工作表与结构化表 | 创建、复制、重命名、移动和 ListObject 读回正确 |
| 透视表 | 字段布局可读；金额汇总为 600 |
| 图表 | 创建、更新类型/标题、清点、删除测试对照图表和导图可用；首次标题缺失见既有问题 |
| 形状与图片 | 几何及文本读回、编组成员数=2、层级调用和图片插入可用；ZOrderPosition 仍按宿主限制提示 |
| 属性与自定义视图 | 属性落盘；无表对照视图创建及保存重开保留；含表情形见既有问题 |
| 变更检测与事件 | 单格变更可定位；选择 B4 后真实选择读回，重连/心跳上报候选指纹 |
| 打印、预览、PDF 与持久化 | A4 横向、打印区域读回；真实图片与单页 PDF 已检查；数据、公式及对象保存重开一致 |

首轮原始 [suite-results.json](suite-results.json) 为 15 组通过、4 组未通过，保持原样；最终判定结合补验和既有问题复现，不把原始失败改成全量成功。

## 发现并修复的 CLI 问题

PDF 已在 WPS 中导出，桥接侧随后调用 `await import("node:fs")`，在 bytenode 的 V8 上下文报 `A dynamic import callback was not specified`。图表导图的落盘检查也使用同一方式，会将真实落盘误报为失败。

先核对已生成 PDF，未自动重放：见 [211-pdf-side-effect-readback.json](211-pdf-side-effect-readback.json)。用户明确批准后，将 `src/bridge/gateway/excel.ts` 两处改为静态 fs 依赖，保留工具契约和宿主写入逻辑。新增 [实际字节码回归](../../../../tests/bytecode-exports.test.ts)，同一 Electron 运行时编译/加载字节码，覆盖两类导出、缺失文件及宿主调用次数。

- 实际字节码回归通过：[bytecode-regression.log](bytecode-regression.log)。
- 重建、类型及完整回归通过，153/153、失败 0、跳过 0：[构建](cli-fix-build.log)、[类型](cli-fix-typecheck.log)、[完整回归](cli-fix-full-test.log)。
- 142 个工具契约不变：[cli-fix-contract.json](cli-fix-contract.json)。
- 只核对并重启本任务的后台，加载新 CLI 字节码，没有关闭 WPS：[224-fixed-cli-identity.json](224-fixed-cli-identity.json)。
- 真机 PDF 与图表导出均确认落盘：[226-pdf-export-fixed.json](226-pdf-export-fixed.json)、[225-chart-export-fixed.json](225-chart-export-fixed.json)。

## 既有问题与适用条件

### 图表首次创建标题

在 `B1:C6` 数据源及首轮参数下，新图表返回请求标题，但真实 `HasTitle=false`、ChartTitle 为空。使用固定基线提交的原始 `addChart` 函数，在同一测试工作簿、相同数据源及图例参数下复现，见 [222-baseline-exact-chart-add.json](222-baseline-exact-chart-add.json)、[223-baseline-exact-chart-read.json](223-baseline-exact-chart-read.json)。更新标题可生效并保存重开保留，见 [212-chart-update.json](212-chart-update.json)、[213-chart-read-after-update.json](213-chart-read-after-update.json)。

这不是本次搬迁新增的行为；未修改 WPS 图表业务实现。完整产品质量放行仍需另行修复“请求标题未生效却返回成功”的既有问题。较小的独立数据源对照能正常创建标题，不能以其中一例通过掩盖上述复现。

首轮多系列各只有一个数据点，改成折线后无可见线段，属于测试数据拓扑。最终夹具使用独立双数据点并显式设置按列绘制，读回单系列公式与真实折线；只调整测试文档，没有新增工具参数或修改业务源码。

### 含结构化表的自定义视图

含 ListObject 的测试工作簿创建视图后 Count=0，接口明确报失败。原始 `manageWorkbookViews` 函数同样复现；无表对照工作簿可创建并保存重开保留，见 [214-baseline-view.json](214-baseline-view.json)、[209-view-control-add.json](209-view-control-add.json)、[210-view-control-list.json](210-view-control-list.json) 及最终读回。这是本轮实际观察到的使用条件，不将无表成功扩大为所有工作簿可用。

### 探针修正与几何边界

- 编组参数应为对外 schema 的 `shapeNames`，不能直接传宿主内部的 `names`。错误参数在本地被拒绝，补验见 [201-group-corrected.json](201-group-corrected.json)。
- 打印区域按 `$A$1:$S$28` 规范化后比较，不能按裸字符串包含 `A1` 判失败。
- `freezeRowIndex` 是首个未冻结行，2 才表示冻结第 1 行；先解除旧分割，再读回实际 SplitRow/SplitColumn。首轮传 1 的请求不能证明表头冻结，见 [freeze-layout-final.json](freeze-layout-final.json)。
- 图表、图片部分几何由宿主调整，记录实际值，不把请求值当作读回。最终夹具调整宽度与位置避免遮挡，仅用于形成清晰验收截图，不代表所有像素定位场景均已通过。

## 最终文件与读回

[final-readback.json](final-readback.json) 保存前后数据、公式、校验、条件格式、名称、链接、工作表和形状数量的严格对比，以及最新文件 SHA-256。最终实际冻结 1 行、0 列；主文件保留 4 张表、4 个形状、1 张统计图表和 2 个编组成员。

另从 XLSX 包核对实际工作表 XML：`ySplit=1`、无冻结列、`topLeftCell=A2`，WPS 写入的状态为 `frozenSplit`；图表和结构化表 XML 均存在，见 [xlsx-validation.json](xlsx-validation.json)。

- [主测试工作簿](CharFloat_M1H_20261012.xlsx)。
- [无表视图对照工作簿](CharFloat_M1H_ViewControl_20261012.xlsx)。
- [真实区域预览](final-preview.png)。
- [WPS 原生导出 PDF](acceptance.pdf) 与 [渲染页](pdf-page.png)，为单页 A4 横向，见 [pdf-validation.json](pdf-validation.json)。
- [原生图表导图](chart.png)。

退出验收 stdio 客户端后，独立后台与 WPS 连接仍保持，见 [stdio-exit-lifecycle.json](stdio-exit-lifecycle.json)。本次后台保留运行，审计与包含凭据的备份在专用临时运行目录；不把它们复制进仓库。默认安装记录未改，桌面发布包未构建或发布。

## 后续边界

Windows WPS、其他宿主版本、Microsoft Office 真实运行及 Word/PPT 后续模块化均未在本轮验收。M1-H 本机项完成，跨平台项待验；P5 总体或发布质量不得因本报告而自动标记全部完成。

本目录的执行脚本会修改指定测试文件，不能作为普通单元测试自动运行。恢复验收先核对当前文件、服务 PID、构建指纹和专用运行目录，再选择未完成步骤；不要在已有文档上重放整批创建/写入操作。
