# 原生脚本能力扩展

专用工具 → 只读探测 → 原生脚本 → 读回和预览。没有专用工具不能直接判定不支持。脚本是正常的扩展通道，无需仅因使用 JS 再次要求用户授权；操作仍须在用户任务范围内。

WPS 使用 `wps_execute_script`，明确传 component 和精确 workbookName/documentName/presentationName。上下文为 app、wb、doc、pres、wps、params 和 console；支持 async/await。Microsoft 应使用实际工具清单中的脚本入口和对应对象模型，不可照搬 WPS 示例。

可用 `wps_inspect_api` 检查 `wb.Worksheets.Item(1).Shapes` 或 `pres.PageSetup`。部分宿主对象无法完整枚举成员，空列表不是 API 不支持的证据，可针对具体属性进行只读访问。捕获具体异常，区分 API 缺失、目标错误、参数错误和连接问题。

## 先确认变量绑定：按组件只有一个变量有值

脚本形参固定是 `app, doc, wb, pres, wps, params, console`（`wps-addon/src/dispatch.js:297-318`）。**宿主按名字逐个解析，找不到就是 `null`，不会报错**：

| 组件 | 有值的变量 | 其余变量的实测状态 |
|---|---|---|
| WPS 表格 / Excel | **`wb`** | `doc` 为 `null`，`pres` 为 `null` |
| WPS 文字 / Word | **`doc`** | `wb`、`pres` 为 `null` |
| WPS 演示 / PPT | **`pres`** | `wb`、`doc` 为 `null` |

`wps` 是宿主注入的宿主全局（本机实测为 `undefined`；实现里只以 `typeof wps !== "undefined"` 的形式探测 `wps.EtApplication` / `wps.alert` 等成员）。**不要依赖它**，用 `app`、`wb`、`doc`、`pres` 就够了。

> **已证实（[ISS-52](../../../docs/acceptance/2.1.0-p0p1/issues.md)）**：本文档旧版称 `doc`"已自动绑定"。真实宿主下 Excel 场景 `doc` 是 `null`，`doc.Worksheets` 直接抛 `原生脚本执行异常: Cannot read properties of null (reading 'Worksheets')`。**Excel 一律用 `wb`**。

每个批次开头加一句自检，缺变量时立刻报错，比在几十行之后抛空指针好定位：

```js
if (!wb) throw new Error('本批次需要 WPS 表格上下文：wb 为空，请确认 component 与 workbookName');
params = params || {};
```

## 枚举必须传整数

字符串枚举**不报错也不生效**：`range.HorizontalAlignment = 'center'` 被静默吞成左对齐（读回 `-4131`），居中必须传 `-4108`（[ISS-11](../../../docs/acceptance/2.1.0-p0p1/issues.md)：子代理首版所有居中文字跑到左上角，全量重绘一次才修好）。

取值表（对齐、边框、颜色、条件格式、形状类型、箭头线型、PPT 图表与版式、Word 域与版式）见 **[枚举速查表](enumeration.md)**，每条都标了证据等级；取不到的标"未实测"，别猜。

写法上坚持"**写一个，读回一个，断言一个**"：

```js
const want = -4108;                       // xlCenter
shape.TextFrame2.TextRange.ParagraphFormat.Alignment = want;
const got = Number(shape.TextFrame2.TextRange.ParagraphFormat.Alignment);
if (got !== want) throw new Error(`对齐未生效：期望 ${want}，读回 ${got}`);
```

## WPS JS API ≠ VBA 差异清单

AI 的默认先验多是 VBA / Office.js 文档，**直接照抄会抛空指针或静默失真**。下面是已证实的差异（来源：[ISS-12/26](../../../docs/acceptance/2.1.0-p0p1/issues.md)、`p5/mcp-sweep/03-word.md`、本仓库实现）：

| VBA / 常见写法 | WPS JS API 实际 | 正确做法 |
|---|---|---|
| `TextFrame.TextRange`（形状/文本框取文本） | **`undefined`**；继续 `.Font` 会抛 `Cannot read properties of undefined (reading 'Font')`，整个脚本中断（表格形状实测） | 表格 Shapes 用 `TextFrame.Characters()`；PPT 文本框实测可用 `TextFrame.TextRange`（见本仓库实现）；Word 侧未实测 |
| `ParagraphFormat`（形状文本框段落格式） | **不存在**（反射为空成员表） | 用存在的那条路径并先只读探测；不要照抄 VBA |
| `ws.Cells(r, c)` | **不是函数**：`ws.Cells is not a function` | 用字符串地址 `ws.Range("A1")`，或 `ws.Range("A1").Offset(r, c)` |
| `AddChart2` + `SetSourceData` | **默认按行取系列**（一行一个系列），图表与预期不符 | 显式传 `PlotBy = 2`（按列），或逐个赋 `Series.Values` / `Series.XValues` |
| `Shapes.Range([...]).Group()` | **不报错但产出损坏的分组对象**：名称错乱、`GroupItems` 不可枚举（[ISS-14](../../../docs/acceptance/2.1.0-p0p1/issues.md)） | 不要用分组做构图；用命名前缀 + 坐标对齐替代，或先探测确认可用性 |
| `Range("5:6").Hidden = true` | 整行隐藏**静默 no-op**（列路径正常，行为不一致，[ISS-60](../../../docs/acceptance/2.1.0-p0p1/issues.md)） | 逐行 `sheet.Rows.Item(r).Hidden = want` 并逐行读回 |
| `range.Sort(...)` 旧式排序 | 本机 WPS **静默 no-op** | 用 `sheet.Sort.SortFields.Clear()/Add()/SetRange()/Header/Apply()`，并读回校验（[ISS-38](../../../docs/acceptance/2.1.0-p0p1/issues.md)） |
| `Cell.Formula("=SUM(ABOVE)")` / `AutoSum()` | **不生效**，返回 `null`、单元格文本不变（[03-word.md](../../../docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/03-word.md) §3.1） | 脚本里算出结果写静态文本；不要因为"没抛异常"就判定成功 |
| `ExportAsFixedFormat(path, 17)` | 不抛异常也**不落盘**（Word 侧实测） | 交付/预览不要依赖它；用带落盘校验的保存工具 |
| `for (const k in v) v[k]` 全量反射 | **会让 WPS 进程崩溃**（[ISS-89](../../../docs/acceptance/2.1.0-p0p1/issues.md)：3 份崩溃报告、2 份栈逐帧一致：`kso → etcore → etapi → jsetapi → ksojscore`） | **不要自己写反射循环**，用 `wps_inspect_api`（默认只列名字、不求值）；危险成员清单见下节 |

### 反射的危险成员清单（ISS-89 取证结果）

**取证方法**：崩溃前的探测数据已落盘，比对"清单里的表达式"与"实际探到的"——**第 5 条 `wb.Worksheets.Item("Data").Cells` 是崩溃点**（前 4 条 `wb` / `Worksheet` / `Range` 成功，之后全部变成"加载项未连接"）。

**机理**：列成员名（`Object.getOwnPropertyNames`）是安全的；**对成员求值（`obj[name]`）才是扎进宿主原生层的动作**。

| 分级 | 成员 | 说明 |
|---|---|---|
| **已证实危险** | `Cells`、`Rows`、`Columns`、`UsedRange`、`EntireRow`、`EntireColumn` | 整表/整列/整行范围，求值会构造覆盖整表的原生对象 → 进程崩溃 |
| **保守跳过** | `Comment`、`CommentThreaded`、`Comments`、`CommentsThreaded`、`Sort`、`SortFields`、`AutoFilter`、`Filters`、`FormatConditions`、`Validation`、`Names`、`QueryTables`、`Connections`、`ChartObjects`、`PivotCaches`、`PivotTables`、`ListObjects`、`Styles`、`CommandBars`、`CurrentRegion`、`Precedents`、`Dependents`、`SpecialCells` | 触达宿主内部集合的原生 getter；未做崩溃取证，但属最易出问题的一类 |

**`wps_inspect_api` 的护栏**（已实现）：

- **默认 `evaluate: false`**：只列名字，连 `typeof` 都不取（取 `typeof` 同样会触发一次属性访问）
- `evaluate: true` 时按上表跳过危险成员，并在 `skipped` 里逐条说明原因
- `maxMembers`（默认 150）限制求值数量，超出部分记入 `skipped`

**需要上表里的能力时**：不要反射它，**直接调用封装好的工具**——例如 `get_sheet_outline` 的 `sheetState` 已能读回保护/筛选/条件格式，`get_range_styles` 的 `validation` 能读回数据有效性。

另有两条与 API 无关、但同样会静默失真的行为：

- **返回值嵌套过深会丢属性**（[ISS-56](../../../docs/acceptance/2.1.0-p0p1/issues.md)）：原实现 `depth > 2` 直接 `String(val)`，三层以上属性静默变 `undefined`。**已修为**：上限放宽到 6 层，超限节点写成显式占位符并返回 `truncated: true` 与 `truncatedPaths`（不再静默丢数据）。稳妥做法仍是**返回扁平的字符串/数字数组**，或自己 `JSON.stringify` 成字符串再返回。
- **脚本成功 ≠ 写入生效**：Bridge 的包装层可能返回 `success`。调用方必须同时检查 `returnValue` 里的读回值和 `failed`/`pendingCount`，本文档末尾给了约定结构。

## Excel / 表格：矢量绘图（优先用工具，不要手写脚本）

**表格也能画矢量图**，而且 WPS 与 Microsoft 的能力**不一样**——这一点直接决定你能画到什么程度：

| 能力 | WPS 表格 | Microsoft Excel |
|---|---|---|
| 几何形状 / 文本框 | ✅ | ✅ |
| **直线连接符** | ✅ | ❌ 报「当前对象不允许此操作」 |
| **艺术字**（`AddTextEffect`） | ✅ | ❌ 无此 API |
| **分组 / 解组** | ✅ `Shapes.Range([...]).Group()` | ❌ 经工具未打通 |
| 层级调整 | ✅ | ✅ |
| 形状导图 | ❌ 未实现 | ✅ `getAsImage` |
| SVG 形状 | 未实测 | ❌ 本机无 `addSvg` |

**结论：要画复杂矢量图（信息图、流程图、带艺术字标题的看板），WPS 表格比 Microsoft Excel 更能画。**

用工具（`host` 传 `wps` 或 `microsoft`）：

- `excel_add_shape` —— `kind` 可选 `geometric` / `textBox` / `line` / `wordart`（艺术字，WPS）；几何形状名见 `shapeType`（`rectangle` / `rounded_rectangle` / `oval` / `arrow_right` / `flow_chart_decision` 等 30 种，**写错会在错误信息里列出可用值**）
- `excel_list_shapes` —— **画完必须读回**：返回每个形状的名字/类型/位置/尺寸/旋转/填充色/线条色线宽/文字/所在单元格
- `excel_update_shape` —— 移动/改尺寸/旋转/改色/改文字/改名，`action:"delete"` 删除
- `excel_group_shapes` / `excel_ungroup_shapes` —— 分组解组（WPS 可用）
- `excel_set_shape_zorder` —— 层级：`bringToFront` / `sendToBack` / `bringForward` / `sendBackward`

**构图要点**：形状位置用**磅值**（`left/top/width/height`），不是单元格坐标；先画底层的分组框与背景条，再叠内容，最后用 `list_shapes` 核对每个形状的落位——**表上的形状没有"网格"约束，不读回很容易叠在一起**。

## Excel：可编辑矢量绘图

Shapes 是独立于单元格和统计图表的绘图层，可用于流程图、标注、信息卡片。下例是 WPS 对象模型示例，使用前检查当前宿主的 Shapes API。坐标为 pt，依据目标单元格 Left/Top/Width/Height 定位，不是行列号或屏幕像素。

只读探测代码：

```js
const sheet = wb.Worksheets.Item(params.sheetName);
const anchor = sheet.Range(params.address);
return { workbookName: wb.Name, sheetName: sheet.Name,
  left: anchor.Left, top: anchor.Top, width: anchor.Width, height: anchor.Height,
  shapeCount: sheet.Shapes.Count };
```

绘制可编辑矩形的代码（params 传 sheetName、address、shapeName、text）：

```js
const sheet = wb.Worksheets.Item(params.sheetName);
const anchor = sheet.Range(params.address);
const shape = sheet.Shapes.AddShape(1, anchor.Left, anchor.Top, 180, 72);
shape.Name = params.shapeName;
shape.TextFrame.Characters().Text = params.text;
return { name: shape.Name, left: shape.Left, top: shape.Top,
  width: shape.Width, height: shape.Height, shapeCount: sheet.Shapes.Count };
```

`AddShape` 首参是形状类型枚举（矩形 `= 1`），见 [枚举速查表](enumeration.md)。`TextFrame.Characters()` 是 Excel 侧已验证的取文本路径，**没有 `TextFrame.TextRange`**。

连接线、自由曲线可继续探测 `AddConnector`、`BuildFreeform` 等宿主 API（**未实测**；本项目实现里从未调用过 `AddConnector`，见 [07-tool-gap.md](../../../docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/07-tool-gap.md)）。**分组（`Range(...).Group`）不要用于构图**（见上表 ISS-14）。不要把截图当作可编辑矢量。样式、文字接口可能随宿主版本不同；创建后若后续设置失败，先检查已创建对象再继续，不能重新创建整套。

## Excel：画布换算与多元素构图范例

### 画布换算

**已证实**（[ISS-15](../../../docs/acceptance/2.1.0-p0p1/issues.md)，WPS 默认列宽/行高实测）：默认列宽 **48pt**、行高 **16pt**；`A1:U42`（21 列 × 42 行）= **1008 × 672pt**。行列号与 pt 的换算是线性的：

```
x(col)  = anchor.Left + (col   - 1) * 48        // col 从 1 起
y(row)  = anchor.Top  + (row   - 1) * 16
w(cols) = cols * 48        h(rows) = rows * 16
```

例如 1008×672 画布上，`A` 列左边缘 = `0`，`U` 列左边缘 = `20 * 48 = 960`；第 42 行下边缘 = `41 * 16 + 16 = 672`。

**不要只靠 48/16 硬算**（列宽被改过、缩放或自定义行高都会偏）。每个批次开头读一次真实几何，把换算建立在读回值上：

```js
const sheet = wb.Worksheets.Item(params.sheetName);
const area = sheet.Range(params.canvas);            // 例："A1:U42"
const colW = Number(sheet.Columns.Item(1).Width);   // 默认 48pt
const rowH = Number(sheet.Rows.Item(1).Height);     // 默认 16pt
return { left: area.Left, top: area.Top, width: area.Width, height: area.Height, colW, rowH };
```

坐标系原点在画布左上角：元素用 `canvasLeft + 偏移` 定位，完成后再核对整体包围盒不超过 `canvasLeft + canvasWidth` / `canvasTop + canvasHeight`。

### 信息图 / 流程图范例（标题条 + 分组框 + 流程带 + 应用卡 + KPI 条）

下面这套布局对应已交付成品 `sweep图形_钙钛矿产业信息图` 的元素构成——**134 个矢量元素**：标题条 + 三栏分组 + 5 段流程带箭头 + 4 应用卡 + 3 技术路线箭形块 + 4 格 KPI 条 + 页脚（见 [mcp-sweep/README.md](../../../docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/README.md)，渲染图 [shapes-infographic-final.png](../../../docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/shapes-infographic-final.png)）。

> **口径**：结构与元素数量来自已交付成品；下面给出的坐标是"按 1008×672 推算的起始值"，**本轮未在真实宿主复跑**。实测读数只有卡片的 `190.67 × 346.67`（`.scratch/ppt3/03-existing-deck-scan.json`，PPT 卡片）与画布的 48/16。请按内容体量微调，不要当成硬约束。

以 `A1:U42`（1008×672pt）为画布，用 `params.blocks` 传元素清单，**一次批次画完全部元素**：

```js
if (!wb) throw new Error('需要 WPS 表格上下文');
const sheet = wb.Worksheets.Item(params.sheetName);
const canvas = sheet.Range(params.canvas || 'A1:U42');
const L = Number(canvas.Left), T = Number(canvas.Top), W = Number(canvas.Width);
const M = 16;                                   // 统一留白
const done = [], failed = [];

function box(cfg) {                             // 统一创建 + 命名 + 读回
  try {
    const s = sheet.Shapes.AddShape(cfg.type || 1, cfg.l, cfg.t, cfg.w, cfg.h);
    s.Name = cfg.name;
    s.TextFrame.Characters().Text = cfg.text || '';
    done.push({ name: s.Name, l: Number(s.Left), t: Number(s.Top), w: Number(s.Width), h: Number(s.Height) });
    return s;
  } catch (e) { failed.push({ name: cfg.name, error: e.message }); return null; }
}

// 1) 标题条：整幅宽度
box({ name: 'IG_TITLE', l: L, t: T, w: W, h: 44, text: params.title });
// 2) 三栏分组框：等宽三栏，两栏之间留 16pt
const gap = 16, colW = (W - gap * 2) / 3, groupTop = T + 60, groupH = 220;
params.groups.forEach((g, i) => {
  const l = L + i * (colW + gap);
  box({ name: `IG_GROUP_${i + 1}`, l, t: groupTop, w: colW, h: groupH, text: g.title });
});
// 3) 流程带：5 段水平排布，段间留 8pt，箭头落在间隙内
const flowTop = groupTop + groupH + 24, flowH = 56, segs = params.flow.length;
const segW = (W - 8 * (segs - 1)) / segs;
params.flow.forEach((txt, i) => {
  const l = L + i * (segW + 8);
  box({ name: `IG_FLOW_${i + 1}`, l, t: flowTop, w: segW, h: flowH, text: txt });
  if (i < segs - 1) box({ name: `IG_ARROW_${i + 1}`, type: params.arrowType, l: l + segW, t: flowTop + flowH / 2 - 6, w: 8, h: 12, text: '' });
});
// 4) 应用卡：四卡等宽
const cardTop = flowTop + flowH + 24, cardGap = 16, cardW = (W - cardGap * 3) / 4, cardH = 96;
params.cards.forEach((c, i) => box({ name: `IG_CARD_${i + 1}`, l: L + i * (cardW + cardGap), t: cardTop, w: cardW, h: cardH, text: c }));
// 5) KPI 条：四格，红/黄/绿阈值由数据决定
const kpiTop = cardTop + cardH + 24, kpiW = W / 4;
params.kpis.forEach((k, i) => box({ name: `IG_KPI_${i + 1}`, l: L + i * kpiW, t: kpiTop, w: kpiW, h: 48, text: `${k.label} ${k.value}` }));
// 6) 页脚
box({ name: 'IG_FOOTER', l: L, t: kpiTop + 64, w: W, h: 20, text: params.footer });

return { canvas: { l: L, t: T, w: W, h: Number(canvas.Height) }, completed: done, failed: failed.length ? failed : null };
```

要点：

- **命名即索引**：`IG_GROUP_1` / `IG_FLOW_3` / `IG_KPI_2` 这种前缀+序号命名，既是重绘时的定位依据，也是失败后续做的锚点。不要依赖 `Shapes.Item(i)` 的序号（新增/删除会串位）。
- **一次批次画完一组**，不要每个形状一次 MCP 调用；批次返回每个元素的**实际** `left/top/width/height`，而不是回显输入。
- **字号与容量**：文本框高度固定时先估行数（`估算行数 * 字号 * 1.25 <= 内高`），溢出就缩短文案或加高，不要无限缩小字号。
- **失败不重绘整套**：`failed` 非空时只补建失败的元素（按名字判断是否已存在），已成功的元素不要再动。
- **箭形块**：`AddShape` 的箭头类型枚举**未实测**（见枚举表 §1.5）。在确认取值前，用矩形 + 短窄矩形拼近似箭头，或先做一次只读探测。
- **分组**：不要用 `Range(...).Group()`（ISS-14 会产出损坏对象）；需要"看起来是一组"时靠坐标对齐、统一配色和命名前缀。

## PPT：页面与坐标基准

PPT 使用 `pres.PageSetup.SlideWidth/SlideHeight` 获取真实尺寸（本机实测 **960×540**）。按页面比例计算几何，所有结果转为 pt 后写入；字体与版面一起考虑。先读现有形状，设置文本框内边距、换行和字号，读回后检查真实预览。需要拆页或改动内容时结合用户目标判断，不用无限缩小字体掩盖溢出。

要点（来自 `p5/mcp-sweep/02-ppt.md` 与 [ISS-53/55/82](../../../docs/acceptance/2.1.0-p0p1/issues.md)）：

- `wps_ppt_generate_deck` 的内部坐标基准是 **720×405 设计基准**，宿主再按页面尺寸换算；返回体里的 `pageWidth`/`layoutWarnings` 是判断"是否真的换算过"的指纹。自己用脚本排版时，**要么全用设计基准交给缩放，要么全用真实尺寸自己算，不要混用**。
- `layoutIndex` 是 **`ppLayout` 枚举**（`1` 标题幻灯片 / `2` 标题和文本 / `7` 标题和图示），不是"第几个版式"；越界（如 `12`）不报错。见枚举表 §2.3。
- **原生图表**：`wps_ppt_insert_native_chart` 在本机 100% 失败（ISS-80），`generate_deck` 的 chart 布局又有数值写不进去的问题（ISS-75）；图表仍要先读回类型与系列数据再判定成功。

## Word：分节、横向页、书签、域、样式

Word 侧的样式、书签、内容控件、交叉引用、分节、表格行列尺寸、文档属性**都没有专用工具**，只能走脚本（[03-word.md](../../../docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/03-word.md) §11）。下面每个片段都在真实宿主验证过，写到"读回断言"为止才停。

### 样式：新建、应用、判定口径

```js
if (!doc) throw new Error('需要 WPS 文字上下文：doc 为空');
const st = doc.Styles.Add(params.styleName, 1);          // 1 = 段落样式，2 = 字符样式
st.Font.Name = '微软雅黑'; st.Font.Size = 12;
st.Font.Color = 0x404040;                                // RGB
st.ParagraphFormat.LineSpacingRule = 4;                  // 固定行距
st.ParagraphFormat.LineSpacing = 20;
st.ParagraphFormat.SpaceAfter = 6;
st.ParagraphFormat.FirstLineIndent = 21;

const p = doc.Paragraphs.Item(params.paragraphIndex);
p.Style = doc.Styles.Item(params.styleName);
// 判定必须读样式名：读 Range.Font 会拿到继承自基样式的值，会误判“没生效”
return { styleName: p.Style.NameLocal, outlineLevel: Number(p.OutlineLevel) };
```

**已证实**：套上自定义样式后 `p.Range.Font.Name` 可能读回基样式（如 `宋体`）的字体，而 `p.Style.NameLocal` 已是自定义样式名——**"样式名生效"和"直接格式读回"是两套口径**，判生效要读 `Style.NameLocal`；套「标题 3」时 `OutlineLevel` 会同步，可用它确认真的进了大纲。样式属性读回口径同上例：`LineSpacingRule=4`（固定值）、`LineSpacing=20`、`FirstLineIndent=21`、段后 `6`，逐项读回即可断言。

### 分节 + 横向页 + 独立的页眉页脚

```js
const D = doc;
D.Sections.Add(D.Content.End - 1);                       // 新增 1 节
const s2 = D.Sections.Item(2);
s2.PageSetup.Orientation = 1;                            // 1 = 横向（0 = 纵向）
s2.Headers.Item(1).LinkToPrevious = false;               // 必须先断开，否则会连带改第 1 节
s2.Headers.Item(1).Range.Text = '第四部分 · 附录（横向页）';
s2.Footers.Item(1).Range.Text = '附录 A';

const s1 = D.Sections.Item(1);
return { count: D.Sections.Count,
  s1: { o: Number(s1.PageSetup.Orientation), w: Number(s1.PageSetup.PageWidth), h: Number(s1.PageSetup.PageHeight) },
  s2: { o: Number(s2.PageSetup.Orientation), w: Number(s2.PageSetup.PageWidth), h: Number(s2.PageSetup.PageHeight) },
  s1Header: s1.Headers.Item(1).Range.Text, s2Header: s2.Headers.Item(1).Range.Text };
```

**已证实**：横向节读回 `orientation=1, pageWidth=841.9, pageHeight=595.3`，纵向节仍是 `0` / `595.3×841.9`；**没断开链接时节 1 会被连带改**，所以要同时读回两节的页眉来断言隔离成立。
**另注**：`wps_word_page_layout_and_watermark` 的页眉/页脚/水印**只作用于第 1 节**且不提示（[ISS-72](../../../docs/acceptance/2.1.0-p0p1/issues.md)），多节文档的其余节必须自己用上面的脚本处理。

### 书签与交叉引用域

```js
const D = doc;
const target = D.Paragraphs.Item(params.paragraphIndex).Range;
D.Bookmarks.Add('OAB_交付状态', target);

// 域：Type 已决定域名，Code 里不能再写域名
const fRef  = D.Fields.Add(target, 33, 'OAB_交付状态 \\h ');    // REF：交叉引用
const fPage = D.Fields.Add(target, 37, 'OAB_交付状态 \\h ');    // PAGEREF：页码引用
fRef.Update(); fPage.Update();
return { bookmarks: D.Bookmarks.Count, refCode: fRef.Code.Text, refResult: fRef.Result.Text };
```

**已证实**：`Bookmarks.Add(name, 段Range)` 生效，读回范围是**整段**（标题里的制表位与页码也一起进去）；`Fields.Add` 返回 `result="三、验收测试结果\t4"`（REF）与 `"2"`（PAGEREF）。
**两条避坑**：① `Code` 里写成 `' REF OAB_交付状态 \h '` 会得到重复域名 ` REF  REF …`，`result="错误！未定义书签。"` 且**不抛异常**；② 删除被书签覆盖的段落会让书签消失（子代理怀疑是 Range 重叠，**属猜测、未隔离验证**），所以书签要在结构改动完成后再建。

### 表格：行高、底纹、边框

```js
const T = doc.Tables.Item(params.tableIndex);
T.Rows.Item(2).HeightRule = 2;              // 2 = Exactly，1 = AtLeast，0 = Auto
T.Rows.Item(2).Height = 28;                 // 默认 HeightRule=0 时 Height 设多少都读回 0
T.Cell(1, 1).Shading.BackgroundPatternColor = 0xE8EEF7;   // RGB，不是 BGR
T.Rows.Item(1).HeadingFormat = true;        // 跨页重复表头（读回 -1）
return { rowCount: T.Rows.Count, cellCount: T.Rows.Item(1).Cells.Count, uniform: T.Uniform,
  row2Height: Number(T.Rows.Item(2).Height), shade: T.Cell(1, 1).Shading.BackgroundPatternColor,
  header: T.Rows.Item(1).HeadingFormat };
```

**已证实**：列宽/行高单位是 **pt**；`HeadingFormat` 读回 `-1`；`Shading.BackgroundPatternColor` 是 **RGB**（与 Excel 的 BGR 相反）。**表内公式不支持**：`Cell.Formula("=SUM(ABOVE)")` 与 `AutoSum()` 都返回 `null` 且文本不变——脚本里算出结果写静态文本，并读回断言。

## 批量执行与返回

把一页、一张表或一组相关形状作为批次边界。用 params 传操作列表，在脚本内循环，不为每个对象单独发起 MCP 调用。所需 API 已在当前宿主验证时复用，无需每次反射。

执行前校验整批目标和输入；每步成功后记录对象标识。脚本可返回以下普通 JSON 结构（这是脚本自定义返回格式，不是新增 MCP 参数）：

```json
{"completed":[{"id":"实际对象标识","left":50,"top":30}],"failed":null,"pendingCount":0}
```

**返回结构要扁平**：嵌套超过两层会丢属性（ISS-56），深层结构请自己 `JSON.stringify` 成字符串再返回。失败时记录当前对象、步骤和异常文本，并停止依赖该步骤的修改。completed 中使用宿主读回值；不返回完整宿主对象、全文或无关数据。Bridge 的脚本包装层可能仍返回 success，调用方必须同时检查 returnValue 中的 failed/pendingCount。响应超时导致整个结果丢失时，按目标标识重新读取，不依靠内存完成列表猜测。

不要为省调用合并本来需要人工决策的内容，也不要将需要审计回滚的单元格写入移到脚本。

## 高频陷阱：这几条不写下来，每个 Agent 都会重踩一遍

以下每一条都是真机上被反复踩到的，不是理论。**坑的共同点是「调用成功但结果不对」**——
只看返回值判断不了，必须按这里写的方式核对。

### 1. 图表数据源必须是**交叉表**，不能用平铺明细表

平铺表（`月份 | 区域 | 营收`，18 行）直接当 `dataRange`，宿主会**按行生成 18 个系列**——
图例挤成一团、颜色随机、完全不可读。

正确做法是先做**交叉表**（行=月份、列=区域），再基于它画图：

```
月份   华东   华南   华北
1月   1280   960    740
2月   1420  1080    820
```

→ 得到 3 个系列（华东/华南/华北）× 6 个类别（月份）。
**核对方法**：建完图读回 `seriesCount`，等于"列数-1"才对；远大于列数就是踩了这个坑。

### 2. 图表数据源**不支持跨工作表**

`dataRange` / `sourceAddress` 传 `明细!A1:C19` 一律报
`Parameter type error source (arg 0)`；`dataRanges` 只收字符串数组，传对象报类型错。

**所以：图表必须建在数据所在的那张表上**。想把明细数据画到独立看板页，得先把交叉表
写到看板页（可以放远一点的空列里），再从那里画图。

### 3. `patch_cells` 的参数是 `address` + **二维数组**，不是逐格对象

```jsonc
// ✅ 正确
{ "name": "excel_patch_cells", "arguments": {
    "address": "A1:B3", "values": [["产品","销量"],["甲",120],["乙",80]] }}

// ❌ 错误：没有 cells 参数，会被 schema 拒掉
{ "arguments": { "cells": [{ "address": "A1", "value": "产品" }] } }
```

**踩坑代价特别大**：写数据失败后如果**不检查写入结果**，后面所有步骤都会在空区域上失败，
而错误信息指向的是"找不到工作表/区域不存在"，**完全指不到真正的错因**。

### 4. 用 `Item(name)` 探测"删没删掉"会得到**假阳性**

宿主集合对象的 `Item(name)` 在名字**不存在时返回 `null`，不抛异常**：

```js
// ❌ 错：删掉之后这一句仍然不抛错 → 判定"还在"
try { wb.Names.Item("x"); stillExists = true; } catch (e) { stillExists = false; }

// ✅ 对：枚举集合逐个比对
let stillExists = false;
for (let i = 1; i <= wb.Names.Count; i++) if (String(wb.Names.Item(i).Name) === "x") { stillExists = true; break; }
```

同样的陷阱存在于 `Shapes.Item(name)`、`CustomDocumentProperties.Item(name)`、`Worksheets.Item(name)`。
**判存在性一律用「集合 `Count` + 逐项枚举」，不要用 `Item(name)`。**

### 5. 写入类操作一律**读回宿主真实状态**再判定成功

`success: true` 只说明"函数返回了"，不说明"改生效了"。本仓库已发现多处
**返回体回显请求值、实际没生效**的情况（对齐方式、条件格式公式、保存到已存在路径…）。

**判据只有一条**：改完之后用脚本或读回工具**直接问宿主**，拿到真实值再比对。
不一致就当作失败处理，不要因为 `success` 是 true 就往下走。
