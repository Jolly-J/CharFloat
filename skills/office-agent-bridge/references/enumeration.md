# 枚举常量速查表

本表回答"该传哪个整数"。`wps_inspect_api` **只反射成员名与方法名，不返回枚举常量**（[issues.md ISS-13](../../../docs/acceptance/2.1.0-p0p1/issues.md)），所以这些值只能靠文档固定下来。**不要用字符串枚举猜**——见下一节。

## 铁律：枚举一律传整数

```js
// ✗ 不报错，但被静默吞掉落到左对齐（读回 -4131）
shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 'center';
range.HorizontalAlignment = 'center';

// ✓ 居中
range.HorizontalAlignment = -4108;
```

已证实（2026-09-22 真实宿主实测，[ISS-11](../../../docs/acceptance/2.1.0-p0p1/issues.md)）：传字符串 `'center'` **既不抛异常也不生效**，赋值被吞成左对齐 `-4131`；子代理首版全部居中文字跑到左上角，只能全量重绘一次。所以要**写完立刻读回断言**：

```js
const want = -4108;
range.HorizontalAlignment = want;
const got = Number(range.HorizontalAlignment);
if (got !== want) throw new Error(`对齐未生效：期望 ${want}，读回 ${got}`);
```

同一条规则适用于 `VerticalAlignment`、`LineStyle`、`Type`、颜色等一切枚举型属性。

## 证据等级（每行都标）

| 标记 | 含义 |
|---|---|
| **实测** | 在真实 WPS 宿主写入并读回验证过（证据见 issues.md / `p5/mcp-sweep/`） |
| **代码可证** | 本仓库实现里直接使用该常量且已有真实宿主复验（如 7/7、14/14 复验批次） |
| **未实测** | 该枚举名在本机未验证取值，**不要照抄**，需要自己只读探测确认后再用 |

---

## 1. Excel / WPS 表格

### 1.1 单元格对齐

| 属性 | 常量 | 值 | 等级 |
|---|---|---|---|
| `Range.HorizontalAlignment` | 左 | `-4131` | **实测**（`'center'` 被吞后读回即此值） |
| | 居中 | `-4108` | **实测** |
| | 右 | `-4152` | **代码可证**（`wps-addon/src/excel.js:206` 读回映射） |
| | 常规 | `1` | **代码可证**（同上） |
| `Range.VerticalAlignment` | 靠上 | `-4160` | **代码可证**（`excel.js:205`） |
| | 居中 | `-4108` | **代码可证**；`excel.js:477` 直接使用 |
| | 靠下 | `-4107` | **代码可证**（`excel.js:205`） |

> 注意：`wps_format_cells` 的 `verticalAlignment` 只实现了 `center` 一个分支（[07-tool-gap.md](../../../docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/07-tool-gap.md) 已列），垂直靠上/靠下要走脚本并传整数。

### 1.2 边框线型与颜色

| 用途 | 写法 | 取值 | 等级 |
|---|---|---|---|
| 线型 | `borders.LineStyle` | `1` = 连续实线（xlContinuous） | **代码可证**（`excel.js:501`） |
| 颜色 | `Font.Color` / `Interior.Color` / `Format.Fill.ForeColor.RGB` | **BGR 字节序**：`(b<<16)\|(g<<8)\|r` | **代码可证**（`excel.js:1274` 起，变量名即 `bgr`） |
| 白色 | `Font.Color = 16777215` | `0xFFFFFF`（BGR 下对称，看不出顺序） | **代码可证** |
| 颜色读回 | `get_range_styles` 返回的十六进制 | 已由实现反算回 `#RRGGBB` | **代码可证**（`excel.js:1425` `excelColorToHex`） |

> **与 Word 相反**：Excel 侧是 BGR，Word 侧 `Shading.BackgroundPatternColor` 是 **RGB**（见 §3.2）。两边都实测过，别互抄。

### 1.3 条件格式（`Range.FormatConditions`）

| 用途 | 写法 | 取值 | 等级 |
|---|---|---|---|
| 规则类型 | `FormatConditions.Add(Type, Operator, Formula1, Formula2)` | `1` = xlCellValue | **实测**（`04-sheet-advanced.md` 阈值规则读回 `type=1`） |
| | `FormatConditions.Add(2, 0, '=E4>AVERAGE(...)')` | `2` = xlExpression（公式规则） | **实测**（同上，`type=2`） |
| | `AddColorScale(2)` | 色阶规则 `type=3` | **实测**（`type=3`、`ColorScaleCriteria.Count=2`） |
| | `AddDatabar()` | 数据条规则 `type=4` | **实测**（`type=4`） |
| 运算符 | `Operator` | `6` = xlLessThan | **实测**（`Operator=6`，公式 `"=0.9"`） |

> 脚本通道才有的能力：公式规则、图标集。MCP 工具的 `ruleType` 枚举只有 `cell_value` / `color_scale` / `data_bar`（[04-sheet-advanced.md](../../../docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/04-sheet-advanced.md)）。

### 1.4 排序 / 图表

| 用途 | 写法 | 取值 | 等级 |
|---|---|---|---|
| 排序字段 | `sheet.Sort.SortFields.Add(range, Order, Key)` | `0` = 升序、`1` = 降序 | **实测**（`.scratch/sheet3/sort_fix.js`，对应 ISS-38 修复后唯一有效路径） |
| 表头 | `sort.Header` | `2`（含表头）；修复实现用 `1` | **实测**（探针脚本 `Header = 2`） |
| 图表类型 | `Shapes.AddChart2(-1, Type, ...)` | `51` 簇状柱 / `57` 簇状条 / `4` 折线 / `-4120` 环形 / `-4169` 散点 | **实测**（`p5/mcp-sweep/README.md` 成品读回：条形 57 / 折线 4 / 环形 -4120 / 散点 -4169 / 柱状 51） |
| 散点必须走脚本 | `AddChart2(201, -4169, ...)` | `201` = xlXYScatter | **实测**（[ISS-17](../../../docs/acceptance/2.1.0-p0p1/issues.md)：工具侧 `scatter` 会被降级成 51） |

> `PlacedChartType` 与 `AddChart2` 的用法见 [native-scripting.md](native-scripting.md) §差异清单：**`SetSourceData` 默认按行取系列，必须显式 `PlotBy=2`**。

### 1.5 Shapes 形状类型

| 用途 | 取值 | 等级 |
|---|---|---|
| `Shapes.AddShape(Type, left, top, w, h)` 矩形 | `1`（msoShapeRectangle） | **实测**（`native-scripting.md` 原例 + `03-word.md` 表格水印用的 `AddTextEffect` 同族）；`shapes-infographic-final.png` 成品的 134 个元素即由该调用族构成 |
| 圆角矩形 / 椭圆 / 箭头 / 连接符 / 自由曲线 | **未实测** | 需自行只读探测；`Shapes` 集合可反射出方法与成员名，但**取不到常量值**，别猜 |

---

## 2. PowerPoint / WPS 演示

### 2.1 形状与文本框类型（`Shapes.AddShape` / `AddTextbox` 首参）

| 类型 | 值 | 等级 |
|---|---|---|
| 矩形 | `1` | **实测**（`.scratch/ppt3/03-existing-deck-scan.json`：`type=1` 全为卡片底/色条；`ppt.js:78` 同值） |
| 圆角矩形 | `5` | **代码可证**（`ppt.js:660`） |
| 椭圆 | `9` | **代码可证**（`ppt.js:661`） |
| 箭头 | `13` | **代码可证**（`ppt.js:662`） |
| 文本框 | `17` | **实测**（同扫描：`type=17` 全为文字形状） |
| `AddTextbox(Orientation, ...)` | `1` = 横排 | **代码可证**（`ppt.js:45`；`AddTextEffect` 同族，`word.js:797` 也用 `0/1` 语义参数） |

### 2.2 图表类型（`Shapes.AddChart` / `AddChart2` 的 Type）

| 类型 | 值 | 等级 |
|---|---|---|
| 折线 | `4` | **代码可证**（`ppt.js:126`） |
| 饼图 | `5` | **代码可证**（`ppt.js:127`） |
| 簇状柱形 | `51` | **代码可证**（`ppt.js:125`） |
| 堆积柱形 | `52` | **代码可证**（`ppt.js:129`） |
| 簇状条形 | `57` | **代码可证**（`ppt.js:128`） |
| 堆积条形 | `58` | **代码可证**（`ppt.js:130`） |
| 复合条饼图 | `68` | **代码可证**（`ppt.js:131`） |

> ⚠️ **PPT 的 `4`/`5` 与 Excel 的 `4`/`5` 语义不同，`51`/`57` 才重合**——跨组件拷贝脚本时逐个核对。
> ⚠️ `wps_ppt_insert_native_chart` 在本机 **100% 失败**（[ISS-80](../../../docs/acceptance/2.1.0-p0p1/issues.md)：`AddChart/AddChart2/AddOLEObject` 返回 `null` 且不建形状）；PPT 原生图表要走 `generate_deck` 的 `chart` 布局，且其数值写入另有缺陷（[ISS-75](../../../docs/acceptance/2.1.0-p0p1/issues.md)）。

### 2.3 版式枚举（`layoutIndex` 实为 `ppLayout`，不是版式序号）

| 值 | 语义 | 等级 |
|---|---|---|
| `1` | 标题幻灯片 | **实测**（`.scratch/ppt3/21-layoutindex-align.json` 读回 `layout=1 clName=标题幻灯片`） |
| `2` | 标题和文本 | **实测**（同上，`layout=2`） |
| `7` | 标题和图示或组织结构图 | **实测**（同上，`layout=7`） |
| `12` | 越界但**不报错**（[ISS-82](../../../docs/acceptance/2.1.0-p0p1/issues.md)） | **实测** |

> 其余 `ppLayout` 取值**未实测**；`21-layoutindex-align.json` 里只验证了 1/2/7 与越界 12。

---

## 3. Word / WPS 文字

### 3.1 版式与结构

| 属性 | 取值 | 等级 |
|---|---|---|
| `PageSetup.Orientation` | `0` = 纵向、`1` = 横向 | **实测**（`03-word.md` §4.2：`orientation=1` → `841.9×595.3`，节 1 仍 `0` / `595.3×841.9`） |
| `Rows.Item(n).HeightRule` | `0` = Auto、`1` = AtLeast、`2` = Exactly | **实测**（§3：**默认 0 时 `Height` 设多少都读回 0，必须先改 rule**） |
| `Columns.Item(n).Width` / `Rows.Item(n).Height` | 单位 **pt**（100/150/80/46/70、28、32） | **实测**（§3） |
| `Rows.Item(n).HeadingFormat` | `-1` = true（跨页重复表头） | **实测**（§3 读回 `-1`） |
| `Table.Borders.Item(idx)` | `-1` = 上/下外框、`-3` = 左/右外框 | **代码可证**（`word.js:614-619`，三线表实现） |
| `Borders.OutsideLineWidth` / `InsideLineWidth` | 线宽枚举，实测用 `12` / `4` | **实测**（§3） |
| `Styles.Add(Name, Type)` | `1` = 段落样式、`2` = 字符样式 | **实测**（§2.1 读回 `type=1` / `type=2`） |
| `ContentControls.Add(Type, Range)` | `0` | **实测**（§5.3 读回 `type:0`，Title/Tag 可写可读） |
| `ComputeStatistics(n)` | `2` = 页数（wdStatisticPages） | **实测**（§7.4：`pages=8`） |
| `BuiltInDocumentProperties.Item(n)` | 用数字索引逐项读写 Title/Subject/Keywords/Category/Comments | **实测**（§1 全部逐字读回） |
| `CustomDocumentProperties.Add(name, LinkToContent, Type, Value)` | `Type=4`（字符串） | **实测**（§1：「验收批次」= `2.1.0-p0p1-p5`） |

### 3.2 颜色：**RGB，不是 BGR**

```js
cell.Shading.BackgroundPatternColor = 0xE8EEF7;   // 读回 15265527 = 0xE8EEF7
```

**实测**（`03-word.md` §3.3）：脚本直写 `0xE8EEF7` 读回 `15265527`（即值本身）；工具路径 `#F1F5F9` 也走同一顺序。**与 Excel 侧习惯相反**（Excel 是 BGR），别照搬。

### 3.3 域类型（`doc.Fields.Add(range, Type, Code, PreserveFormatting)`）

| 用途 | `Type` | 等级 |
|---|---|---|
| `AutoNum`（表格内编号） | `-1` | **实测但无效**：域被创建，单元格只留 `\u0015` 占位，**不渲染数字**（§3.2） |
| `REF`（交叉引用） | `33` | **实测**（§5.4 域码 ` REF OAB_验收结果 \h  \* MERGEFORMAT`，结果正确） |
| `PAGEREF`（页码引用） | `37` | **实测**（§5.4 `Update()` 后 `result="2"`） |
| `PAGE` / `NUMPAGES` | **未实测**（成品页脚里的 PAGE/NUMPAGES 域是宿主既有的，**没有记录其 Type 数字**） | 需要用 `Fields.Add` 插入时，自行只读探测确认 Type |

> **两条实测避坑**：
> 1. `Code` 里**不能重复写域名**——传 `' REF OAB_验收结果 \h '` 会得到 ` REF  REF …` 且 `result="错误！未定义书签。"`，**不报任何异常**；只写参数部分 `'OAB_验收结果 \h '` 才对（§5.4）。
> 2. `doc.ExportAsFixedFormat(path, 17)` **不抛异常也不落盘**（§7.4），不能作为交付或预览的验证手段；桥接侧 `wps_word_save_document` 的导出已加落盘校验，以它的返回为准。

### 3.4 分节与页眉层

| 用途 | 写法 | 等级 |
|---|---|---|
| 新增一节 | `doc.Sections.Add(doc.Content.End - 1)` | **实测**（§4.1） |
| 断开与上节的页眉链接 | `sec.Headers.Item(1).LinkToPrevious = false` | **实测**（§4.2） |
| 节内页眉/页脚 | `sec.Headers.Item(1).Range.Text` / `sec.Footers.Item(1).Range.Text` | **实测**（§4.2；`DifferentFirstPageHeaderFooter` 由工具的 `differentFirstPage` 打开） |
| 跨页水印（宿主工具只落第 1 节正文层） | 删正文层 WordArt 后在 `sec.Headers.Item(1)` 重建 | **实测**（§7.1，见 [ISS-58](../../../docs/acceptance/2.1.0-p0p1/issues.md)） |

---

## 4. 拿不准时的探测方法（安全边界）

`wps_inspect_api` 能列出**成员名**，但不给常量值，也不能给签名（`03-word.md` §11：`Cell.Formula` 的最多参数个数只能靠试错）。

```js
// 受控探测：一次只读一个属性，且只读“已经存在”的对象
const s = wb.Worksheets.Item(params.sheetName).Shapes.Item(1);
return { type: Number(s.Type), lineVisible: !!s.Line.Visible, weight: Number(s.Line.Weight) };
```

**必须避免**：对每一个成员逐个求值的"全量反射"（`for (const k in v) v[k]`）。已证实这会让 **WPS 进程崩溃**（[ISS-89](../../../docs/acceptance/2.1.0-p0p1/issues.md)：3 份崩溃报告、2 份栈逐帧一致，`jsetapi → etcore`），崩溃后组件掉线且桥接不会主动报"疑似崩溃"。**具体哪些成员危险目前没有安全清单**——所以只探你真正要用的那一个属性，并把它写进批次返回里复用。
