# 03 Word 高级能力实测（WPS 文字）

- 目标文档：已打开的 `测试文字文稿.docx`（含初稿：封面、多级大纲、表格 ×2、页眉、水印、页脚页码、批注 ×1、修订 ×2）
- 工作目录：`/Users/Python/Office Agent Bridge`
- 脚本 / 导出物目录：`.scratch/word3/`
- 状态：**已完成**（每完成一步即时追加；两处早期误判已在 §3.3 显式更正）

## 待填

- [x] 基线读回（文档属性、段落/表格/样式清单）
     基线（脚本只读探测，2025 实测）：
     - 文档：`测试文字文稿.docx`，114 段 / 2 表 / 1027 字；**1 节**，纵向 595.3×841.9pt
     - 页眉「Office Agent Bridge 2.1 版本发布与验收说明」；页脚「第 3 页 / 共 5 页」
     - 样式总数 **478**，实际 InUse **13** 个（正文/标题1-3/批注文字/目录1-3/页脚/页眉/普通表格/网格型/默认段落字体）
     - 书签 **0**、内容控件 **0**、索引 0、图文框 2（水印）、批注 1、修订 2
     - 域 27 条：TOC(13) ×1、HYPERLINK(88) ×13、PAGEREF(37) ×13
     - 表格：表1 5×5、表2 5×4（均 Uniform）
     - 内置属性：Author=jolin，Title/Subject/Keywords/Category/Comments 全为空
- [x] 样式体系：标题样式 / 正文样式 / 自定义样式新建+应用
- [x] 表格高级：合并单元格、边框底纹、列宽行高、表内公式或编号
- [x] 多节与分页：分节符、分节页眉页脚差异、横向页
- [x] 内容控件 / 书签 / 交叉引用
- [x] 查找替换高级用法：通配符、格式替换
- [x] 文档属性：标题、作者、主题等元数据
- [x] 汇总判定表（§8）

---

## 1. 文档属性（元数据）

**MCP 工具清单里没有任何文档属性工具**（`wps_word_*` 12 个定义中无 metadata / properties 分支），只能用 `wps_execute_script` 走原生 API。

| 动作 | 通道 | 返回 | 读回判定 |
|---|---|---|---|
| 写内置属性 Title/Subject/Keywords/Category/Comments | script `BuiltInDocumentProperties.Item(x).Value = ...` | 无异常 | ✅ 全部生效（读回值与写入值逐字一致） |
| 读内置属性 Author | script | `jolin` | ✅ 只读属性可见 |
| 新增自定义属性「验收批次」 | script `CustomDocumentProperties.Add(name,false,4,value)` | 无异常 | ✅ 生效，读回 `验收批次=2.1.0-p0p1-p5` |

读回原文（`BuiltInDocumentProperties`）：

```
Title    = "Office Agent Bridge 2.1 版本发布与验收说明"
Subject  = "2.1.0 P0/P1 阶段验收材料"
Keywords = "WPS;Office Agent Bridge;验收;MCP"
Category = "阶段验收"
Comments = "由文档高级能力实测生成"
Author   = "jolin"（只读）
CustomDocumentProperties = KSOProductBuildVer / ICV / 验收批次=2.1.0-p0p1-p5
```

> 结论：能力**支持**，但**只能靠脚本**；MCP 无对应工具，schema 也无处传元数据。

## 2. 样式体系

MCP 同样**没有样式工具**（无 style / 样式 相关 action 或参数；`wps_word_format_document` 是**直接格式**，不是样式）。

### 2.1 新建自定义样式（script `doc.Styles.Add`）

| 样式名 | 类型 | 写入属性 | 读回 |
|---|---|---|---|
| `OAB 实测正文` | 段落(1) | 微软雅黑 12pt / 非加粗 / 色 0x404040 / 固定行距 20pt / 段后 6pt / 首行缩进 21pt | ✅ `font=微软雅黑, farEast=微软雅黑, size=12, bold=0, color=4210752, lineRule=4, lineSpacing=20, firstLineIndent=21, builtIn=false, inUse=true` |
| `OAB 强调标记` | 字符(2) | 微软雅黑 / 加粗 / 色 0xC00000 | ✅ `type=2, bold=-1, color=12582912, inUse=true` |

### 2.2 应用样式

| 目标 | 写法 | 读回判定 |
|---|---|---|
| 第 23 段 | `p.Style = Styles.Item("OAB 实测正文")` | ✅ `style="OAB 实测正文"`，字号 12 生效 |
| 第 24 段整段 Range | `p.Range.Style = ...` | ✅ `style="OAB 实测正文"`，字体=微软雅黑 |
| 第 102 段套「标题 3」 | `p.Style = Styles.Item("标题 3")` | ✅ `style="标题 3"`，`OutlineLevel=3`（说明能真正进大纲，不是假套用） |
| 改基样式「正文」字体 | `Styles.Item("正文").Font.Name = "宋体"` + Size=12 | ✅ 读回 `宋体/12` |

**踩坑记录（重要）**：第 23 段套上自定义样式后，`p.Range.Font.Name` 读回是 `宋体`——因为它继承了被我同时改过的「正文」基样式；第 24 段读回 `微软雅黑`。这不是样式没生效（`p.Style.NameLocal` 已是自定义样式名），而是**「样式名生效」和「Range 直接格式读回」是两套口径**，判定时必须读 `Style.NameLocal`，只看 `Range.Font` 会误判。

## 3. 表格高级能力

以**表 2（5×4）**为试验田（初稿已有），全部通过 script 原生 API 完成。

| 能力 | 写法 | 读回判定 |
|---|---|---|
| 加列 | `T.Columns.Add(T.Columns.Item(4))` | ✅ 4 列 → 5 列 |
| 加行 | `T.Rows.Add()` ×2 | ✅ 5 行 → 7 行 |
| **合并单元格（跨列标题）** | `T.Cell(1,1).Merge(T.Cell(1,5))` | ✅ `Rows.Item(1).Cells.Count = 1`，`T.Uniform` 由 `true` → `false`，单元格文本合成 `2.2 模块交付状态 · 高级表格实测汇总` |
| **合并单元格（跨列合计）** | `T.Cell(7,1).Merge(T.Cell(7,4))` | ✅ `Rows.Item(7).Cells.Count = 2`（原 5） |
| 底纹 | `cell.Shading.BackgroundPatternColor = 0xE8EEF7` | ✅ 读回 `15265527`(=0xE8EEF7)，**注意是 RGB 顺序，不是 BGR** |
| 边框 | `T.Borders.Enable=true` + `OutsideLineWidth=12` / `InsideLineWidth=4` | ✅ 读回 outside=12 / inside=4 |
| 列宽 | `T.Columns.Item(i).Width = 100/150/80/46/70` | ✅ 读回逐列一致，`T.PreferredWidth` 388→446 |
| 行高 | `Rows.Item(2).HeightRule=2`(Exactly)+`Height=28`；`Rows.Item(1).HeightRule=1`(AtLeast)+`Height=32` | ✅ 读回 `{height:28, rule:2}` / `{height:32, rule:1}`。**注意 `HeightRule` 默认 0(Auto) 时 `Height` 设多少都读回 0**，必须先改 rule |
| 表头跨页重复 | `Rows.Item(2).HeadingFormat = true` | ✅ 读回 `-1`(true) |
| 表格样式 | `T.Style = "网格型"` | ⚠️ 赋「网格型」时抛异常，见下 |

### 3.1 表内公式：**不支持（且是"静默无效"型失败）**

```
cell.Formula("=SUM(ABOVE)", null, 0)   → 返回 null，单元格文本不变（仍为 \u0007 空）
cell.Formula("=SUM(ABOVE)", 0, "0")    → 抛 "Too many parameters."
cell.Formula("=SUM(ABOVE)", true, 0)   → 抛 "Too many parameters."
cell.Formula()                          → 返回 null，无变化
cell.AutoSum()                          → 返回 null，单元格文本不变
手动写 "=SUM(ABOVE)" 到 cell.Range.Text  → 文本原样保留，Range.Fields.Count = 0（没变成域）
```

读回实测：`afterWrite="=SUM(ABOVE)"`、`fields=0`、`afterUpdate` 仍是纯文本、`fieldErr="Cannot read properties of null (reading 'Type')"`。

**判定**：WPS for Mac 的 `Cell.Formula` 只接受 ≤2 个参数且**不产生任何效果**，`AutoSum` 同样无效，公式域也无法通过 `Range.Text` 写入。**这就是"返回 success 但实际没生效"的典型**——若只看 `Formula()` 没抛异常，会误判为成功。
**绕过**：脚本算出结果后写入静态文本（本次写入 `合计 296.2`，由 `100+0+99.4+96.8` 求和得出，✅ 读回一致）。

### 3.2 表内编号

| 方式 | 结果 |
|---|---|
| 静态序号列（写入 1/2/3/4） | ✅ 生效，读回 `["1","2","3","4"]` |
| `doc.Fields.Add(range, -1, "AutoNum", false)` | ⚠️ 单元格读回 `\u0015`（域占位符）——**域被创建了但不渲染数字**，表格内 AutoNum 不可用 |

### 3.3 `wps_word_manage_table` 三个 action 的**独立**实测（结论：都生效）

> 修正说明：我一开始拿"表 1 第 1 格"当读回对象，但那次调用其实没落到同一格上，据此写过"success 但没生效"是**误判**；下面是对同一格做的前后对比，结论已改正。

| action | 调用 | 前 → 后（native 读回同一格） | 判定 |
|---|---|---|---|
| `format_cell` | `{cellRow:2, cellColumn:2, cellFormat:{bold:true, fontSizePt:14, fontName:"微软雅黑", backgroundColor:"#F1F5F9"}}` | `shading 15265527 → 16381425`、`size 10.5 → 14`、`bold 9999999 → -1` | ✅ 全部生效 |
| `merge_cells` | `{mergeRange:{startRow:3,startCol:1,endRow:3,endCol:2}}` | `Rows.Item(3).Cells.Count 5 → 4`，格内文本由 `加载项入口与源码一致性` 变为 `加载项入口与源码一致性无漂移` | ✅ 生效 |
| `add_row` | `{tableIndex:2}` | `rowCount 7 → 8` | ✅ 生效 |

**颜色不是 BGR**：`hexToExcelColor('#F1F5F9')` = `(b<<16)|(g<<8)|r` = `16381425`，而脚本直写 `0xE8EEF7` 读回 `15265527` = `0xE8EEF7` 本身 → 两条路径都是 **RGB（r 在最低字节）**，表面看像 BGR 只是数值巧合，别照搬 Excel 的 BGR 经验。

## 4. 多节与分页

### 4.1 新增分节符（script `doc.Sections.Add`）

| 动作 | 写法 | 读回 |
|---|---|---|
| 新增 1 节 | `D.Sections.Add(D.Content.End - 1)` | ✅ `Sections.Count` 1 → 2 |

### 4.2 第 2 节横向页 + 独立页眉页脚

| 动作 | 写法 | 读回 |
|---|---|---|
| 第 2 节横向 | `s2.PageSetup.Orientation = 1` | ✅ `orientation=1, pageWidth=841.9, pageHeight=595.3`（第 1 节仍 0 / 595.3×841.9） |
| 断开与上节的页眉链接 | `s2.Headers.Item(1).LinkToPrevious = false` | ✅ 读回 `false`，第 2 节页眉文本独立 |
| 第 2 节页眉 | `s2.Headers.Item(1).Range.Text = "..."` | ✅ 读回=`第四部分 · 附录（横向页）` |
| 第 2 节页脚 | `s2.Footers.Item(1).Range.Text = ...` | ✅ 读回=`附录 A · 高级能力实测原始记录` |
| 第 1 节页眉是否被连带改 | 只读 | ✅ 第 1 节页眉仍=`Office Agent Bridge 2.1 版本发布与验收说明`（**分节隔离成立**） |

## 5. 书签 / 内容控件 / 交叉引用

### 5.1 书签（script `doc.Bookmarks.Add`）

| 动作 | 读回 |
|---|---|
| `Bookmarks.Add("OAB_交付状态", 段落Range)` | ✅ `Bookmarks.Count=2`，读回 `["OAB_交付状态","OAB_验收结果"]` |
| `Bookmarks.Add("OAB_验收结果", 标题Range)` | ✅ 读回 `{start:565,end:636,text:"三、验收测试结果\t4"}`（**标题里的制表位与页码也一起进去了**，书签范围是整段而非纯标题文字） |

### 5.2 `wps_word_write_content` 的 `location:"bookmark"` —— ⚠️ **放错位置**

调用：`{location:"bookmark", targetBookmark:"OAB_交付状态", content:"【书签定位写入测试】..."}`
返回：`success=true, insertedLines=1, location="bookmark"`

读回结果：
- 新段落落在 **第 13 段**（目录区域内），样式=`目录 1`，**不是**书签处
- 书签 `OAB_交付状态` 的位置 `{492,567}` 完全没变
- 反而把**另一个**书签 `OAB_验收结果` 的范围从 `{565,636}` 撑大成 `{567,689}`，文本变成了「【书签定位写入测试】…三、验收测试结果　4」

**判定：`location:"bookmark"` 返回 success，但内容没有插进书法处，而是污染了相邻书签的范围。** 目录被插入了一行非目录内容。

### 5.3 内容控件（ContentControls）：**API 存在且可用**

| 动作 | 读回 |
|---|---|
| `D.ContentControls.Add(0, 段落Range)` | ✅ `ContentControls.Count=1` |
| 写 `Title` / `Tag` | ✅ 读回 `{title:"OAB 实测控件", tag:"oab-test", type:0}` |

> MCP 无任何内容控件工具；`wps_word_write_content` 的 location 枚举里也没有 contentControl。

### 5.4 交叉引用（REF / PAGEREF 域）

| 写法 | 结果 |
|---|---|
| `Fields.Add(rng, 3, ' REF OAB_验收结果 \\h ')` | ❌ 域码变成 ` REF  REF OAB_验收结果 \\h  \\* MERGEFORMAT `（**域名重复**）→ `result="错误！未定义书签。"` |
| `Fields.Add(rng, 3, 'OAB_验收结果 \\h ')` | ✅ 域码 ` REF OAB_验收结果 \\h  \\* MERGEFORMAT `，`result="三、验收测试结果\t4"`（**交叉引用生效**） |
| `Fields.Add(rng, 37, 'OAB_验收结果 \\h ')` | ✅ 域码 ` PAGEREF OAB_验收结果 \\h  \\* MERGEFORMAT `，`Update()` 后 `result="2"`（**页码引用生效**） |

**避坑**：`Fields.Add` 的 Type 参数已经决定了域名，`Code` 里**不能再写域名**，否则会得到"错误！未定义书签。"且不报任何异常。MCP 无交叉引用工具。

## 6. 查找替换的高级用法

### 6.1 通配符替换：**生效**

`searchQuery="([0-9]{2,3})%"`, `useWildcards=true`, `replaceText="[\1 pct]"`

读回：命中 **4 个段落 + 4 个表格单元格**（`doc.Content` 之外还遍历了每张表的 Range），文本由 `100%` 变成 `[\x01 pct]`。

### 6.2 反向引用（backreference）：**不支持**

| replaceText | 读回 |
|---|---|
| `[\1 pct]`（Python `'\\1'`） | ❌ 写出的是 **控制字符 0x01**，而不是捕获组 `100`。`\1` 在参数传递途中已被解释成 0x01 |
| `A: [\1] B: [^&] C: [\2]`（`\\1`/`\\2` 字面量） | `\1`/`\2` 原样保留为字面文本（**无反向引用**）；`^&` 替换成了**整个匹配** `pct` → 得到 `B: [pct]` |
| 通配符模式下 `searchQuery="回填"`, `replaceText="<<^&>>"` | ✅ 得到 `<<<<回填>>>>`——说明 `^&` 在通配符模式下**确实代表整个匹配串**，可作反向引用的替代 |

### 6.3 格式替换：**不稳定（部分生效）**

`searchQuery="(达标)"`, `replaceText="达标(已核)"`, `replaceFormatting={bold:true, fontName:"微软雅黑", fontSizePt:11}`

- 段落 82 读回 `bold=-1(true), font=微软雅黑, size=11` ✅
- 表格 2 `[2,5]` 读回 `bold=-1(true)`，但 `font/size` **未读到**
- 另一次同参数对 `≥ 99%` 的替换：段落 91/97 的 `bold` 生效，`font=""`、`size=9999999`（未生效）

**判定：`bold` 相对稳定，`fontName`/`fontSizePt` 时好时坏**；返回值里没有任何格式读回信息，调用方无法判定。

### 6.4 只查找模式（不传 replaceText）：**matchCount 永远是 0（假信息）**

| searchQuery | matchCount | message |
|---|---|---|
| `验收` | `0` | `已为 "验收" 找到并应用格式化 (0 处)` |
| `达标` | `0` | 同上 |
| `Office` | `0` | 同上 |
| `MCP` | `0` | 同上 |

**根因（源码 `wps-addon/src/word.js:876`）**：`else if (hasFmt)` 分支在没传 `replaceFormatting` 时根本不会进入，`matchCount` 保持 0；message 还谎称"已应用格式化"。**同一文件 889 行的 `catch (e) {}` 把所有 Find 异常吞掉**，所以"0 处"既可能是没找到、也可能是 API 报错。

### 6.5 替换数量也是假的

源码 `wps-addon/src/word.js:875`：`matchCount++` 写在 `Execute(..., wdReplaceAll)` **之后**，收到的是 `wdReplaceAll` 的布尔返回值，每个 Range 只 +1。所以段落数+表数=2 时报的计数与实际替换处数无关。

## 7. 页面版式 / 水印 / 目录 / 排版工具

### 7.1 `wps_word_page_layout_and_watermark`

调用：`{headerText:"页眉：Office Agent Bridge 2.1 验收", footerText:"页脚：MCP Sweep Word 03", differentFirstPage:true}`
返回：`success=true`，message=`已成功更新 [测试文字文稿.docx] 页面版式与水印`

native 读回：

| 节 | orientation | 页眉 | 页脚 | DifferentFirstPage | FirstPageHeader |
|---|---|---|---|---|---|
| 1 | 0 纵向 | `页眉：Office Agent Bridge 2.1 验收` | `页脚：MCP Sweep Word 03` | `-1`(true) | **空** |
| 2 | 1 横向 | `第四部分 · 附录（横向页）`（**未被工具改动**） | `附录 A · 高级能力实测原始记录`（**未改**） | `-1`(true) | 空 |

**判定**：
- ✅ 第 1 节页眉页脚生效
- ⚠️ **只作用于第 1 节**，第 2 节完全没动，但 message 说"已成功更新页面版式"，没有任何"仅首节"的提示 → 多节文档下**静默漏改**
- ⚠️ `differentFirstPage:true` 把 `DifferentFirstPageHeaderFooter` 打开（读回 `-1`），但**首页页眉内容为空**，即"首页不同"= 首页什么都不显示，工具没有为首页写入任何内容

水印调用：`{watermarkText:"内部机密 严禁外传"}` → `success=true`，读回 `D.Shapes.Count` 2 → **3**，新增 `WordArt 8/15`；**但只加在第 1 节**（`Sections(1).Range.ShapeRange.Count=3`，`Sections(2)...=0`）→ ✅ 水印生效 / ⚠️ 同样只覆盖第 1 节。

### 7.2 `wps_word_insert_table_of_contents`

调用：`{upperHeadingLevel:1, lowerHeadingLevel:3, insertLocation:"start", includePageNumbers:true}`
返回：`success=true`，message=`已成功在 [测试文字文稿.docx] 插入标准目录`

读回：`TablesOfContents.Count` 1 → **2**，新目录 `UpperHeadingLevel=1 / LowerHeadingLevel=3`，15 条条目。
`TablesOfContents.Item(2).Update()` 后条目页码由 `3` 刷新为 `4` → ✅ **目录生成 + 域刷新都真实生效**。

### 7.3 `wps_word_format_document`

调用：`{target:"range", paragraphRange:[126,127], preset:"custom", fontName:"仿宋_GB2312", fontSizePt:16, lineSpacingPt:28, firstLineIndentChars:2, bold:false}`
返回：`已成功格式化段落 P126 ~ P127`

native 逐段读回 P126/P127：
```
font=仿宋_GB2312  NameFarEast=仿宋_GB2312  size=16
LineSpacing=28  LineSpacingRule=4(固定值)  CharacterUnitFirstLineIndent=2
```
→ ✅ **全部参数逐项生效**，是本轮唯一"参数全中"的排版类工具。

### 7.4 `wps_word_save_document`

| 调用 | 返回 | 落盘验证 |
|---|---|---|
| 原地保存 `{documentName}` | `savedPath=/Users/jolin/Downloads/测试文字文稿.docx` | ✅ 文件 mtime 刷新、大小 35KB |
| 导出 PDF `{filePath:".../word3-final.pdf", format:"pdf"}` | `success=true`，`savedPath` 回显该路径，message=`已成功导出为 PDF` | ❌ **磁盘上文件不存在**（`ls` 报 No such file）→ **"success 但没生效"** |

`doc.ExportAsFixedFormat(path, 17)` 在 WPS for Mac 上不抛异常也不落盘；`doc.ComputeStatistics(2)` 同时读回 `pages=8`，说明分页计算本身可用。

### 7.5 `wps_word_capture_preview` —— 实现了但没有 MCP 定义

- 处理器确实存在：`src/bridge/gateway/word.ts:170` + 注册表 `src/bridge/gateway.ts:106`
- 但 HTTP/MCP 调用返回：`{"success":false,"error":"未知工具 wps_word_capture_preview"}`
- 源码注释已自认这是"实现了但未注册"台账项
- **后果**：Agent 无法主动生成 Word 视觉预览，只能自己拼 `ExportAsFixedFormat`（而它在本机不落盘，见 7.4）→ **Word 侧的视觉验收出现能力缺口**

## 8. 工具级判定总表

| 工具 / 能力 | 通道 | 判定 |
|---|---|---|
| 文档属性（Title/Author/Subject/Keywords/Category/Comments） | 仅 script | ✅ 生效（MCP 无工具） |
| 自定义文档属性 | 仅 script | ✅ 生效 |
| 新建段落样式 / 字符样式 | 仅 script | ✅ 生效（MCP 无工具） |
| 应用样式（段落 / Range / 标题 3） | 仅 script | ✅ 生效，OutlineLevel 同步 |
| 修改基样式（正文） | 仅 script | ✅ 生效 |
| 表格加行/加列 | 仅 script | ✅ 生效（`add_row` 亦已单独复测：7→8 行） |
| **合并单元格** | 仅 script / MCP 工具 | ✅ 两条通道都生效，`Uniform` 正确变 false |
| 单元格底纹 | 脚本 + MCP 工具 | ✅ 都生效（**RGB 顺序**，不是 BGR） |
| 边框 / 列宽 / 行高 | 仅 script | ✅ 生效（行高需先设 `HeightRule`） |
| 表内公式 `Cell.Formula` / `AutoSum` | 仅 script | ❌ **不支持**，静默返回 null |
| 表内 AutoNum 域编号 | 仅 script | ❌ 不支持（只留 `\u0015` 占位） |
| 分节符 / 横向页 / 分节独立页眉页脚 | 仅 script | ✅ 生效 |
| **`page_layout_and_watermark` 页眉页脚** | MCP 工具 | ⚠️ 只改第 1 节，其余节静默漏改 |
| 水印 | MCP 工具 | ✅ 生效，但同样只覆盖第 1 节 |
| **书签创建/读取** | 仅 script | ✅ 生效（MCP 无工具） |
| **`write_content`.`location:"bookmark"`** | MCP 工具 | ❌ **没插到书签处**，反而污染了另一个书签的范围 |
| **内容控件（ContentControls）** | 仅 script | ✅ API 可用，能建能读 Title/Tag（MCP 无工具） |
| **交叉引用 REF / PAGEREF** | 仅 script | ✅ 生效（code 不能重复写域名） |
| **通配符查找替换** | MCP 工具 | ✅ 生效（段落 + 表格都命中） |
| **反向引用 `\1`** | MCP 工具 | ❌ 不支持（变 0x01）；`^&` 可代表整个匹配 |
| **格式替换** | MCP 工具 | ⚠️ 不稳定：bold 稳，fontName/fontSizePt 时好时坏 |
| **只查找模式 matchCount** | MCP 工具 | ❌ **永远返回 0**，message 还谎称"已应用格式化" |
| 目录生成 + 域刷新 | MCP 工具 | ✅ 生效 |
| `wps_word_format_document`（range+custom） | MCP 工具 | ✅ 全部参数生效 |
| 原地保存 | MCP 工具 | ✅ 生效 |
| **导出 PDF** | MCP 工具 | ❌ **success 但磁盘无文件** |
| `wps_word_capture_preview` | 无定义 | ❌ 不可调用（"未知工具"） |
| 批注 1 条 / 修订 2 条 / 1 条内容控件 | — | 初稿原有批注与修订**未被我破坏**（读回仍是 1 / 2） |

## 9. 卡住 / 未竟项

1. **`Cell.Formula` 全签名试错**：`(f)`、`(f,0)`、`(f,null,0)`、`(f,"\\# 0.0")` 都返回 null 无效果；`(f,0,"0")`、`(f,true,0)` 抛 `Too many parameters.`。已用静态求和绕过，结论是宿主不支持。
2. **`doc.ExportAsFixedFormat` 不落盘**：脚本通道直接调用与工具通道都返回成功但文件不存在。**未找到可用的 Word→PDF 路径**，因此本轮的"视觉验证"只能依赖 native 属性读回 + 分页统计（`ComputeStatistics(2)=8`），**没有真实 PDF/截图**证据。
3. **书签在后续操作后被清除**：先建的 `OAB_验收结果` / `OAB_交付状态` 两个书签，在删除一个被误插入目录的段落后**双双消失**（`D.Bookmarks.Count` 回到 0，`refFields` 也被清空）；重建的 `OAB_复测锚点` 在两次 TOC `Update()` 后仍存活。**怀疑**（标注为猜测）是删除段落时 Range 重叠导致书签失效，未做进一步隔离验证。
4. ~~**`wps_word_write_content` 会吃掉小写字母 `a`**~~ → **【2026-09-22 复核后撤回：非缺陷，是探针自己造的】** 见 §10。**注意：本文件 §10 当年的结论是错的**，保留原文仅为留痕"验证方法出错"这一类问题。

## 10. 【已撤回】写入路径吞掉小写字母 `a`——实为验证探针的正则坑

> **本节的原始结论已被推翻（2026-09-22 复核）。** 下面先写"当年怎么错的"，再写"复核结论与原始证据"，最后留下真正可复用的教训。**不要**再按本节的"规避"改写入路径。

### 10.1 当年（错误）的结论

现象被描述为：`wps_word_write_content` 返回 `success=true`，但写入文本里**所有小写字母 `a` 消失**；并用下表"隔离"到 `Paragraphs.Add(targetRange)` 路径：

| 写法（当年的表） | 当年写的读回 | **复核后的真相** |
|---|---|---|
| `doc.Paragraphs.Add()` 无参 + `p.Range.Text = S` | ✅ `S` 完整保留 | ❌ 该行原始输出其实是**空串**，不是 `S` |
| `doc.Paragraphs.Add(targetRange)` + `p.Range.Text = S` | ❌ 变 `"X  b bc bnn A A 啊阿"` | 写入**完好**；读回串被探针自己删了 `a` |
| `doc.Range(endPos,endPos).InsertAfter(S)` | ❌ 同样吞 `a` | 同上，写入完好 |
| `doc.Paragraphs.Add(targetRange)` + `InsertAfter(S)` | ❌ 同样吞 `a` | 该行原始输出是**空串** |
| transport 回显 | ✅ 完整 | ✅ 这条是对的（传输没问题） |

教训之一：**上表把"空串"读成了"完整保留"，把"探针删的"读成了"宿主吞的"**——一张表同时犯了两个方向的反向误读。

### 10.2 复核结论：原现象不成立

`wps_word_write_content` 写入的内容**一直是对的**，本仓库代码里也没有任何删 `a` 的逻辑。假象来自本文件当年的验证探针：

```js
// 当年的 verify_tool.js（逐字引用）
const t = String(D.Paragraphs.Item(i).Range.Text).replace(/[\r\a]/g, "");  // ← 这里删掉了所有 a
if (t === want) found.want.push(i);
if (t === stripped) found.stripped.push(i);     // stripped = want.replace(/a/g, "")
```

**原因**：`/[\r\a]/` 在 **JavaScript** 里 `\a` 是**恒等转义**＝字母 `a`（`\a` = BEL `0x07` 是 C/PCRE/.NET 的语义，JS 里 BEL 必须写 `\x07`）→ 探针先把读回文本里所有小写 `a` 删掉，再与原始串比较，必然得出"少了 `a`"。大写 `A`、中文完好也正因如此。

**原始输出（逐字引用，来自本子代理会话的 tool/result）**：

```
SENT to tool: '确认字符: a ab abc banana A Aa 啊阿'
tool success: True
{ "sentByPython": "确认字符: a ab abc banana A Aa 啊阿",
  "exactMatch": [], "strippedMatch": [148],
  "strippedForm": "确认字符:  b bc bnn A A 啊阿", "total": 149 }
```

`strippedMatch:[148]`（共 **149** 段）恰好证明内容**正确写到了文末**（`location:"end"` 语义正确）。

**第二轮隔离的原始输出（同样是逐字引用）**——四个写法**每个**都以同一行 `.replace(/[\r\a]/g, "")` 收尾：

```
SENT: 'Xa a ab abc banana A Aa 啊阿'
Paragraphs.Add(targetRange)+Range.Text     -> X  b bc bnn A A 啊阿   （= S 去掉全部 a → 写入正确）
Paragraphs.Add()+Range.Text                -> （空串）
Paragraphs.Add(targetRange)+InsertAfter    -> （空串）
ContentEnd Range.InsertAfter               -> X  b bc bnn A A 啊阿   （写入正确）
```

### 10.3 排除清单（方法可复核）

| 假设 | 方法 | 结果 |
|---|---|---|
| 本仓库有删 `a` 的代码 | 把 `src/**`、`wps-addon/src/**`、`office-addon/src/**`、`scripts/**` 里 1161 条正则字面量 + `new RegExp` 逐条当清洗器实测 | **0 条**能删掉每个 `a` |
| 历史版本曾用 `\a` | `git log --all -S` 搜 `[\r\n\a]` / `\a` | 从未出现 |
| 构建改坏转义 | 比对 `addon-core.js` 与 `src/**` 的 `[\r\n\x07]` | 一致 |
| 传输吞字符 | transport 回显 + 已部署副本与仓库产物同 sha256 | 排除 |
| 宿主 API 吞 `a` | 上面原始输出：读回串只差探针删掉的部分 | **不成立** |

### 10.4 真正可复用的教训（比原结论有价值）

**触发条件**：写脚本读回 Word 文本、想清掉段落标记/单元格标记 → **错误做法**：`.replace(/[\r\n\a]/g, "")` → **已证实原因**：JS 正则里 `\a` 是恒等转义＝字母 `a`（BEL 要写 `\x07`），会把读回文本里所有小写 `a` 删掉，制造"宿主吞字符"的假报警（本文件当年据此浪费了一整轮排查，还差点去改正常的写入路径） → **正确做法**：一律写 `[\r\n\x07]`；看到"宿主行为诡异"时**先逐行回读探针的清洗/比较/断言逻辑** → **证据**：本节 10.2 的原始探针与原始输出。

**现行处置**：`wordWriteContent` 的读回校验已从"只比长度"升级为**逐字比对**（等长改写也拦）+ 报错带出写入/读回原文 + 每行返回 `readBackMatches`；**写入路径未改**。详见 [agent-tests/word-fixes.md](../../agent-tests/word-fixes.md) §1 与 [issues.md](../../issues.md) 的"撤稿记录：ISS-67"。

## 11. MCP / skill 说明的不足

| 位置 | 问题 |
|---|---|
| 工具清单 | **完全没有**：文档属性、样式、书签、内容控件、交叉引用、分节、页眉页脚分节控制、表格列宽行高、导出预览 —— 全部只能靠 `wps_execute_script` |
| `wps_word_read_document` scope 枚举 | schema 写了 `paragraphs`/`tables` 两个枚举值，但**参数描述里只解释了 outline/full/selection**，`paragraphs`/`tables` 没有说明差异 |
| `wps_word_manage_table` | 无 `columnWidth`/`rowHeight`/`border`/`formula` 参数；`mergeRange` 描述只有"合并单元格范围"，**没说 `startRow` 等是否 1-based**（实测是 1-based） |
| `wps_word_manage_table`.`format_cell` | `backgroundColor` 只写"十六进制底色，例如 '#F1F5F9'"，**没说明是 RGB 顺序还是 BGR**（实测 RGB，与 Excel 侧习惯相反，易写反） |
| `wps_word_find_and_replace` | 描述"支持通配符、全字匹配、区分大小写"——**没提反向引用不支持**（`\1` 会静默变 0x01）；**没提只查找时 matchCount 恒为 0** |
| `wps_word_write_content` | `location:"bookmark"` 描述没有说明"书签不存在时的回退行为"，实测**书签存在也没插对位置** |
| `wps_word_page_layout_and_watermark` | 描述"支持设置页眉页脚（支持奇偶页不同、首页不同）"，**没说明只作用于第 1 节** |
| `wps_word_save_document` | `format:"pdf"` 描述没有平台差异说明；本机 macOS 不落盘 |
| skill `office-agent-bridge-word-batch-edit/SKILL.md` | 第 8 行只提 `wps_word_create_document`/`wps_word_save_document`；**没有"Word 视觉验收在本机不可用"的说明**，第 23 行的"宿主无预览时明确视觉验证缺口"过于笼统，实际 `wps_word_capture_preview` 根本调不到 |
| skill `references/native-scripting.md` | 只给了 Excel Shapes 与 PPT 示例，**Word 表格/样式/分节/书签一个例子都没有**，本轮全靠 `wps_inspect_api` 反射试出来 |
| `wps_inspect_api` | 反射 `doc.Tables.Item(2).Rows.Item(1).Cells.Item(1)` 能列出 `Formula`/`AutoSum`/`Merge` 等方法，**但列不出签名**，导致必须逐个参数个数试错才知道 `Formula` 最多 2 参 |

## 12. 成品现状（native 最终读回）

```
文档: /Users/jolin/Downloads/测试文字文稿.docx   已保存=true
段落 173 / 分页 8 页 / 2 节（节1 纵向 595×842，节2 横向 842×595）
样式: 自定义在用样式 2 个 —— "OAB 实测正文"(段落)、"OAB 强调标记"(字符)
表格: 表1 5×5(uniform)  表2 7×5(非 uniform，含 2 处合并)
  表2 标题行(合并 1×5) = "2.2 模块交付状态 · 高级表格实测汇总"，底纹 0xE8EEF7
  表2 第 3 行已合并 [3,1]-[3,2]；第 2 行第 2 格 bold+14pt+底纹 #F1F5F9（工具写入）
  表2 列宽 100/150/80/46/70，行2 Height=28pt(Exactly)、行1 Height=32pt(AtLeast)
  表2 合计行 = "合计 296.2"（静态结果，公式域不支持）
书签: OAB_复测锚点
内容控件: 1（Title="OAB 实测控件", Tag="oab-test"）
目录: 2 个 TOC 域（均已 Update，页码已刷新）
交叉引用: REF + PAGEREF 域各 1 条，指向 OAB_复测锚点
文档属性: Title/Subject/Keywords/Category/Comments 全部写入；自定义属性「验收批次」
页眉页脚: 节1="页眉：Office Agent Bridge 2.1 验收"/"页脚：MCP Sweep Word 03"
          节2="第四部分 · 附录（横向页）"/"附录 A · 高级能力实测原始记录"
水印: 3 个 WordArt（2 原有 + 1 新增"内部机密 严禁外传"，仅节1）
原有批注 1 / 修订 2 保持未动
```
