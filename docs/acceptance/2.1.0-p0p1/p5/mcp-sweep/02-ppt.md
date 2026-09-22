# PPT 复查报告（p5 / mcp-sweep / 02-ppt）

状态：进行中（第一优先实验已完成，第二优先能力探测进行中）
执行者：演示文稿复查员（子代理）
环境：WPS 演示（macOS）+ Office Agent Bridge 2.1.0，服务端口 19890，凭据 `~/.wps-bridge/token`
调用方式：HTTP `POST http://127.0.0.1:19890/api/v1/tool/call`，body 为 `{"name": <工具名>, "arguments": {...}}`（注意**不是** `{"tool":..., "args":...}`，实测会报 `Cannot read properties of undefined (reading 'startsWith')`）
脚本与原始数据：`.scratch/ppt3/`
被测对象：新建空白演示文稿 `ppt3-blank.pptx`（`.scratch/ppt3/`，`Presentations.Add()` 后 `SaveAs`，0 页起）

---

## 0. 环境先决发现：WPS 里跑的加载项是**旧构建**

这一条影响后面所有读数的解释，先单独列。

| 位置 | 文件 | 大小 | mtime | 含 `fitGeneratedPptShapes` |
|---|---|---|---|---|
| 仓库 | `wps-addon/addon-core.js` | 203647 B | 09-22 15:53 | 是 |
| WPS 已部署（演示，`_` 与 `_2.1.0` 两份，sha `c30e548afaac`） | `addon-core.js` | 196693 B | 09-22 11:51 | 是 |
| WPS 备份 | `addon-core.js.backup-1789967115643-147a9c`（sha `3b34c6a5d7c7`） | 186997 B | 09-21 13:05 | 否 |
| WPS 备份 | `addon-core.js.backup-1789962205405-7ffb57`（sha `707aa2d8f580`） | 178161 B | 09-21 11:43 | 否 |

判据（**行为指纹**，非猜测）：本次 `wps_ppt_generate_deck` 的返回体字段集合为
`{success, presentationName, createdSlidesCount, createdSlideIndices, themeColor, message}`，
**没有** `pageWidth/pageHeight/unit`、**没有** `layoutWarnings`、**没有** `visualVerificationRequired`。
文本比对确认：这个字段集合与 09-21 13:05 备份（`3b34c6a5d7c7`）的 `pptGenerateDeck` return 块**逐字一致**；
而磁盘上 09-22 11:51 的部署版与仓库版都返回 `...page, layoutWarnings, visualVerificationRequired`。

→ 结论：**WPS 演示进程内存里运行的是 09-21 13:05（无缩放）的旧构建；磁盘上 11:51 的新构建未生效。**
证据：`.scratch/ppt3/01-generate-deck.json`、`.scratch/ppt3/04-build-scan.json`。

---

## 1. 干净实验：`wps_ppt_generate_deck` 几何复现

### 1.1 实验输入（唯一一次生成调用）

```json
{"presentationName":"ppt3-blank.pptx","themeColor":"#0F4C81",
 "slides":[{"layout":"title","title":"实验标题页","subtitle":"几何复现 · 干净实验"},
           {"layout":"content","title":"内容页A（有正文）","bulletPoints":["要点一","要点二","要点三"]},
           {"layout":"content","title":"内容页B（故意不传正文）"},
           {"layout":"end","title":"谢谢观看"}]}
```

### 1.2 工具原始返回（逐字，`.scratch/ppt3/01-generate-deck.json`）

```json
{"success":true,"data":{"success":true,"presentationName":"ppt3-blank.pptx","createdSlidesCount":4,
"createdSlideIndices":[1,2,3,4],"themeColor":"#0F4C81",
"message":"已成功基于大纲批量生成 4 页专业商业演示胶片！"}}
```

**没有任何警告字段。** 工具描述写"返回 layoutWarnings，需逐页预览验证"，实际返回里连键都没有。

### 1.3 原始读数（`app.Presentations.Item("ppt3-blank.pptx")` 直接读宿主对象，`.scratch/ppt3/02-raw-shapes.json`）

页面：`PageSetup.SlideWidth=960`，`SlideHeight=540`（单位 pt），`Slides.Count=4`，新页 `Layout=12`（空白版式）。

| 页 | 形状序 | Id | 名称 | type | Left | Top | Width | Height | FontSize | Text |
|---|---|---|---|---|---|---|---|---|---|---|
| 1 | 1 | 2 | 矩形 1 | 1 | 0 | 0 | **720** | **405** | — | — |
| 1 | 2 | 3 | 文本框 2 | 17 | 60 | 140 | 600 | 50.8 | 36 | 实验标题页 |
| 1 | 3 | 4 | 文本框 3 | 17 | 60 | 230 | 600 | 29 | 18 | 几何复现 · 干净实验 |
| 2 | 1 | 2 | 文本框 1 | 17 | 50 | 30 | 620 | 36.25 | 24 | 内容页A（有正文） |
| 2 | 2 | 3 | 直接连接符 2 | 9 | 50 | 75 | 70 | 0.05 | — | — |
| 2 | 3 | 4 | 文本框 3 | 17 | 60 | 100 | 600 | 65.35 | 16 | 要点一\n要点二\n要点三 |
| 3 | 1 | 2 | 文本框 1 | 17 | 50 | 30 | 620 | 36.25 | 24 | 内容页B（故意不传正文） |
| 3 | 2 | 3 | 直接连接符 2 | 9 | 50 | 75 | 70 | 0.05 | — | — |
| 4 | 1 | 2 | 矩形 1 | 1 | 0 | 0 | **720** | **405** | — | — |
| 4 | 2 | 3 | 文本框 2 | 17 | 60 | 160 | 600 | 55.65 | 40 | 谢谢观看 |

### 1.4 回答三个问题

**问题 1：页面尺寸？标题页背景多大？是否超框？**
- 页面 960×540 pt（16:9）。
- 第 1 页背景（`矩形 1`）= **720×405 pt，位于 (0,0)**。
- **不超框，反而小 25%**：宽 720/960 = 75%，高 405/540 = 75%。页面右侧空白 240 pt、底部空白 135 pt，右下角 1/4 面积裸露。
- 第 4 页（`end`）背景同样是 720×405，问题一样。

**问题 2：坐标是"最终位置"还是"还要乘系数"？系数多少？**
- **本次实测系数 = 1.0，即没有任何换算**：背景、标题框、正文框全是设计基准坐标原样落到页面上（设计基准就是 `ppt.js` 里写死的 `pageWidth=720 / pageHeight=405` 与 `AddTextbox(1, 60, 140, 600, 80)` 等）。
- 正确系数应为 `sx = 960/720 = 1.3333`、`sy = 540/405 = 1.3333`（代码 `pptScaleForPage` 就是这么写的），本次**未被应用**。
- 静态核对（**预测，非宿主实测**）：`wps-addon/src/ppt-layout.js` 的 `pptScaleForPage` 在 960×540 上返回 sx=sy=1.3333，若磁盘上那份新代码真的在跑，背景应读作 960×540、标题框应读作 Left=80/Top=186.67/Width=800。仓库版代码逻辑本身能覆盖满页；**当前运行版不具备该逻辑**。

**问题 3：`content` 布局不传正文时生成了什么？有警告吗？**
- 第 3 页只生成了 **2 个形状**：标题文本框（50,30,620,36.25，字号 24）+ 标题下装饰线（直接连接符，50,75,70,0.05）。
- **正文形状完全没有被创建**——不是空文本框，是连形状都没有（对比第 2 页的第 3 个形状）。
- **工具没有任何警告**：返回体里根本没有 `layoutWarnings` 键（见 1.2）；`message` 仍是"已成功基于大纲批量生成 4 页专业商业演示胶片！"。
- **根因（源码级确认，非推断）**：`wps-addon/src/ppt.js` 第 346–355 行，`else` 分支里只有
  `if (Array.isArray(spec.bulletPoints) && spec.bulletPoints.length > 0) { ...建正文框... }`，
  **没有 else**，缺正文即静默跳过；同一函数里的 `warnings` 只由 `fitGeneratedPptShapes` 在"文字可能溢出"时推送（第 40 行），**从不校验正文是否存在**。新旧两版构建在这段代码上完全相同（已逐字节比对），因此这是**稳定可复现的设计缺陷**，与环境新旧无关。

### 1.5 与上一轮"问题 A（1280×720 超框 33%）"的对照

上一轮的原始产物还在：`测试演示文稿.pptx`（960×540，10 页，`/Users/jolin/Downloads/`）。逐形状扫描（`.scratch/ppt3/03-existing-deck-scan.json`）结果：

| 页 | 形状 | Left | Top | Width | Height | 说明 |
|---|---|---|---|---|---|---|
| 1 | 矩形(type 1) | 0 | **-180** | **1280** | **720** | 唯一超框形状，正是问题 A |
| 1 | 文本框 | 80 | 186.67 | 800 | 67.73 | = 设计 (60,140,600) × 4/3 |
| 1 | 文本框 | 80 | 306.67 | 800 | 38.67 | = 设计 (60,230,600) × 4/3 |
| 10 | 矩形(type 1) | 0 | 0 | **960** | **540** | = 设计 (0,0,720,405) × 4/3，**满页正确** |
| 10 | 文本框 | 80 | 213.33 | 800 | 74.2 | = 设计 (60,160,600) × 4/3 |
| 2–9 | 全部 | 66.67 / 133.33 / 826.67 / 346.67 … | | | | 全部 = 设计值 × 4/3 |

要点：
1. 同一份文件里，**第 10 页（end 布局）的背景是 960×540，满页正确**；只有第 1 页的背景是 1280×720。同一函数、同一页面尺寸，不可能一次调用产出两种背景 → **第 1 页那个 1280×720 不是 `generate_deck` 的设计输出**（设计输出是 720×405，乘 4/3 是 960×540，都不等于 1280×720）。
2. 1280×720 恰好等于 960×540 pt 在 96 dpi 下的**像素尺寸**（1 pt = 4/3 px），也等于设计基准 720×405 × (4/3)²。
3. 第 1 页的**文字**框位置（80 / 186.67 / 306.67）证明这一页确实被 ×4/3 缩放处理过，唯独背景是 1280×720@(0,-180)，Top 也不是 0。
4. 本次干净实验**无法复现** 1280×720：新建空白文稿上生成的背景是 720×405。要产生 1280×720，需要背景基准 960×540 再乘 4/3，或 720×405 乘 (4/3)²，两条路径当前生成代码都不存在。
5. 结论与"猜测"的划分：**已证实**——当前 `generate_deck` 稳定产出 720×405 背景（欠 25%），且 1280×720 那个形状就在上一轮文件第 1 页上原文存在、Top=-180；**未证实（标注为待定）**——它是被谁、在哪一步改成 1280×720 的。最省事的判定办法是查该 page 形状的创建来源（上一轮调用记录/审计），本轮无法从当前宿主回溯。

### 1.6 会话内构建切换（决定性对照）

第一次生成（14:5x）跑的是旧构建，几分钟后再调用 `wps_ppt_generate_deck` 时返回体**已经变成新格式**（多出 `pageWidth/pageHeight/unit/layoutWarnings/visualVerificationRequired`）——WPS 演示里的加载项在这两次调用之间被重载成了磁盘上 11:51 的新构建。于是同一份文稿内出现了天然对照：

| 页 | 生成时构建 | 背景矩形 | 标题框 (L,T,W) | 标题字号 | 返回体 |
|---|---|---|---|---|---|
| 1–4 | 旧（无缩放） | **720×405** | 60, 140, 600 | 36 | 无 layoutWarnings |
| 6–7 | 新（有缩放） | **960×540（满页）** | 80, 186.67, 800 | 48 | `layoutWarnings: []` |
| 8–9 | 新 | — | 66.67, 40, 826.67 | 32 | `layoutWarnings: []`（无正文，仍无警告） |
| 20–24 | 新（单次 5 页批量） | 960×540 | 按 4/3 换算 | 48 | `layoutWarnings: []`，逐页几何检查 0 越界 |

原始读数见 `.scratch/ppt3/07-after-reload-shapes.json`、`.scratch/ppt3/22-batch-multipage.json`。

**结论**：几何问题（欠缩放/超框）的根因是**运行中的加载项构建版本**，代码侧 11:51 的新构建在 960×540 上缩放系数 sx=sy=4/3，背景正好满页、无越界。上一轮那次 1280×720 读数无法在任一构建里由 `generate_deck` 复现，属该页上的既存/外来形状（见 1.5）。

---

## 2. 剩余能力探测（全部在 `ppt3-blank.pptx` 上实做）

### 2.1 能力总览

| 能力 | 有无专用工具 | 实测结果 |
|---|---|---|
| 母版 / 版式 | 无专用工具 | 脚本可读写；`manage_slides add` 的 layoutIndex 可套用内置版式并生成占位符 |
| 表格 | `wps_ppt_manage_table` | 可用，但有 3 个缺陷（见 2.3） |
| 形状增删改 | `wps_ppt_manage_shapes_and_media` | 可用；`swap_shapes`/`align_shapes` 语义与描述不符 |
| 图片 | **无**（无 add_picture/OLE 工具） | 未验证到可用路径 |
| 原生图表 | `wps_ppt_insert_native_chart` | **100% 失败**（宿主 API 返回 null） |
| 卡片排版 | `wps_ppt_add_business_cards` | 可用，缩放正确，超栏有明确报错 |
| 切换 / 动画 | **无专用工具** | 脚本可读写 `SlideShowTransition.EntryEffect`（写 3844 读回 3844）、`TimeLine.MainSequence.Count` 可读 |
| 多页批量排版 | `wps_ppt_generate_deck` | 单次 5 页 0.24 s，逐页几何 0 越界 |
| 预览截图 | `wps_ppt_capture_slide_preview` | **100% 失败**；用脚本 `slide.Export` 绕过并完成真实视觉核查 |
| 演讲者备注 | 随 `generate_deck` 的 `notes` | 可用，`NotesPage.Shapes.Placeholders.Item(2)` 读回一致 |
| 保存 | **无**（Word/Excel 各有 save 工具） | `wps_ppt_save_presentation` → `未知工具`；需脚本 `pres.Save()` |

### 2.2 母版 / 版式（脚本读 + 工具写）

`ppt3-blank.pptx` 原始读数：

```
Designs=1  SlideMaster.Name="WPS"  SlideMaster.Shapes=5
  母版占位符: 标题占位符1 / 文本占位符2 / 日期占位符3 / 页脚占位符4 / 灯片编号占位符5 (Type=14)
CustomLayouts.Count=11
  1 标题幻灯片(5 shapes)  2 标题和内容(5)  3 节标题(5)  4 两栏内容(6)  5 比较(8)
  6 仅标题(4)  7 空白(3)  8 图片与标题(6)  9 竖排标题与文本(5)  10 内容(4)  11 末尾幻灯片(5)
```

`wps_ppt_manage_slides action=add layoutIndex=N` 的真实语义（实测）：

| layoutIndex | 生成页的 `Layout` | `CustomLayout.Name` | 页内形状 |
|---|---|---|---|
| 1 | 1 | 标题幻灯片 | 2 个占位符（标题 1 / 副标题 2） |
| 2 | 2 | 标题和文本 | 2 个占位符（标题 1 / 文本占位符 2） |
| 7 | 7 | 标题和图示或组织结构图 | 2 个占位符（标题 1 / 智能图形占位符 2） |
| 12 | 12 | 空白 | 0 个形状 |

→ `layoutIndex` 走的是 **ppLayout 旧枚举**，与上面 `CustomLayouts` 的 1..11 序号**不同源**、名字也对不上（序号 2 是"标题和文本"而不是 `CustomLayouts.Item(2)` 的"标题和内容"，序号 7 是"标题和图示或组织结构图"而不是 "空白"）。传 12 不报错（`CustomLayouts.Count` 只有 11）。工具描述写的"版式序号(默认 12 空白版式)"没有说明这一点。

占位符（`typeCode 14 / typeName "placeholder"`）可用 `wps_ppt_manage_shapes_and_media action=update_shape shapeId=<占位符 Id> text=...` 写入：实测第 17 页"文本占位符 2"写入"写进占位符的正文"、fontSize=28，读回 `hasText:true`。**这是当前唯一可用的"按母版版式排版"路径**（原生命令行占位符填充技巧）。脚本侧可直接读写 `slide.CustomLayout`、`pres.SlideMaster`、`pres.Designs`。

### 2.3 表格（`wps_ppt_manage_table`）

| 动作 | 参数 | 结果 |
|---|---|---|
| create_table | rows=3,columns=3,left=80,top=120,width=800,height=300 | 成功；实测 L=80,T=120,W=800,**H=122**（`height` 被宿主行高覆盖） |
| create_table（带 data） | data 含数字 `1000` | 失败：`arguments.data[1][2]: 类型不正确` |
| create_table（带 data，全字符串） | 3×3 中文数据 | 成功，`read_table` 逐格读回一致 |
| read_table | tableIndex=1 | 成功，返回二维数组 |
| set_cell_text | row=2,column=2 | 成功，读回 `改过的单元格` |
| set_cell_text | row=4（表只有 3 行） | 失败：`Cannot read properties of null (reading 'Shape')` |
| set_table_data | — | **失败：`未知的表格操作: set_table_data`**（工具 enum 里有、宿主 switch 里没有） |
| style_table | headerFillColor/字号/zebra/borderColor | 成功，读回样式生效 |

附带：`create_table` 的返回 message 恒为"已成功创建 **8 行 5 列**表格…"，与实际的 3 行 3 列不符（请求什么尺寸都一样）。`get_slide_shapes` 把表格形状报成 `hasText:true, text:"企业"`（首个单元格文本外泄到父形状）。

### 2.4 形状增删改（`wps_ppt_manage_shapes_and_media`）

- `add_textbox`：坐标原样生效；请求 height=50 实测 36.25（宿主按文字自动收缩）。
- `add_shape`：`rectangle|rounded_rectangle|oval` 均成功，坐标精确；`shapeType:"star"` 被 catalog 拒绝：`arguments.shapeType: 不在允许值中`（清晰）。
- `update_shape`：几何+文本+字号+填充色全部生效，返回读回值；`shapeId=999` → `未在第 11 页找到形状: 999`（清晰）。
- `swap_shapes`：**只交换 Top（垂直换位），不动 Left**。返回体里直接写着 `{oldTop,newTop,height}`。两个横向并排的形状调用后位置实际无变化，与描述"智能互换两个形状位置"不符。
- `set_z_order`：`bring_to_front` 等返回成功。
- `align_shapes`：**不是页面对齐**，而是"把列表里其余形状对齐到**第一个形状**的对应边/中心"（源码 `refVal` 只从 `i===0` 取）。更麻烦的是 `shapeIds` 解析先按 **shape Id** 再按 **页内索引** 兜底（`findShapeOnSlide`），传 `[1,2,3]`（真实 Id 是 2/3/4）会把形状 1 落到索引兜底，实测把 3 个矩形的对齐做成了"只动中间那个"。
- `delete_shape`：成功；`shapeId=999` → `未找到形状: 999`。
- 读回：`rounded_rectangle`/`oval` 的 typeCode 都是 1、typeName 都是 `shape`，**无法区分矩形/圆角/椭圆**。

### 2.5 原生图表（`wps_ppt_insert_native_chart` / `generate_deck` 的 chart 布局）

宿主级探针（`wps_execute_script` 直接调用宿主 API）：

```json
{"hasAddChart":"function","hasAddChart2":"function",
 "addChart_ret":"null","addChart_isNull":true,"shapesAfter":0,
 "addChart2_ret":"null","addChart2_isNull":true,"shapesAfter2":0,
 "ole":"null"}
```

→ `Shapes.AddChart` / `AddChart2` / `AddOLEObject` 三个方法**都存在但返回 null 且不创建任何形状**（页内形状数保持 0）。包装层随后读 `chartShape.Chart`，抛出：

```
图表已创建，但数据配置未完成：Cannot read properties of null (reading 'Chart')
```

这个文案是**误导**：图表根本没有创建（读回该页 `shapeCount: 0`）。`generate_deck` 用 `layout:"chart"` 时同样失败，而且**页面已经被插入、标题和装饰线已经画上**，整次调用报错返回——部分写入 + 硬错误，没有回滚。

### 2.6 切换 / 动画

无任何 MCP 工具覆盖。脚本层可用（**只读探测 + 一次写入验证**）：

```
SlideShowTransition.EntryEffect : number（第 1 页 = 0）
写入 3844 后读回 3844  →  切换效果可设
SlideShowTransition.Speed = 3, AdvanceOnTime = 0
TimeLine.MainSequence.Count = 0  →  动画序列对象可达
```

### 2.7 多页批量排版 + 真实视觉核查

一次调用生成 5 页（title / content / cards_3 / cards_4 / end），0.24 s，返回 `layoutWarnings: []`；脚本逐页扫描：页 20=3 形状、21=3、22=17、23=22、24=2，**越界形状 0 个**。

预览工具不可用，改用脚本 `slide.Export(path,"PNG",1280,720)` 导出并**真的看了图**：结束页紫色满幅（无白边、无欠缩放），内容页标题+要点+装饰线正常。→ 新版构建的几何在视觉上成立；旧构建那批页面（第 1–4 页背景 720×405）按读数就是右下角 25% 裸露。

### 2.8 其他实测缺陷

- **`set_background` 会串页（受控复现）**：新建空白文稿 → 加 2 张空白页 → 只给第 1 页设 `#FF0000` → 导出两张图取中心像素：第 1 页 `(255,0,0)`、**第 2 页也是 `(255,0,0)`**。返回文案却是"已将第 1 页背景设为 #FF0000"。原文稿里设过背景的页（第 5 页）让全部页都变 `RGB(15,23,42)`。
- **没有 PPT 保存工具**：`wps_ppt_save_presentation` → `未知工具 wps_ppt_save_presentation`；只能用 `wps_execute_script` 调 `pres.Save()`（实测 `Saved=-1`，25 页落盘）。
- **`wps_ppt_add_chart` 别名不可调用**：`未知工具 wps_ppt_add_chart`（源码注释里承认这是"实现了但未注册"的死分支）。
- `duplicate / move / delete` 幻灯片、`add_business_cards` 独立调用（缩放正确，超栏报 `卡片数量超过分栏数，请拆分为多页，避免内容被截断`）、`notes` 备注写入、`read_presentation` 大纲读取全部正常。

---

## 3. 卡点与报错原文

| # | 场景 | 原文 |
|---|---|---|
| 1 | HTTP 调用格式用 `{tool,args}` | `{"success":false,"error":"Cannot read properties of undefined (reading 'startsWith')"}`（正确 body 是 `{name,arguments}`） |
| 2 | `Presentations.Add()` + `SaveAs` | `WPS execute_script 超时，结果未知；先读取状态，不要自动重放写入`（21 s 后超时，但读回确认文稿已建好：`ppt3-blank.pptx, 960×540, 0 页`） |
| 3 | `wps_ppt_insert_native_chart`（默认/显式几何、column/pie） | `图表已创建，但数据配置未完成：Cannot read properties of null (reading 'Chart')` |
| 4 | `generate_deck` `layout:"chart"` | 同上；且页面已插入、标题已画，调用整体失败 |
| 5 | `create_table` 的 data 含数字 | `arguments.data[1][2]: 类型不正确` |
| 6 | `set_table_data` | `未知的表格操作: set_table_data` |
| 7 | `set_cell_text` 行号越界 | `Cannot read properties of null (reading 'Shape')` |
| 8 | `wps_ppt_capture_slide_preview`（3 个目标 × 6 次） | `PPT 未生成预览`；对照实验 `slideIndex=9999` 会正常回传 `Cannot read properties of null (reading 'Export')`，证明该工具的错误是会透传的，这句兜底文案意味着 addon 报了"成功"而 bridge 在 `~/.wps-bridge/previews/` 找不到文件（20 ms 轮询 3 s 也没出现文件） |
| 9 | 想自己指定预览输出路径 | `arguments.outputPath: 未知参数`（schema 未开放，无法让 AI 自己选落盘位置） |
| 10 | 保存 / 别名 | `未知工具 wps_ppt_save_presentation`、`未知工具 wps_ppt_add_chart` |

预览卡的绕行做法（已用于本次视觉核查）：`wps_execute_script` + `component:"ppt"` + `pres.Slides.Item(n).Export("<绝对路径>.png","PNG",1280,720)`，实测在第 3 页导出成功（93 ms、84 KB），输出到 `.scratch/ppt3/` 与 `~/.wps-bridge/previews/` 均可。

---

## 4. MCP 与 skill 说明的不足（具体到工具与参数）

1. **`wps_ppt_generate_deck`**：描述写"返回 layoutWarnings，需逐页预览验证"，但①运行中的旧构建**完全不返回该字段**；②即便新版返回，`warnings` 只由"文字可能溢出"触发，**不覆盖正文缺失**。应补一句："`content`/`cards` 不传 `bulletPoints`/`cards` 时只生成标题，不生成正文框"。另外它一次调用会插入 N 页，`chart` 布局失败时**已插入的页不会回滚**，描述里没有任何"部分写入"提示。
2. **`wps_ppt_manage_slides.layoutIndex`**：描述为"版式序号(默认 12 空白版式)"。实测它是 **ppLayout 枚举**（1/2/7/12…），与 `CustomLayouts` 的 1..11 序号不同源，且 12 越界时不报错。描述应写明映射，或直接接受版式名。
3. **`wps_ppt_manage_table`**：enum 里的 `set_table_data` **宿主未实现**；`data` 只接受字符串（描述写"批量注入二维数据"，但混入数字会被 schema 拒），错误信息不指出期望类型；`height` 参数实际被行高覆盖；成功 message 固定写"8 行 5 列"。
4. **`wps_ppt_manage_shapes_and_media`**：`swap_shapes` 只换 Top、`align_shapes` 是"对齐到列表第一个形状"而非页面对齐；`shapeIds`/`shapeId` 的"先按 Id、失败再按索引"二义性完全没写，最容易被 AI 误用（本次就误用了）。也没有 `add_picture`，图片能力在描述里缺位。
5. **`wps_ppt_insert_native_chart`**：macOS WPS 上**完全不可用**（宿主 `AddChart` 返回 null）。描述与 `bridge_get_capabilities` 只笼统写"复杂图表与母版需实机验证，不保证高保真复刻"，不足以让 AI 事先避开；建议直接标注平台可用性。相关地，`generate_deck` 的 `chart` 布局描述里也没标"macOS 不可用"。
6. **`wps_ppt_capture_slide_preview`**：描述承诺"导出高保真图片供多模态核查"，实际 100% 报 `PPT 未生成预览`；且该错误**不携带 addon 的原始返回**（handler 里 `res?.error || 'PPT 未生成预览'` 在 `success:true` 但文件不存在时把信息丢光），AI 无法自查。schema 也不允许传 `outputPath`（实测 `未知参数`），连绕行到指定路径都不行。PPT skill 要求"每个修改页完成后检查一次预览"，却既没提该工具当前不可用，也没给脚本兜底（`slide.Export`）。
7. **保存链路缺口**：Word/Excel 各有 save 工具，**PPT 没有**；而 `office-agent-bridge` skill 的通用流程要求"按需保存…调用对应保存接口"，在 PPT 上会落空（本次只能用 `wps_execute_script` + `pres.Save()`）。
8. **`set_background` 的副作用没写**：描述是"背景颜色设置"（按页），实测会**改到全部页**，且返回文案仍按单页汇报。这是最容易让 AI 静默毁掉整份配色的一处，建议在描述里显式警告或修实现。

---

## 5. 结论一句话

- 问题 A：**当前构建不复现**；几何欠缩放的根因是 `wps_ppt_manage_*`/`generate_deck` 跑在**旧的、无 `fitGeneratedPptShapes` 的加载项构建**上（旧构建产出 720×405 背景，占页面 75%）；换成磁盘上的新构建后背景正好 960×540。上一轮读到的 1280×720 是 `测试演示文稿.pptx` 第 1 页上的既存形状（Top=-180），不是 `generate_deck` 的设计输出。
- 问题 B：**稳定复现且已定位到源码**（`wps-addon/src/ppt.js:346-355` 缺 else 分支，warnings 不校验正文），新旧构建都一样，`layoutWarnings` 恒为空数组。
- 能力面：表格/形状/卡片/版式占位符/多页批量可用；原生图表、预览截图、保存、切换动画四块在 MCP 层缺失或不可用。

