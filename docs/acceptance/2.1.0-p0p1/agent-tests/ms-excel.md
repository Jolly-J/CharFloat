# Microsoft Excel（Office.js 加载项）侧能力核实

- 任务：共享任务 `task-3`
- 执行时间：本机 2026-09-22
- 目标宿主：**Microsoft Excel for Mac 16.113（CFBundleVersion 16.113.26091433）**，工作簿 `工作簿1.xlsx`
- 通道：桥接 `http://127.0.0.1:19890/api/v1/tool/call`，Office.js 任务窗格走 `host: "microsoft"`
- 写区：仅本文件；**未修改任何产品源码**
- 判据口径：**「能用」= 有真实读数**。工具报"未实现"只证明桥接层拒了它，不证明宿主不行。

---

## 0. 一句话结论

> **13 项登记为"Microsoft 未实现"的能力里，只有 2 项在 Office.js 交付面上真的做不到**
> （`format_text_segment`、`export_sheet_pdf`）；**其余 11 项宿主 API 都在**——其中 **7 项 Office.js 加载项里已经有实现**，只是被桥接层的缺口声明在调用宿主之前挡住了。
>
> 细分：**A（确实做不到）= 2**、**B（只是没接上）= 11**、整体判为 **C（待确认）= 0**（另有 5 项局部待确认，见 §6）。
> B 的 11 项里，7 项加载项已有代码（§4.2），4 项要新写（§4.3）；
> 其中 2 项带尾巴：`set_sheet_view` 的 **视图 zoom** 在 Office.js 没有对应 API（只能做网格线/行列标题），
> `create_workbook` 的桌面端行为未实测。这两处都在 §4 逐个写明，不折进"已解决"。
>
> ⚠️ **证据层级提醒**：13 项里**没有任何一项**拿到过"送到宿主并成功"的真机读数——桥接层在调用宿主前就把它们全拒了。
> 形状域（自检 + 工具路径）有真机读数；其余 12 项是"工具路径读数 + 加载项源码 + 需求集对照 + 版本表"推出来的，**不属实机通过**。

另外，任务背景里三条"已知结论"有两条是**误判**，根因都不是宿主：

| 已知结论 | 本次实测结论 |
|---|---|
| 直线报「当前对象不允许此操作」 | ❌ **误判**。`addLine` 真实创建成功（读回到 `type=Line` 的形状，包围盒与入参一致）；报错来自工具**写后读回**时对 Line 形状 load `fill`。 |
| `addSvg` / `getActiveShape` 本机没有这两个 API | ⚠️ **只对一半**。`addSvg` 确实是 **ExcelApi BETA（预览）专属**，GA 面没有；`getActiveShape` **存在**，但它是 `Workbook.getActiveShape()`（ExcelApi **1.19**），加载项错调成了 `sheet.shapes.getActiveShape()`。本体 16.113 ≥ 1.19 要求的 16.96。 |
| 分组经工具路径对已存在的表报「当前对象不允许此操作」，根因未定位 | ✅ **已定位**。`addGroup` **执行成功**（用 `ungroup_shapes` 反证组合确实存在）；报错来自 `loadShapeDetail()` 对 **Group 形状** load `fill`/`textFrame` 后整批 `sync()` 抛错。 |

---

## 1. 判定分级（本报告统一口径）

| 分级 | 含义 |
|---|---|
| **A · 宿主真的不支持** | Office.js 交付面（GA 需求集）没有该 API，且没有可用的替代 API → 加载项通道做不到，只能如实记录 |
| **B · 只是工具没接上** | 宿主 API 存在、本机 Excel 版本支持该需求集，但桥接缺口声明/加载项分支/参数映射没接齐 → **要补实现** |
| **C · 待确认** | 有证据指向某结论，但缺一次真机读数（或本机平台无法验证） |

> ⚠️ 贯穿全文的证据层级：**真机读数** > **源码可证** > **官方文档可证** > **推测**。
> 凡是"源码/文档可证但没有真机读数"的，一律不写成"已实测"。

---

## 2. 环境与取证前提（先证明"读的是哪一份代码"）

### 2.1 宿主与通道真实读数

```
bridge_get_capabilities → connection.components.msExcel
  { connected: true, version: "2.0.0", activeDocument: "工作簿1.xlsx",
    activeSheet: "OPPO销售驾驶舱", summary.selection.address: "OPPO销售驾驶舱!I13" }
office_get_status → { runningComponents: { excel: true }, openDocuments: { excel: ["工作簿1.xlsx"] } }
```

`components.msExcel` 就是 Office.js 任务窗格通道（`ws-server.ts` 里 `key === 'ms-excel'` 落到 `msExcel`）。
`components.word/excel/ppt` 是 **WPS** 加载项通道（Word 文稿/PPT 都在 WPS 里打开）——两者不要混。

### 2.2 任务窗格是否运行最新代码 —— 三层证据

| # | 证据 | 读数 | 说明 |
|---|---|---|---|
| 1 | **运行期能力探针**（最有力） | `excel_get_charts {"host":"microsoft"}` → `data.addon = { version:"2.0.0", capabilities:{ cap08Shapes:true, shapeMethods:[add_shape, list_shapes, update_shape, group_shapes, ungroup_shapes, set_shape_zorder, get_active_shape, export_shape_image, shape_self_test, insert_image, reload] } }` | `addon` 字段由**窗格进程内**的 `ADDON_CAPABILITIES` 常量产出（[state.js:7](../../../../office-addon/src/state.js)、[chart.js:55](../../../../office-addon/src/excel/chart.js)）。拿到它 = 窗格里跑的就是带 CAP-08 分支的那份代码。 |
| 2 | **服务端字节一致性** | `sha256(https://127.0.0.1:19891/office-addon/taskpane.js?_=…) = a73ad3a94e25b80d7fa96119e22b97a3e639a4e18ba3fb3d466f36d6206f518f`，与磁盘 [office-addon/public/taskpane.js](../../../../office-addon/public/taskpane.js) **逐字节相同**（149884 字节）；响应头 `Cache-Control: no-cache, no-store, must-revalidate`。 | 桥接托管的就是当前构建，且不会被缓存。 |
| 3 | **功能自检跑得通** | `excel_get_charts {"host":"microsoft","shapeSelfTest":true}` 返回 13 项逐项结果 | 自检函数 `handleShapeSelfTest` 只存在于新构建（[shape.js:486](../../../../office-addon/src/excel/shape.js)）。 |

> ### ⚠️ 对旧判据的更正（重要）
>
> 旧文档 [16-cap-ms-shapes.md](../p5/mcp-sweep/16-cap-ms-shapes.md) §2.1 用
> `components.msExcel.version === "2.0.0"` + `buildStatus === null` 判定"窗格跑的是旧代码"。
> **这条推理对 Office.js 通道不成立**，理由源码可证：
>
> - `version` 来自 [office-addon/src/state.js:3](../../../../office-addon/src/state.js) 的 **硬编码常量** `ADDON_VERSION = "2.0.0"`，历次改动从未更新它 → 它永远是 2.0.0，`versionStatus.staleSuspect: true` 属**误报**。
> - `buildFingerprint` / `buildStatus` **只在 WPS 分支被写入**（[ws-server.ts:464-486](../../../../src/bridge/ws-server.ts)）：`if (isMs) {…}` 分支里没有 `buildFingerprint` 字段 → Office.js 通道**永远**读不到，与新旧无关。
>
> **正确判据是证据 #1 的 `addon.capabilities` 探针**（本次读数 `cap08Shapes: true` → 新代码）。

### 2.3 "不碰用户数据"的执行与验证

- 全部探测在自建临时表 `_ms_probe` / `_ms_probe2` / `_cap08_shapetest` 上进行；
- 收尾已删除三张表，读回工作表清单恢复为探测前的 4 张：
  `['Sheet1', '2026Q3销售分析', 'OPPO大中华区Q3销售总报表', 'OPPO销售驾驶舱']`，活动表回到 `OPPO销售驾驶舱`；
- 用户表完整性读回：`excel_get_charts {"host":"microsoft","sheetName":"OPPO销售驾驶舱"}` → `count=3`（`Chart 1 ColumnClustered` / `Chart 2 Doughnut` / `Chart 3 BarClustered`），与探测前一致；
- 全程**未调用** `save_workbook`，未对任何用户表写入。

---

## 3. 形状域：三条已知结论的根因定位

### 3.1 任务窗格内置自检真实读数（13 项）

`excel_get_charts {"host":"microsoft","shapeSelfTest":true,"selfTestKeep":true}`

```
total=13  passed=10  failed=['addLine（连接符/直线）','addSvg','getActiveShape']

OK   addGeometricShape(rectangle) + 填充/线条/旋转读回   54ms
FAIL addLine（连接符/直线）                              16ms  当前对象不允许此操作。
FAIL addSvg                                              2ms  sheet.shapes.addSvg is not a function. (…) 'sheet.shapes.addSvg' is undefined
OK   addTextBox + 字体读回                              104ms
OK   shapes 列表读回（几何属性 + zOrder）                 23ms
OK   shapes 填充/线条读回（逐个 load）                     1ms
OK   setZOrder(bringToFront/sendToBack)                   4ms
FAIL getActiveShape                                       0ms  sheet.shapes.getActiveShape is not a function. (…) 'sheet.shapes.getActiveShape' is undefined
OK   addGroup + 组合成员读回                               7ms
OK   ShapeGroup.ungroup()                                 3ms
OK   getAsImage(png)                                     19ms  → 5076 字节，iVBORw0KGgo…
OK   scaleWidth/scaleHeight                               2ms  100×100 → 150×150
OK   delete + 读回确认                                     2ms
```

**同一个返回体的 `readback` 里，`cap08_line` 是存在的**：

```json
"list": [ {"name":"cap08_rect","type":"GeometricShape", …},
          {"name":"cap08_line","type":"Line","left":200,"top":20,"width":160,"height":90,"rotation":0,"z":"1"},
          {"name":"cap08_text","type":"GeometricShape", …} ]
```

入参是 `addLine(200, 20, 360, 110)` —— 位置尺寸完全对得上。**自检的 FAIL 标记错了对象：`addLine` 成功了，失败的是紧接着的 `loadShapeDetail()`。**

### 3.2 工具路径实验（自建临时表，逐步隔离）

| 步骤 | 调用 | 真实读数 |
|---|---|---|
| A1 | `excel_add_shape {kind:"geometric", shapeType:"roundRectangle", name:"p_rect", …}` | `success:true`，读回 `fill.foregroundColor=#2F6FEB`、`verified:true` |
| A2 | `excel_list_shapes`（表内**只有 1 个几何形状**） | ❌ `属性"name"不可用。读取属性的值之前，请先对包含对象调用 load 方法，再对关联的请求上下文调用 "context.sync()"。` |
| A3 | `excel_add_shape {kind:"line", name:"p_line", x1:240,y1:40,x2:420,y2:140}` | ❌ `当前对象不允许此操作。` |
| A4 | `excel_list_shapes`（几何 + 直线） | ❌ `当前对象不允许此操作。` |
| A5 | `excel_update_shape {name:"p_line", action:"delete"}` | ✅ `success:true, deleted:"p_line", remaining:["p_rect"], verified:true` ← **直线确实存在过** |
| A6 | `excel_list_shapes`（直线已删） | 回到 A2 的 `属性"name"不可用` |
| C1 | 再加一个 `rectangle`（`p_oval`） | ✅ `success:true` |
| C2 | `excel_group_shapes {shapeNames:["p_rect","p_oval"], groupName:"p_group"}` | ❌ `当前对象不允许此操作。` |
| D5 | `excel_ungroup_shapes {name:"p_group2"}` | ✅ `releasedChildren:["p_rect","p_oval"], childrenStillPresent:["p_rect","p_oval"], verified:true` ← **组合确实存在过** |
| D4 | `excel_update_shape {name:"p_group2", left:110}`（只改位置） | ❌ `当前对象不允许此操作。` ← 该分支只做 `loadShapeDetail` + sync |
| D1 | `excel_export_shape_image {shapeName:"p_rect"}` | ✅ `success:true`，`imageBase64: iVBORw0KGgo…`（真实 PNG） |
| B3 | `excel_set_shape_zorder {name:"p_rect", zOrder:"bringToFront"}` | ✅ `verified:true`，`zOrderBottomToTop` 读回 |

### 3.3 根因（源码 + 真机读数互相印证）

**根因 1 —— `loadShapeDetail()` 对不支持 `fill` 的形状 load `fill`，整批 `sync()` 抛错。**
[shape.js:72-80](../../../../office-addon/src/excel/shape.js) 无条件把 `fill/*`、`lineFormat/*`、`textFrame/*` 都排进加载队列：

```js
function loadShapeDetail(shape) {
  shape.load("name,id,type,left,top,width,height,rotation,zOrderPosition,visible");
  try { shape.load("fill/type,fill/foregroundColor,fill/transparency"); } catch (e) {}
  try { shape.load("lineFormat/color,lineFormat/weight,lineFormat/visible"); } catch (e) {}
  try { shape.load("textFrame/textRange/text"); } catch (e) {}
  …
}
```

`try/catch` 只能拦住**同步**异常；真正的失败发生在 `await context.sync()`，拦不住，整批失败。
- **Line 形状没有 `fill`** → `add_shape{kind:"line"}` 的写后读回（A3）、`list_shapes`（A4）双双抛「当前对象不允许此操作」。
- **Group 形状没有 `fill`/`textFrame`** → `group_shapes`（C2）、`update_shape` on group（D4）抛同一错误。

**根因 2 —— `handleListShapes` 对同一个 `ShapeCollection` 连续 `load()`，后者覆盖前者。**
[shape.js:231-236](../../../../office-addon/src/excel/shape.js)：

```js
shapes.load("items/name,items/id,items/type,…");                                   // 第 1 次
try { shapes.load("items/fill/type,items/fill/foregroundColor,…"); } catch (e) {}   // 第 2 次
try { shapes.load("items/lineFormat/color,items/lineFormat/weight"); } catch (e) {} // 第 3 次
try { shapes.load("items/textFrame/textRange/text"); } catch (e) {}                 // 第 4 次
await context.sync();
const items = shapes.items.map(shapeToJson);   // ← shapeToJson 第一句读 shape.name
```

真机读数完全吻合这个解释：
- 表内 **0 个形状** → `list_shapes` 成功（`Sheet1` → `count:0`）；
- 表内 **≥1 个形状** → `属性"name"不可用`（`name` 在第 1 次 load 里，被后续 load 覆盖掉了）；
- 表内**有 Line** → 更早就在 sync 阶段撞上「当前对象不允许此操作」。

**根因 3 —— 逻辑与实际 enum 键不符的别名表（12/47 条坏）。**
[shape.js:13-31](../../../../office-addon/src/excel/shape.js) 的 `SHAPE_TYPE_ALIASES` 映射目标里有 12 个不是 `Excel.GeometricShapeType` 的键：

| 别名 | 映射到（不存在） | 真实 enum 键 |
|---|---|---|
| `rounded_rectangle` / `roundedrectangle` / `round_rect` | `roundedRectangle` | **`roundRectangle`** |
| `oval` / `ellipse` / `circle` | `oval` | **`ellipse`** |
| `flow_chart_data` / `flowchart_data` | `flowChartData` | `flowChartInputOutput`（推定） |
| `cross` / `plus` | `cross` | **`plus`** |
| `text_box` / `textbox` | `textBox` | 无（文本框是独立方法 `addTextBox`） |

真机读数：`shapeType:"rounded_rectangle"` → `不认识的几何形状类型 "rounded_rectangle"`；`"ellipse"` → 同样报错。
**注意 `rounded_rectangle` 正是 [16-cap-ms-shapes.md](../p5/mcp-sweep/16-cap-ms-shapes.md) §4.2 交接示例里写的取值** —— 那条示例照抄会直接失败。

### 3.4 形状域结论

| 能力 | 判定 | 依据 |
|---|---|---|
| 矩形/几何形状、填充、线条、旋转、文字 | **B（能用）** | 真机读数（自检 OK + A1/C1/E4 工具路径 OK，`verified:true`） |
| 文本框 | **B（能用）** | 自检 OK + 工具路径 OK |
| 层级 `setZOrder` | **B（能用）** | 自检 OK + `excel_set_shape_zorder` `verified:true` |
| 形状导图 `getAsImage` | **B（能用）** | 自检 OK（5076 字节 PNG）+ `excel_export_shape_image` 返回真实 base64 |
| 缩放 `scaleWidth/Height` | **B（能用）** | 自检 OK：100×100 → 150×150 |
| **直线 / 连接符** | **B（宿主完全能做；工具读回代码有缺陷）** | A5 反证直线真实存在；A3/A4 的报错来自 `loadShapeDetail` |
| **分组 / 解组** | **B（宿主完全能做；工具读回代码有缺陷）** | D5 反证组合真实存在；C2/D4 的报错来自 `loadShapeDetail` 读 Group 的 `fill` |
| **`getActiveShape`** | **B（宿主 API 存在，加载项调错对象）** | `Workbook.getActiveShape()` = ExcelApi **1.19**；加载项调 `sheet.shapes.getActiveShape()` → `is not a function`。本机 16.113 ≥ 1.19 要求的 16.96 |
| **`addSvg`** | **A（宿主 GA 面确实没有）** | 官方 API 页标注 `API set: ExcelApi BETA (PREVIEW ONLY)`；`@types/office-js@1.0.533` 中**完全没有** `addSvg` 符号。另：工具 schema 的 `kind` 枚举里也没有 `svg`（见 §5） |

---

## 4. 13 项"Microsoft 未实现"逐项判定

### 4.0 统一前提：工具路径的真实读数

**13 项全部在调用宿主之前被桥接层拒绝**，且（补齐合法 `action` 取值后）返回体完全一致：

```json
{"success":false,"error":"excel_<name> 在 Microsoft 宿主上未实现（加载项缺少该 RPC 分支），已拒绝且未执行。"}
```

拒绝点在 [catalog.ts:193](../../../../src/bridge/catalog.ts)：`if (isUnimplementedOn(host, bareMethod)) throw …`，
判定表是 [host-methods.ts:50-52](../../../../src/bridge/contracts/host-methods.ts) 的 `HOST_IMPLEMENTATION_GAPS.microsoft`。

> ⚠️ **这句话本身是错的。** "加载项缺少该 RPC 分支"对其中 7 项不成立 —— 加载项**有**这些分支与实现，见 §4.2。

**所以工具路径的读数只能证明"桥接层拒了"，不能证明宿主不行。** 下面的"宿主 API"栏改由以下三条独立证据支撑：
1. `@types/office-js@1.0.533` 的 `index.d.ts`（每个成员都标了 `[Api set: ExcelApi x.y]`）；
2. 本机 Excel **16.113** + 微软官方要求集↔版本对照表；
3. 真机需求集下界探针：`excel_manage_cell_comments`（用 `Worksheet.comments` = **ExcelApi 1.10**）在 `host=microsoft` 上**真实成功**（返回批注 id）→ **本机至少支持到 1.10**；形状工具成功 → 1.9 已覆盖；`addSvg` 不可用 → 不含 BETA。

官方对照（<https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets>）：
Office on Mac 支持 **ExcelApi 1.19 需 16.96**、**1.20 需 16.100**、**1.21 需 16.110.1**。
**本机 16.113 → 支持到 ExcelApi 1.21。** 即下面所有 ≤1.21 的需求集本机都满足。

### 4.1 判定总表

| # | 能力（工具） | 工具路径读数 | 宿主 API（Office.js） | 需求集 | 加载项已有分支 | 判定 |
|---|---|---|---|---|---|---|
| 1 | `excel_get_style_token` | 拒绝，未执行 | `Workbook.styles`(1.7)；`Range.format.font` / `.fill`(1.1) | 1.7 | ✗ | **B · 只是没接上** |
| 2 | `excel_format_text_segment` | 拒绝，未执行 | **GA 面不存在**（无单元格内局部格式 API） | — | ✗ | **A · 宿主真的不支持** |
| 3 | `excel_configure_print_layout` | 拒绝，未执行 | `Worksheet.pageLayout`(1.9) + `PageBreakCollection`(1.9) | 1.9 | ✗ | **B · 只是没接上** |
| 4 | `excel_export_sheet_pdf` | 拒绝，未执行 | **GA 面不存在**（Excel 没有 PDF 导出 API） | — | ✗ | **A · macOS 上做不到**（Windows 见 §4.3-4） |
| 5 | `excel_set_sheet_view` | 拒绝，未执行 | `showGridlines`/`showHeadings`(1.8) ✔；**视图 zoom 无 API** | 1.8 | ✗ | **B（部分）· 网格线/标题没接上；zoom 做不到** |
| 6 | `excel_copy_range` | 拒绝，未执行 | `Range.copyFrom` | **1.9** | ✅ `handleCopyRange` | **B · 只是没接上** |
| 7 | `excel_manage_hyperlink` | 拒绝，未执行 | `Range.hyperlink` | **1.7** | ✅ `handleSetHyperlink` | **B · 只是没接上** |
| 8 | `excel_manage_named_range` | 拒绝，未执行 | `Workbook.names` / `Worksheet.names` | **1.1** | ✅ `handleListNamedItems` / `handleUpdateNamedItem` | **B · 只是没接上** |
| 9 | `excel_manage_document_properties` | 拒绝，未执行 | `Workbook.properties` | **1.7** | ✅ `handleGetProperties` / `handleUpdateProperties` | **B · 只是没接上** |
| 10 | `excel_manage_table` | 拒绝，未执行 | `Worksheet.tables` / `TableCollection.add` | **1.1** | ✅ `handleCreateTable` / `handleUpdateTable` | **B · 只是没接上** |
| 11 | `excel_manage_pictures` | 拒绝，未执行 | `ShapeCollection.addImage`（无独立 PictureCollection） | **1.9** | ✅ `handleInsertImage` | **B · 只是没接上** |
| 12 | `excel_export_chart_image` | 拒绝，未执行 | `Chart.getImage(w,h,fittingMode)` | **1.2** | ✅ `handleExportChartImage` | **B · 只是没接上** |
| 13 | `excel_create_workbook` | 拒绝，未执行 | `Excel.createWorkbook(base64?)` | **1.8** | ✗ | **B（待确认桌面行为）** |

**统计：A = 2，B = 11（其中 7 项加载项代码已经写好），C = 0。**
其中 3 项的 `action` 取值还被 schema 拦了一层（原调用取值得不到"未实现"提示）：
`excel_configure_print_layout` 允许 `apply/read`（不是 `get`）、`excel_manage_document_properties` 允许 `read/apply/delete`（不是 `get`）、`excel_manage_table` 允许 `list/apply/delete`（不是 `create`）。

### 4.2 七个"加载项里已经写好、只是被挡住"的（源码可证）

| 工具 | 加载项 RPC 分支 | 实现位置 | 用到的宿主 API | 现状 |
|---|---|---|---|---|
| `excel_copy_range` | **`copy_range`（同名）** | [range.js:314](../../../../office-addon/src/excel/range.js) | `destRange.copyFrom(sourceRange, copyType)` | `HOST_IMPLEMENTATION_GAPS.microsoft` 把它声明成未实现 → **声明有误** |
| `excel_export_chart_image` | **`export_chart_image`（同名）** | [chart.js:506](../../../../office-addon/src/excel/chart.js) | `chart.getImage(w,h)` | 同上，**声明有误** |
| `excel_manage_hyperlink` | `set_hyperlink` | [range.js:330](../../../../office-addon/src/excel/range.js) | `range.hyperlink = {address, textToDisplay, screenTip}` | 分支名与桥接方法名不同，未映射 |
| `excel_manage_named_range` | `list_named_items` / `update_named_item` | [workbook.js:95](../../../../office-addon/src/excel/workbook.js) | `workbook.names.load/add/getItem().delete()` | 同上 |
| `excel_manage_document_properties` | `get_document_properties` / `update_document_properties` | [workbook.js:118](../../../../office-addon/src/excel/workbook.js) | `workbook.properties.load/set` | 同上 |
| `excel_manage_table` | `create_table` / `update_table` | [table.js:4](../../../../office-addon/src/excel/table.js) | `sheet.tables.add(range, hasHeaders)` / `table.style` / `table.resize()` | 同上 |
| `excel_manage_pictures` | `insert_image` | [shape.js:207](../../../../office-addon/src/excel/shape.js) | `sheet.shapes.addImage(base64Image)` | 分支在，但**连工具都没注册**（`excel_insert_image` → `未知工具`；`insert_image` 不在 `EXCEL_METHODS` 里） |

> **这 7 项没有真机读数**——因为桥接层在调用宿主前就拒了，没有任何合法入口能把请求送到窗格。
> 判定依据是"加载项已有实现 + 所用 API 属于本机已支持的需求集"，**属源码可证，不属实机通过**。
> 要变成"已实测"，必须先修 `HOST_IMPLEMENTATION_GAPS.microsoft`。

### 4.3 需要新写实现的 4 项

**4.3-1 `get_style_token` → B（宿主完全能做）**
WPS 侧实现（[wps-addon/src/excel.js:7-35](../../../../wps-addon/src/excel.js)）只做了三件事：读 `sampleAddress` 的字体名/字号/加粗/字色/底色，再读一次 `A1` 的字体名。Office.js 等价物：

```js
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getItem(sheetName);
  const r = sheet.getRange(sampleAddress);          // 默认 "A3"
  r.format.font.load("name,size,bold,color");
  r.format.fill.load("color");
  const title = sheet.getRange("A1");
  title.format.font.load("name");
  context.workbook.styles.load("items/name");        // Workbook.styles = ExcelApi 1.7
  await context.sync();
  return { fontName: r.format.font.name, sampleFontSize: r.format.font.size,
           sampleBold: r.format.font.bold, fontColor: r.format.font.color,
           headerBackgroundColor: r.format.fill.color, titleFontName: title.format.font.name };
});
```

**4.3-2 `configure_print_layout` → B（宿主完全能做）**
`excel_configure_print_layout` 的 schema 字段（`printArea` / `printTitleRows` / `printTitleColumns` / `orientation` / `paperSize` / `zoom` / `fitToPagesWide` / `fitToPagesTall` / `centerHorizontally` / `centerVertically` / `printGridlines` / 六项 margin / 六个页眉页脚 / 分页符）在 `Excel.PageLayout`（**ExcelApi 1.9**）里**逐个都有**。
以下签名逐个取自 `@types/office-js@1.0.533`（行号即该文件行号）：

```js
const pl = sheet.pageLayout;                         // Worksheet.pageLayout = ExcelApi 1.9
pl.setPrintArea("A1:H40");                            // setPrintArea(Range|RangeAreas|string)
pl.setPrintTitleRows("$1:$2");                        // setPrintTitleRows(Range|string)
pl.setPrintTitleColumns("$A:$A");                     // setPrintTitleColumns(Range|string)
pl.orientation = Excel.PageOrientation.landscape;     // {portrait|landscape} = 1.7
pl.paperSize  = Excel.PaperType.a4;                   // {letter…a3,a4,a4Small…} = 1.7
pl.zoom = { scale: 85 };                              // PageLayoutZoomOptions：打印缩放（≠视图缩放）
//   fitToPagesWide/Tall 走同一个对象：pl.zoom = { horizontalFitToPages: 1, verticalFitToPages: 2 }
//   ⚠️ horizontalFitToPages/verticalFitToPages 不是 PageLayout 的直接属性，只在 zoom 里
pl.centerHorizontally = true; pl.centerVertically = true;
pl.printGridlines = true;  pl.printHeadings = false;
pl.setPrintMargins(Excel.PrintMarginUnit.points,      // {points|inches|centimeters} = 1.9
  { left:36, right:36, top:36, bottom:36, header:18, footer:18 });   // PageLayoutMarginOptions
pl.headersFooters.defaultForAllPages.set({            // headersFooters: HeaderFooterGroup → HeaderFooter
  leftHeader: "…", centerHeader: "&A", rightHeader: "&D",
  leftFooter: "…", centerFooter: "第 &P 页 / 共 &N 页", rightFooter: "…"
});
sheet.horizontalPageBreaks.add(sheet.getRange("A30"));   // add(Range|string)：PageBreakCollection = 1.9
sheet.verticalPageBreaks.add(sheet.getRange("H1"));
// 读回：pl.getPrintArea(): RangeAreas、pl.getPrintTitleRows()/Columns(): Range、
//      sheet.horizontalPageBreaks.getCount()
```

**4.3-3 `set_sheet_view` → B（部分）**
```js
sheet.showGridlines = true;   // ExcelApi 1.8
sheet.showHeadings  = true;   // ExcelApi 1.8
```
⚠️ **`zoom`（视图缩放）在 Office.js GA 面没有对应属性**——`Excel.Worksheet` 的完整成员表里没有 zoom；唯一的 `zoom` 是 `PageLayout.zoom`（**打印**缩放）。要还原视图缩放只能走原生通道（macOS AppleScript `window.zoom` / Windows COM `ActiveWindow.Zoom`）。建议把这些字段**如实回报为 unsupportedFields**，不要静默丢弃。

**4.3-4 `export_sheet_pdf` → A（macOS 上做不到）**
- Office.js：`@types/office-js@1.0.533` 全文无 `exportAsFixedFormat`；唯一的 `pdf` 枚举是 **`WordApiDesktop 1.1` 的 `Word.ImageFormat.pdf`**，与 Excel 无关 → **Excel 加载项通道没有 PDF 导出。**
- macOS 原生（`office_execute_script`，本机走 JXA）：导出 Excel 的 AppleScript 脚本字典（`sdef "/Applications/Microsoft Excel.app"`，838 KB，370 条 command）里**没有 `export as fixed format`**；能落 PDF 的只有
  - `save as`（字典描述即 *"Saves changes into a different file."*，会给工作簿**改关联文件名/路径**，不是导出）配合 `XlFileFormat: PDF file format`（0x02bc0039）；
  - `print out`（有 `print to file` 参数，但需要 PostScript 驱动，不是 PDF 通路）。
  → **在 macOS 上"安全地导 PDF 且不改动原工作簿"做不到。** 不建议把它实现成 `save as`：那会改掉用户工作簿的路径关联。
- Windows COM：`resources/office/excel.ps1` 的 28 个方法里没有导出；但 Windows COM 面本身有 `Workbook.ExportAsFixedFormat`。**本机是 macOS，无法验证 → 标 C（待确认）**，建议在 Windows 实机验收时再定。

### 4.4 `format_text_segment` 为什么是 A

`Excel.Range` 的 116 个成员里**没有任何字符级/局部文本格式入口**（无 `textRange`、无 `getTextRange`、无 `characters`）；`Excel.TextRange` 只挂在形状/图表上。
即"只给单元格里某一段文字加粗变红"在 **Office.js GA 面做不到**（`Range.format` 一律作用于整格）。

替代路径（都不是 Office.js 结构化 API）：
- 形状文本整体格式：`shape.textFrame.textRange.font.bold`（已在自检里真机验证 OK）；
- 整格格式：`range.format.font.*`（现有 `excel_format_cells` 已覆盖）；
- 单元格内富文本只能靠 **macOS AppleScript**（脚本字典里有 `character` 类、`change case` 等命令）或 **Windows COM `Range.Characters`**。

### 4.5 关于 `create_workbook`（B，附待确认）

`Excel.createWorkbook(base64?)` = **ExcelApi 1.8**，属性表与本机 16.113（支持到 1.21）都满足，加载项**没有** RPC 分支 → 属"没接上"。
但**桌面端**（尤其 Mac）`createWorkbook` 的实际行为（是否弹"信任加载项"、是否真的新建窗口）**本次未实测** → 该子项标 **C**。
另注：`excel_create_workbook` 的 schema 有 `savePath` / `sheetName`，而 `Excel.createWorkbook` 只接受 base64 初值、不接受落盘路径——落盘仍需 `Workbook.save` 或用户在宿主里另存，这层语义差要在实现时如实回报。

---

## 5. 顺带发现的工具契约缺陷（都会让"能用"变成"报错"）

| # | 缺陷 | 位置 | 真机读数 |
|---|---|---|---|
| D1 | `add_shape` 的 `kind` 枚举**没有 `svg`**，加载项里 `kind === "svg"` 的分支不可达 | schema `excel_add_shape.kind` vs [shape.js:156](../../../../office-addon/src/excel/shape.js) | `arguments.kind: 不在允许值中（允许值：geometric / textBox / textbox / line / connector / wordart）` |
| D2 | 12/47 条形状别名映射到不存在的 enum 键（`rounded_rectangle`→`roundedRectangle`、`ellipse`→`oval`、`plus`→`cross`…） | [shape.js:13-31](../../../../office-addon/src/excel/shape.js) | `不认识的几何形状类型 "rounded_rectangle"` / `"ellipse"` |
| D3 | `kind` 枚举里有 `wordart`（WPS 独有），Office.js 侧会静默落进 `geometric` 分支建一个矩形 | schema vs [shape.js:145-198](../../../../office-addon/src/excel/shape.js) | 源码可证，未单独跑 |
| D4 | 参数名不一致：`export_shape_image` 用 `shapeName`（不是 `name`）、`ungroup_shapes` 用 `shapeName`/`name`（不是 `groupName`）、`add_shape` 用 `name`（不是 `shapeName`） | schema | `arguments.name: 未知参数；本工具允许的参数：shapeName, shapeId, format, scale, sheetName, host` |
| D5 | `HOST_IMPLEMENTATION_GAPS.microsoft` 的报错文案"（加载项缺少该 RPC 分支）"对 7 项**与事实不符** | [host-methods.ts:50-52](../../../../src/bridge/contracts/host-methods.ts) + [catalog.ts:197](../../../../src/bridge/catalog.ts) | §4.2 |
| D6 | `insert_image` 分支在加载项里存在，但不在 `EXCEL_METHODS`，也没有对应工具 → 死分支 | [host-methods.ts:19-29](../../../../src/bridge/contracts/host-methods.ts) | `{"name":"excel_insert_image"}` → `未知工具 excel_insert_image` |

---

## 6. 待确认清单（明确不写成确定结论）

| # | 待确认项 | 为什么不能定 | 下一步 |
|---|---|---|---|
| C1 | §4.2 那 7 项在真机上是否真的能跑通 | 桥接层在调用宿主前拒绝，**没有合法入口**能把请求送进窗格 → 只有源码可证 | 修 `HOST_IMPLEMENTATION_GAPS.microsoft` + 对齐分支名后重测 |
| C2 | `Workbook.getActiveShape()`（1.19）在本机是否真能取到值 | 无法在窗格里执行任意 Office.js（见 §7 限制）；版本表说支持（16.113 ≥ 16.96），但 API 在无选中形状时会抛 `ItemNotFound` | 补实现时加 `Office.context.requirements.isSetSupported('ExcelApi','1.19')` 守卫 + `getActiveShapeOrNullObject` 实测 |
| C3 | `Excel.createWorkbook` 在桌面版 Mac 的真实行为 | 未实测（会弹窗/新建窗口，属可见副作用） | Windows/macOS 实机各跑一次 |
| C4 | Windows COM 的 `ExportAsFixedFormat` 能否补上 `export_sheet_pdf` | 本机是 macOS，`resources/office/*.ps1` 无法执行 | Windows 实机验收 |
| C5 | 用户表 `OPPO销售驾驶舱` 上 `list_shapes` 报「当前对象不允许此操作」的确切形状类型 | 该表有 3 张图表；**未在用户表上做破坏性隔离实验**（铁律） | 在临时表上单独建一张图表后复现（本次未做，因临时表已清理） |

---

## 7. 本次探测的方法学限制（如实说明）

1. **本机没有"把任意 Office.js 送进任务窗格"的合法通道。**
   `tool-registry.ts:17` 只把 `excel_*` + `host=microsoft` 的调用路由到 Office.js；`wps_execute_script` 无论传什么 `component` 都路由到 **WPS** 加载项（所以会出现 `Excel is not defined`）。`EXCEL_METHODS` 里没有 `run_script`，因此也没有 `excel_execute_script` 工具。
   **窗格里唯一的"直试宿主 API"入口就是 `handleShapeSelfTest`**（它自己在注释 [shape.js:478-482](../../../../office-addon/src/excel/shape.js) 里说明了这一点）。
   → 形状域拿到了逐项真机读数；其余 12 项只能靠"工具路径读数 + 加载项源码 + 需求集对照 + 版本表"。
2. **`office_execute_script` 不是 Office.js**：本机走 macOS JXA（`Application('Microsoft Excel')`），是 Excel 的 AppleScript 模型，**不能用来回答"Office.js 有没有这个 API"**，只能用来回答"Excel 这个应用程序有没有这个能力"。
   ⚠️ 注意 JXA 的 `typeof wb.任意不存在的属性` **一律返回 `"function"`**（已做对照实验：`typeof wb.zzzNotAThing === "function"`），所以**不要用 typeof 探测 JXA 成员是否存在**——本报告改用 `sdef`（脚本字典）作为该通道的证据来源。
3. `sdef` 只能证明"脚本字典里声明了"，**不等于运行时可用**；本报告据此下的结论都标了"可证"级别。

---

## 8. 复现命令

```bash
TOKEN=$(cat ~/.wps-bridge/token)
call() { curl -s -X POST http://127.0.0.1:19890/api/v1/tool/call \
  -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" \
  -d "{\"name\":\"$1\",\"arguments\":$2,\"clientName\":\"probe\",\"sessionId\":\"ms-probe-tester\"}"; echo; }

# ① 窗格新代码探针（看到 addon.capabilities.cap08Shapes=true 即新代码）
call excel_get_charts '{"host":"microsoft"}'
# ② 服务端字节一致性
curl -sk "https://127.0.0.1:19891/office-addon/taskpane.js?_=$(date +%s)" | shasum -a 256
shasum -a 256 office-addon/public/taskpane.js
# ③ 形状域逐项真机自检（selfTestKeep 会把明细留在 _cap08_shapetest，记得删）
call excel_get_charts '{"host":"microsoft","shapeSelfTest":true}'
# ④ 13 项"未实现"的工具路径读数
call excel_copy_range '{"host":"microsoft","sourceRange":"A1:B2","destRange":"D1:E2","sheetName":"<自建临时表>"}'
call excel_manage_hyperlink '{"host":"microsoft","action":"add","address":"A1","url":"https://example.com"}'
# ⑤ 需求集下界真机探针（Worksheet.comments = ExcelApi 1.10）
call excel_manage_cell_comments '{"host":"microsoft","action":"add","address":"A1","text":"probe","sheetName":"<自建临时表>"}'
# ⑥ macOS 原生通道能力取证（不是 Office.js）
sdef "/Applications/Microsoft Excel.app" | grep -o '<command name="[^"]*"' | sort -u   # 无 export as fixed format
defaults read "/Applications/Microsoft Excel.app/Contents/Info.plist" CFBundleShortVersionString   # 16.113
```

---

## 9. 对后续补实现的建议（按收益排序）

1. **先修 3 个"读回炸掉写操作"的缺陷**（影响最大、改动最小）：
   - `loadShapeDetail` 改为**按 shape type 分支 load**（Line/Group 不 load `fill`/`textFrame`），或改成逐个属性 load 并容忍单项失败；
   - `handleListShapes` 合并成**一次** `load()` 调用（或先 load 标量、sync、再对需要的形状单独 load 嵌套属性）——自检里已有正确写法可抄（[shape.js:604-612](../../../../office-addon/src/excel/shape.js)）；
   - 别名表按真实 enum 键修正（`roundRectangle` / `ellipse` / `plus` …），并对 `text_box` 走 `addTextBox`。
   修完后 `excel_list_shapes` / `excel_group_shapes` / `excel_add_shape{kind:"line"}` 都能立刻真机验证。
2. **修 `HOST_IMPLEMENTATION_GAPS.microsoft`**：把已实现的 7 项移出缺口表（`copy_range`、`export_chart_image` 直接可通；其余 5 项同时补 `normalizer` 的方法名映射），并把报错文案改成不臆断原因。
3. **补 3 项新实现**：`get_style_token`、`configure_print_layout`、`set_sheet_view`（网格线/标题；zoom 明确回报 unsupported）。
4. **`getActiveShape` 改调 `workbook.getActiveShape()`**，用 `getActiveShapeOrNullObject()` 避开 `ItemNotFound`，并加 1.19 的需求集守卫。
5. **`export_sheet_pdf` / `format_text_segment` 不要承诺 Office.js 实现**：在工具描述里写清"Microsoft 侧不支持，macOS 请改用 …，Windows 走 COM"，避免 AI 反复试错。
6. **`insert_image` 注册成工具**（或明确并入 `excel_manage_pictures`），消除死分支。
