# 16 · CAP-08 / CAP-52：MS 侧矢量绘图补强

- 开始时间：2026-09-22（本轮）
- 工作目录：`/Users/Python/Office Agent Bridge`
- 写区：`office-addon/src/**`、`office-addon/public/taskpane.html`（仅必要时）；本文件
- 只读参考：`src/bridge/**`、`skills/**`（不改；需要的注册行与 skill 文案见 §6 交接）
- 目标宿主：本机 Microsoft Excel（`工作簿1.xlsx`），Office.js 任务窗格已连接

> ## ⚠️ 读本页前必读：窗格版本
>
> 与 [README](README.md) 里 WPS 侧同一个坑：**磁盘上是新的，宿主进程里跑的是旧的**。
> 本机 Excel 的任务窗格在我改完之前就已打开，而它的热重载链路失效（证据见 §2.1），所以：
> - §3 的**宿主实现 / 读回口径**是"源码 + 离线批处理语义检查"可证的；
> - §4.2 的**真机读数**必须在**重新打开任务窗格之后**取一次；
>   在取到之前，本能力的状态是「实现完成、**真机待验**」，不是「已实测」。

## 台账原文（`docs/acceptance/2.1.0-p0p1/capability-backlog.md`）

> **CAP-08** | **MS 侧矢量绘图**：几何形状、连接符、SVG、文本框、分组、层级、旋转缩放、形状导图
> 依据：Office.js ExcelApi 1.9 `Shape` 全套（**文档可证，未实测**）
> **CAP-52** | MS 侧**矢量绘图域**整体补齐（与 CAP-08 同源）

完成定义（台账末尾统一口径，缺一不算）：宿主实现 + 工具定义 + 读回能力 + 契约快照差异 + skill 说明 + 真机读数 + 回归。
本页交付其中的**宿主实现 / 读回能力 / 真机读数通道 / 回归 / 交接材料**；工具定义与契约快照落在
`src/bridge/**`，不在本任务写区，注册行见 §6。

## 1. 静态盘点结论（源码可证）

- **宿主实现原本只有三个函数**：`office-addon/src/excel/shape.js` 的
  `handleInsertImage` / `handleListShapes` / `handleUpdateShape`（移动/尺寸/删除）。
- **这三个函数还是"死分支"**：`src/rpc.js` 的 `dispatchExcelTool` 有
  `insert_image` / `list_shapes` / `update_shape` 三个 case，但
  `src/bridge/contracts/host-methods.ts` 的 `EXCEL_METHODS` 不含它们，
  所以 `unifiedExcelTools()` 派不出 `excel_insert_image` / `excel_list_shapes` / `excel_update_shape`；
  裸名也不在 `HANDLERS` 里。实测：
  `{"name":"excel_list_shapes","arguments":{"host":"microsoft"}}` → `未知工具 excel_list_shapes`。
- **读回字段严重不足**：原来只 load `name/id/type/left/top/width/height`，
  没有填充色、线条、旋转、层级、文字。
- **`addLine` / `addSvg` / `addTextBox` / `addGroup` / `setZOrder` / `getAsImage` 完全没有实现。**
- 形状层面的"读回"在 MCP 层**没有任何入口**：AI 现在连"这张表里有哪些形状"都问不到。

## 2. 真机探测：先踩到的坑

### 2.1 结论：当前 Excel 里跑的任务窗格是旧版代码，且热重载失效

证据（都是本机读数，不是推断）：

| 探针 | 读数 | 说明 |
|---|---|---|
| `bridge_get_capabilities` → `connection.components.msExcel` | `version: 2.0.0`、**`buildStatus: null`** | 新版 `connection.js` 注册时会带 `buildFingerprint`；为 `null` 说明窗格跑的**不是**磁盘那份 `office-addon/public/taskpane.js` |
| 临时把诊断对象塞进 `handleGetSheetOutline` 返回体后重开构建 | `excel_get_sheet_outline(host=microsoft)` 读回**没有**该字段 | 窗格返回的仍是旧实现 |
| `POST /api/v1/office/reload`（桥接既有热重载路由，发 `run_script: window.location.reload(true)`） | 连续两次返回成功，窗格指纹与诊断**始终不变** | 信号发出去了，窗格没有真正重新加载 |
| `excel_get_charts(host=microsoft)` 读 `addon` 部署探针字段 | 旧窗格**没有** `addon` 字段 | 本轮新增的判据：看到 `addon.capabilities.cap08Shapes=true` 就说明窗格已换新代码 |

原因：`window.location.reload(true)` 的 `true` 是已废弃的 forceGet 参数，Excel for Mac 的 WKWebView
不保证据此重新拉取子资源。窗格里「刷新连接」按钮用的是同一个调用，是同一类问题（本轮已在 `bootstrap.js` 一并修掉）。

### 2.2 为什么不用其它通道探测

- `wps_execute_script`：走 **WPS 表格**加载项（`sockets['excel']`）。实机读数
  `原生脚本执行异常: Excel is not defined` —— 那不是 Office.js，穿不过去。
- `office_execute_script{component:'excel'}`：本机走 **macOS JXA**（`Application('Microsoft Excel')`），
  是 Excel 的 AppleScript/JS 脚本接口，**不是 Office.js**；实测 `target.name()` 报
  `The object you are trying to access does not exist (-1728)`。
- 自己注册 `ms-excel` 通道再下发 `run_script`：只赢 2.5s（真实任务窗格每 2.5s 重连夺回该键），
  实测窗口内拿不到 `run_script` 回包；而且那属于"抢连接"，不该作为常规手段。

### 2.3 由此固化的真机自检通道（本轮新增，**不需要新增 MCP 工具**）

`office-addon/src/excel/shape.js` 的 `handleShapeSelfTest` + `src/excel/chart.js` 里
`get_charts` 的 `shapeSelfTest` 开关：用**已经注册**的 `excel_get_charts(host=microsoft)` 触发，
在真实窗格里逐项跑一遍全部形状 API，返回逐项可用/不可用 + 错误原文 + 耗时。

```
excel_get_charts {"host":"microsoft","shapeSelfTest":true}                        # 跑完自动清理探测表
excel_get_charts {"host":"microsoft","shapeSelfTest":true,"selfTestKeep":true}    # 明细留在 _cap08_shapetest 表
```

窗格侧另有两个入口（重开窗格后可用）：地址带 `#selftest` 自动跑；控制台调 `window.cap08ShapeSelfTest({})`。

## 3. 补齐的宿主实现（`office-addon/src/**`）

`office-addon/src/excel/shape.js`（重写并扩充）：

| 能力 | 实现 | 读回方式 |
|---|---|---|
| 几何形状 | `handleAddShape`（kind=geometric，`addGeometricShape`；别名表 → 枚举键，未知类型抛出可用候选） | 同批次 `loadShapeDetail` 读回位置/尺寸/旋转/填充/线条/文字，`verified` + 逐项差异 |
| 直线 / 连接符 | `handleAddShape`（kind=line，`addLine(x1,y1,x2,y2)`） | 读回包围盒与线条颜色/粗细 |
| SVG | `handleAddShape`（kind=svg，`addSvg(base64)`；接受 `svg` 源码自动转 base64） | 读回 name/位置/尺寸/类型 |
| 文本框 | `handleAddShape`（kind=textbox，`addTextBox`） | 读回文字 + 字体（bold/size/color） |
| 分组 | `handleGroupShapes`（`addGroup([id...])`） | 读回组成员数与成员名，`verified` 比对请求成员数 |
| 解组 | `handleUngroupShapes`（`group.group.ungroup()`） | 读回释放出的子形状是否回到表里；先校验确实是组合类型 |
| 层级 | `handleSetShapeZOrder`（BringToFront/SendToBack/BringForward/SendBackward） | 读回整表 `zOrderPosition` 自底向上序列；端点动作强校验 |
| 形状列表读回 | `handleListShapes`（整批一次 sync） | name/id/type/left/top/width/height/rotation/zOrderPosition/visible/fill/line/text + `zOrderBottomToTop` |
| 形状导图 | `handleExportShapeImage`（`getAsImage(PNG/JPEG, scale)`） | 返回 base64 与估算字节数；空图**抛错**，不返回假图 |
| 缩放 | `handleUpdateShape` 的 `scaleWidth` / `scaleHeight` / `rotationDelta` | 读回 width/height/rotation |
| 删除 | `handleUpdateShape` action=delete | 删除后重新列表确认确实不在 |
| 图片 | `handleInsertImage`（原有，补上读回与参数校验） | 读回完整形状属性 |
| 能力自检 | `handleShapeSelfTest` | 12 项逐项 ok/错误原文/耗时，可写回工作表 |

`office-addon/src/rpc.js`：新增 `add_shape`/`add_geometric_shape`、`group_shapes`、`ungroup_shapes`、
`set_shape_zorder`/`set_z_order`、`get_active_shape`、`export_shape_image`、`shape_self_test`/`probe_shape_api`
分支；并新增 `reload`/`reload_addon` 分支（改写 query 再跳转，替代失效的 `location.reload(true)`）。

`office-addon/src/bootstrap.js`：「刷新连接」按钮改用新的 `reloadTaskPane()`；新增 `#selftest` 自动自检
与 `window.cap08ShapeSelfTest()` 调试入口。

`office-addon/src/state.js`：新增 `ADDON_CAPABILITIES` 部署探针；
`office-addon/src/excel/chart.js`：`get_charts` 返回值里带 `addon: {version, capabilities}`。

### 设计口径（三条硬约束）

1. **写后必读**：每个写操作在同一个批次里把真实落位值读回来返回；请求值与实际值不一致就写进
   `warnings` 并把 `verified` 置 false —— 不做"发了请求就报成功"。
2. **没实现就不假装**：某项属性宿主不接受只影响该项，进 `warnings`，不阻断其它属性、不折叠成成功；
   不支持的整项能力（如该版本没有 `ShapeGroup.ungroup`）**抛错并给出替代路径**。
3. **不用 Canvas 或任何本地合成造图** —— 形状只能由宿主真实创建；`getAsImage` 返回空即报错。

## 4. 真机读数

### 4.1 已完成的真机读数（旧窗格，用于证明"窗格是旧的"）

| 调用 | 读数 |
|---|---|
| `bridge_get_capabilities` | `msExcel.version=2.0.0`、`buildStatus=null` |
| `excel_get_sheet_outline {"host":"microsoft","sheetName":"2026Q3销售分析"}` | `usedRangeAddress='2026Q3销售分析'!A1:F8`、8 行 6 列；**无诊断字段** |
| `excel_get_charts {"host":"microsoft","sheetName":"2026Q3销售分析"}` | `count=1`、`Chart 2`、`ColumnClustered`、`left=334/top=205/width=580/height=340`；**无 `addon` 字段** |
| `excel_read_range {"host":"microsoft","address":"A1"}` | 活动表 `OPPO销售驾驶舱!A1` 读回成功（Office.js 通道活着） |
| `wps_execute_script`（对照） | `原生脚本执行异常: Excel is not defined`（证明 WPS 通道 ≠ Office.js） |

### 4.2 待取读数（窗格重开后一次性执行）

```
① excel_get_charts   {"host":"microsoft","sheetName":"2026Q3销售分析"}
   期望：出现 "addon":{"version":"2.0.0","capabilities":{"cap08Shapes":true,"shapeMethods":[...]}}
② excel_get_charts   {"host":"microsoft","shapeSelfTest":true}
   期望：shapeSelfTest.total=12、逐项 ok、failed=[]（不支持项会带 error 原文）
③ （完成 §6 注册后）
   excel_add_shape          {"host":"microsoft","kind":"geometric","shapeType":"rounded_rectangle","shapeName":"cap08_v","left":40,"top":40,"width":180,"height":90,"fillColor":"#2F6FEB","lineColor":"#12305E","lineWeight":2,"rotation":15,"text":"CAP-08","sheetName":"_cap08_probe"}
   excel_add_shape          {"host":"microsoft","kind":"line","shapeName":"cap08_l","x1":240,"y1":40,"x2":420,"y2":140,"lineColor":"#E4572E","lineWeight":3,"sheetName":"_cap08_probe"}
   excel_add_shape          {"host":"microsoft","kind":"textbox","shapeName":"cap08_t","text":"矢量文本框","left":40,"top":160,"width":200,"height":50,"sheetName":"_cap08_probe"}
   excel_group_shapes       {"host":"microsoft","shapeNames":["cap08_v","cap08_l"],"groupName":"cap08_g","sheetName":"_cap08_probe"}
   excel_set_shape_zorder   {"host":"microsoft","name":"cap08_t","zOrder":"bringToFront","sheetName":"_cap08_probe"}
   excel_list_shapes        {"host":"microsoft","sheetName":"_cap08_probe"}     ← 位置尺寸/填充/线条/层级 全部读回
   excel_ungroup_shapes     {"host":"microsoft","groupName":"cap08_g","sheetName":"_cap08_probe"}
   excel_export_shape_image {"host":"microsoft","name":"cap08_t","format":"png","scale":2,"sheetName":"_cap08_probe"}
   （收尾）excel_delete_sheet {"host":"microsoft","sheetName":"_cap08_probe"}
```

探测表统一用 `_cap08_probe`，用完删除；**不动 `工作簿1.xlsx` 原有的 4 张表**。

## 5. 离线批处理语义检查（不是真机）

`tmp/cap08/offline-check.mjs`：把构建产物 `office-addon/public/taskpane.js` 装进
带 Office.js 批处理语义的 mock（`load`/`sync`/setter 两阶段、`addGroup`/`ungroup`/`setZOrder`/`getAsImage` 行为），
直接跑真实的形状处理函数。**48 项全过**（几何/直线/SVG/文本框/层级/分组/解组/读回列表/更新/删除/错误分支/自检编排）。

它的定位是"动真机之前挡住纯逻辑错误"，**不能替代真机读数**，报告中不得混为一谈。

## 6. 交接：`src/bridge/**` 需要的改动（不在本任务写区）

### 6.1 `src/bridge/gateway.ts` 的 `HANDLERS` 注册行（加在 `// ---- Excel（表格） ----` 段末尾）

```ts
  // CAP-08：MS 侧矢量绘图（形状/连接符/SVG/文本框/分组/层级/导图）
  "excel_add_shape": excel.addShape,
  "excel_list_shapes": excel.listShapes,
  "excel_update_shape": excel.updateShape,
  "excel_group_shapes": excel.groupShapes,
  "excel_ungroup_shapes": excel.ungroupShapes,
  "excel_set_shape_zorder": excel.setShapeZOrder,
  "excel_get_active_shape": excel.getActiveShape,
  "excel_export_shape_image": excel.exportShapeImage,
  "excel_insert_image": excel.insertImage,
```

对应需要在 `src/bridge/gateway/excel.ts` 新增 9 个处理器（写法与既有处理器一致）：

```ts
export const addShape: Handler = async (ctx) => {
  const { args, callOffice } = ctx;
  return await callOffice("add_shape", args);
};
export const listShapes: Handler = async (ctx) => {
  const { args, callOffice } = ctx;
  return await callOffice("list_shapes", args);
};
export const updateShape: Handler = async (ctx) => {
  const { args, callOffice } = ctx;
  return await callOffice("update_shape", args);
};
export const groupShapes: Handler = async (ctx) => {
  const { args, callOffice } = ctx;
  return await callOffice("group_shapes", args);
};
export const ungroupShapes: Handler = async (ctx) => {
  const { args, callOffice } = ctx;
  return await callOffice("ungroup_shapes", args);
};
export const setShapeZOrder: Handler = async (ctx) => {
  const { args, callOffice } = ctx;
  return await callOffice("set_shape_zorder", args);
};
export const getActiveShape: Handler = async (ctx) => {
  const { args, callOffice } = ctx;
  return await callOffice("get_active_shape", args);
};
export const exportShapeImage: Handler = async (ctx) => {
  const { args, callOffice } = ctx;
  return await callOffice("export_shape_image", args);
};
export const insertImage: Handler = async (ctx) => {
  const { args, callOffice } = ctx;
  return await callOffice("insert_image", args);
};
```

### 6.2 必改：`src/bridge/contracts/host-methods.ts` 的 `EXCEL_METHODS`

**只加 HANDLERS 注册行是不够的**：`src/bridge/tools/excel.ts` 的 `unifiedExcelTools()` 用
`EXCEL_METHODS.includes(tool.name.replace(/^wps_/, ''))` 过滤，只有在这里追加方法名，
schema / gateway / host 三层才同时打通（`check:claims` 的「Microsoft 可调用 28 项」口径会随之变成 37 项，
需同步复核 `office-addon/public/taskpane.html:38` 的「28 项能力」文案，否则 `check:claims` 会拦下）：

```ts
  'rollback_cells', 'save_workbook', 'get_style_token', 'format_text_segment',
  // CAP-08：MS 侧矢量绘图
  'add_shape', 'list_shapes', 'update_shape', 'group_shapes', 'ungroup_shapes',
  'set_shape_zorder', 'get_active_shape', 'export_shape_image', 'insert_image'
```

同时：

- `REPLAY_SAFE_METHODS` 与 `MCP_READ_ONLY_TOOLS` 应加入 `list_shapes`、`get_active_shape`（只读、不改文档）；
  **不要**加 `export_shape_image`（会写本地文件）与其它写方法。
- `HOST_IMPLEMENTATION_GAPS.wps` 需加入这 9 个方法名：WPS 加载项没有对应 RPC 分支，
  这样 `host=wps` 的调用才会被正确归类为"宿主未实现"，而不是打到 WPS 加载项的 default 分支。
- `COM_IMPLEMENTATION_GAPS`：这 9 个方法在 `resources/office/excel.ps1` 里都没有实现，必须一并列入，
  否则 Office.js 通道不可用时会误判为"可回退到原生 COM"。
- `src/bridge/office/normalizer.ts`：新方法走 `default:` 原样透传即可
  （加载项读的字段名就是 schema 字段名，不要再维护第二份映射，避免两层漂移）。

### 6.3 `src/bridge/tools/definitions/excel.ts` 的 schema 参数名（与宿主实现一一对应）

| 工具 | 必填 | 可选 |
|---|---|---|
| `add_shape` | `host`, `kind`（geometric/line/svg/textbox） | `sheetName`, `shapeType`, `shapeName`, `text`, `left/top/width/height`, `x1/y1/x2/y2`, `svg`, `base64Image`, `fillColor`, `lineColor`, `lineWeight`, `rotation` |
| `list_shapes` | `host` | `sheetName` |
| `update_shape` | `host`, `name` | `sheetName`, `action`(delete), `left/top/width/height`, `rotation`, `rotationDelta`, `scaleWidth`, `scaleHeight`, `scaleType`(currentSize/originalSize), `fillColor`, `lineColor`, `lineWeight`, `text`, `newName` |
| `group_shapes` | `host`, `shapeNames[]`（≥2） | `sheetName`, `groupName` |
| `ungroup_shapes` | `host`, `groupName` | `sheetName` |
| `set_shape_zorder` | `host`, `name`, `zOrder`(bringToFront/sendToBack/bringForward/sendBackward) | `sheetName` |
| `get_active_shape` | `host` | `sheetName` |
| `export_shape_image` | `host`, `name` | `sheetName`, `format`(png/jpeg), `scale` |
| `insert_image` | `host`, `base64Image` | `sheetName`, `left/top/width/height`, `name` |

返回值统一特征（写进工具描述，让 AI 知道可以自检）：`shape`（宿主读回的实际属性）、
`verified`（请求值是否全部读回一致）、`warnings`（不一致或未生效项）。

### 6.4 skill 说明需要补的要点（`skills/office-agent-bridge/**`，不在本任务写区）

1. MS 侧画图形用 `excel_add_shape`（`kind` 选 geometric / line / svg / textbox）；
   几何类型支持直觉名（`rounded_rectangle`/`oval`/`arrow`/`flow_chart_decision`…），未知类型会报可用候选；
2. 写完**必看**返回的 `shape` 与 `verified`/`warnings`，不要只看 `success`；
3. 形状多时先 `excel_list_shapes` 拿 name（name 才是稳定定位键，id 是 `{GUID}`）；
4. 层级用 `excel_set_shape_zorder`，返回 `zOrderBottomToTop` 可直接核对；
5. 需要视觉自检就用 `excel_export_shape_image` 导 PNG，再自己看；
6. **宿主限制**：`host=wps` 侧没有这 9 个方法（WPS 走 `wps_execute_script` 的形状 API）；
7. 改了 `office-addon/src/**` 之后，**必须重新打开任务窗格**才生效（热重载链路本轮才修，旧窗格仍跑旧代码）。

## 7. 回归结果

| 检查 | 结果 |
|---|---|
| `npm run build:office-addon` | 通过（19 模块，语法校验通过） |
| `node --check office-addon/public/taskpane.js` | 通过 |
| `node scripts/build-office-addon.mjs --check` | 生成物与源码一致 |
| `npm run typecheck` | 退出码 0 |
| `npm run check:claims` | 通过（路由表 34 项 / 对外工具 105 个 / Microsoft 可调用 28 项；**§6.2 改完需重跑**） |
| `npm test` | 95/95 通过 |
| 离线批处理语义检查（`tmp/cap08/offline-check.mjs`） | 48/48 通过（**非真机**） |

## 8. 遗留与如实说明

1. **真机读数待补**：本机 Excel 任务窗格仍是旧代码，§4.2 的读数必须重开窗格后取。
   在拿到之前，本能力状态是「实现完成、真机待验」。
2. **仍未实测的 Office.js 形状细节**（自检会逐项报，不靠文档推断）：
   `Shape.group.ungroup()` 是否存在于本机 API 集、`getAsImage` 的 `PictureFormat` 枚举与可用 scale、
   `zOrderPosition` 的编号基准（0 还是 1）、`addSvg` 对 base64 的容忍度（是否需要 data URI 前缀）、
   `incrementRotation` 是否可用。
3. **`insert_dimension` 等与 CAP-08 无关的既有缺口**不在本轮范围。
4. **`excel_export_shape_image` 会写本地文件**（与 `capture_sheet_preview` 同类），
   不应标 readOnlyHint，也不应进跨通道重放白名单。
5. 本轮**没有**动 `office-addon/public/taskpane.html`（不需要）。
6. `tmp/cap08/**` 是本轮探测脚手架（离线检查 + 探针脚本），属临时材料、不进仓库正式目录；
   若要保留，建议迁到 `scripts/` 并加进导航。
