# MS Office 与 WPS 能力对照 + 补强建议

> 状态：**已完成（§0–§5 全部填充；未实测项已在 §5 逐条声明）**
> 范围：`office-addon/src/**`（Office.js 侧）vs `wps-addon/src/**`（WPS 侧）+ 桥接网关宿主差异声明
> 结论依据等级：**源码可证** / **文档可证** / **推测**
> 实测限制：本机 **Microsoft Excel 加载项未连接**，MS 侧全部结论均来自源码阅读与官方文档，**无实测**
> 只读分析，未修改仓库任何源码或既有文档；中间产物在 `.scratch/msvs/`

---

## 0. 摘要与结论速览

### 0.1 一句话结论

> **不是"WPS 比 MS 弱"。真实情况是：WPS 宿主 API 面（VBA/COM 的近乎全量克隆）≫ Office.js 面；而我们的加载项只用了其中很小一撮，且 MS 侧那一小撮还有 8 处"字段名对不上 → 静默 no-op / 结果错误"。**
>
> 也就是说，本轮最大的发现**不是缺能力，而是"以为改了、其实没改"**。

### 0.2 两条通道覆盖的能力域

| | **WPS 加载项** | **Office.js 加载项（MS）** | MS 侧原生（COM/JXA） |
|---|---|---|---|
| 组件覆盖 | **表格 + 文字 + 演示** | **仅 Excel** | Windows：Excel 结构化（28 项）；word/ppt 仅任意脚本 |
| RPC 方法数 | **61**（表格 30 / 文字 13 / 演示 9 / 平台 9） | **77 个标签 → 43 个处理器**（全是 Excel） | Excel 28 个方法 + `script` |
| 路由表覆盖（`EXCEL_METHODS` 30 项） | **29/30**（缺 `update_chart`） | **29/30**（缺 `rollback_cells`） | **28/30**（缺 `update_chart`、`save_workbook`） |
| 表格：读写与结构 | ✅ 区域读写、查找替换、行列增删隐藏；❌ **无区域复制** | ✅ 读写、查找替换、行列结构、**区域复制**、超链接 | 同 Office.js |
| 表格：格式化 | ✅ 字体/填充/边框/对齐/数字格式/合并（**带读回校验**） | ✅ 同上 + 9 个 `OPPO_*` 命名样式模板 | 同 Office.js |
| 表格：条件格式 | ✅ cell_value / data_bar / color_scale（可设两端色） | ✅ cellValue / colorScale（**不设色**）/ dataBar（可设色） | 同 Office.js |
| 表格：图表 | ✅ 建/查/删；`dataRanges` 联合源、`yAxis`、`seriesSettings`；**无 update** | ✅ 建/查/删/**改**/导出图片；单元格锚定定位；**无 yAxis/逐系列** | 同 Office.js |
| 表格：透视表 | ✅ 建表 + **字段编排**；无刷新 | ✅ 建表 + refresh；**无字段编排** | 同 Office.js |
| 表格：形状/图片 | ❌ **完全没实现**（宿主 `Shapes` 全套能力闲置） | ⚠️ **只有 `addImage`**（宿主有全套：几何形状/连接符/SVG/文本框/分组/导图） | 无 |
| 表格：结构化表格 / 命名区域 / 文档属性 | ❌ 未实现（宿主都支持） | ✅ 已实现 | ✅ 命名区域/属性 |
| 表格：工作表管理 | ✅ rename/move/tab_color/protect/unprotect | ✅ rename/copy；**move/tab_color/protect 静默 no-op** | protect/move |
| 表格：截图/预览 | ✅ **`CopyPicture` 真实像素渲染** | ⚠️ 有图表时真图；**无图表时 Canvas 硬编码伪渲染** | ✅ 真实渲染（`CopyPicture`） |
| 表格：保存/导出 | ✅ `save_workbook`；**无 PDF** | ✅ `save`；**无 PDF** | ⚠️ 结构化方法**无 `save_workbook`**；只能经 `script` 调 `SaveAs` / `ExportAsFixedFormat` |
| 表格：分页与打印 | ❌ 未实现（宿主有 `PageSetup`） | ❌ 未实现（宿主有 `PageLayout`/`PageBreaks`） | 未实现 |
| 文字（Word） | ✅ **13 个工具**：建/存(PDF)/关/删/读/写/排版(3 套模板)/目录/表格/修订+批注/页面设置+水印/查找替换 | ❌ **无结构化通道** | 仅 `office_execute_script` |
| 演示（PPT） | ✅ **9 个工具**：读/形状探测/整册生成/页面/表格/卡片/原生图表/形状媒体/截图 | ❌ **无结构化通道** | 仅脚本 + 导页图 |
| 逃生舱 | ✅ `execute_script`（7 参注入 + console 捕获 + 序列化保护）、`inspect_api`（运行时反射） | ✅ `run_script`（2 参注入，无保护） | ✅ `script` |
| 事件 | 宿主有 45+ 事件 + `wps.ApiEvent`（**我们未用**） | ExcelApi 1.17 有 3 个工作表事件（**我们未用**） | 无 |

### 0.3 最值得补的 3 条

| 排名 | 建议 | 补在哪一侧 | 一句话理由 |
|---|---|---|---|
| **1** | **修 M1–M8 跨宿主字段契约**（§3.5）：`set_data_validation` 静默清空校验、`manage_sheet` 4 个 action 静默 no-op、`find_and_replace` 全表误匹配、`get_charts` 选择器被忽略…… | **桥接网关**（normalizer 补映射）+ **Office.js 加载项**（补 action 与 `rule` 构造） | 这是**正确性缺陷**：现有工具在 MS 侧"返回 success 但文件里没有变化"，直接摧毁 AI 的自检可信度 |
| **2** | **补全 MS 侧 Excel 矢量绘图域**（几何形状 / 连接符 / SVG / 文本框 / 分组 / 层级 / 形状导图） | **Office.js 加载项** | 宿主 API 全套都在（文档可证），我们只实现了 `addImage` → **MS 用户完全用不上"矢量信息图"这一能力**，而 WPS 侧已有对应物 |
| **3** | **把 `capture_sheet_preview` 的伪渲染换成真实渲染** | **Office.js 加载项**（无图表时改走 Windows COM 的真实 `CopyPicture` 路径，或显式标 `synthetic`） | 现在 MS 侧无图表时返回的是**Canvas 硬编码合成的假图**（`cellW=110/rowH=28`），AI 的视觉自检会得出与真实文件不符的结论 |

> 紧随其后的是 **S4 分页/打印/PDF 导出**（两边全空白，但它是交付链路的最后一公里）。

### 0.4 WPS 相对 MS 的明显短板

**宿主层面：几乎没有。** 官方文档显示 WPS 表格的对象模型是 VBA/COM 的近乎全量克隆，甚至**比 Office.js 更全**：WPS 有 `AddSmartArt`、艺术字 `AddTextEffect`、`BuildFreeform` 任意多边形、`ThreeD`/`Shadow`/`Glow`/`Reflection`/`SoftEdge`、格式刷 `PickUp`+`Apply`、`Application.Run`/`Evaluate`、`ExportAsFixedFormat`、45+ 事件、`Range.Characters`、`Names`、`FileSystem`/`PluginStorage`/`ApiEvent` —— 这些 Office.js 全都没有。

**我方实现层面：有 4 处明显短板**（源码可证）：

1. **表格侧形状/图片能力为零** —— 宿主 `Shapes` 全套（`AddShape`/`AddConnector`/`AddTextEffect`/`AddSmartArt`/`BuildFreeform`/`ThreeD`…）一个都没暴露。
2. **冻结窗格依赖 `ActiveWindow`** —— 必须先 `Activate` 工作表，跨簿/多窗口会打错目标；Office.js 侧则是与活动窗口无关的对象化 API。
3. **`update_chart` 缺失** —— 唯一被 `HOST_IMPLEMENTATION_GAPS` 显式声明的方法级缺口，改图表标题/位置只能删了重建。
4. **`add_chart` 的定位是磅值像素而非单元格锚定** —— `position.leftCell` 只取一次 `Left/Top` 当像素初值，行列尺寸一变就漂移；Office.js 侧是真正的 `setPosition(startCell, endCell)`。

另外还有 2 处**粒度**短板（不完全是弱项，属互补）：条件格式的 `dataBar` 不能设渐变/边框色；透视表只能建不能刷新。

### 0.5 哪些结论只是文档推断、没实测

**⚠️ 本机 Microsoft Excel 加载项未连接，MS 侧零实测。** 具体分级：

- **纯文档推断（未实测，风险最高）**：
  - §3.5 的 **M1–M8 全部**——"字段不匹配"本身是**源码可证**，但"因此静默 no-op / 全表误匹配 / 破坏性替换"的**运行时后果是代码路径推断**，必须实机复现才能定级。
  - §4 中 S2 / S4 / S6 / S7 依赖的 Office.js 要求集（`ExcelApi 1.9` 需 Office 1808+，`1.17` 需 2302+）——**用户实际 Office 版本未探测**，若版本偏旧则方案不可用。
- **文档可证但未在加载项沙箱内验证**：WPS 的 `Shapes.AddShape`/`AddSmartArt`、`Range.Characters`、45+ 事件、`PageSetup`、`FileSystem`/`PluginStorage`/`ApiEvent`——**"官方文档有" ≠ "当前 WPS 12.1.28496 已实现" ≠ "加载项沙箱内可调用"**，三层都要实测。
- **明确标注为"推测"的条目**（见 §5.2）：①-2 `RangeAreas`、①-3 SVG、①-4 新图表类型、①-5 图表细粒度、①-8 批处理语义，以及 WPS 侧若干页面对应能力。
- **Windows COM 路径**：`resources/office/excel.ps1` 在本机（macOS）**无法执行**，S3/S4 里"改走 COM"的建议需要 Windows 实机验收。

---

## 1. 我方两条通道的实现盘点

> 盘点口径：**RPC 分发 switch 的 `case` 标签**（宿主侧真实可达的方法），不是 MCP 工具名。
> 静态统计（源码可证）：
>
> | 通道 | 源文件 | RPC 方法标签 | 去别名后业务处理器 | 覆盖组件 |
> |---|---|---|---|---|
> | WPS | `wps-addon/src/**` 9 模块（7250 行） | **61** | 52 | 表格 + 文字 + 演示 |
> | Office.js | `office-addon/src/**` 19 片段 | **77** | 43 | **仅 Excel** |
>
> MS 侧另有**非 Office.js** 的两条原生路径（`resources/office/*.ps1` + `ms-office-driver.ts`）：
> Windows COM（结构化方法**只有 Excel**，另有 `script` 任意脚本 action）与 macOS JXA（**只有任意脚本**，无结构化方法）。
> 也就是说：**Microsoft 上没有任何结构化 Word / PPT 通道**，只能走 `office_execute_script`。

### 1.1 Office.js 侧（`office-addon/src/**`，19 个片段，仅 Excel）

43 个处理器按能力域：

| # | 能力域 | RPC 别名 | 实现要点（源码可证） |
|---|---|---|---|
| 1 | 工作簿与元数据 | `get_workbook_info` / `get_workbook_state` / `get_workspace_summary`、`get_sheet_outline`、`save` / `save_workbook`、`calculate`、`list_named_items`、`update_named_item`、`get_document_properties`、`update_document_properties` | **命名区域（`workbook.names`）与文档属性（`workbook.properties` 8 字段）只有这一侧有**；`calculate` 是独立方法 |
| 2 | 工作表生命周期 | `list_sheets` / `get_sheets`、`add_sheet` / `create_sheet`、`update_sheet` / `rename_sheet`、`delete_sheet`、`copy_sheet` / `duplicate_sheet` / `manage_sheet` | 重命名有独立方法，不靠 `manage_sheet` 的 action |
| 3 | 区域读写与结构 | `read_range` / `get_range_data` / `get_used_range`、`get_range_styles`、`write_range` / `patch_cells` / `set_range_data`、`update_range_structure` / `insert_rows` / `delete_rows` / `modify_rows_columns` / `manage_rows_and_columns` / `insert_dimension`、`clear_range` / `clear_cells`、`copy_range`、`set_hyperlink`、`set_data_validation`、`find_replace_cells` / `find_and_replace` / `find_replace` / `search_cells` | **`copy_range`、`set_hyperlink` 只有这一侧有**；读回带 `text` 与 `numberFormat` |
| 4 | 公式 | `set_formula` / `set_cell_formula` | 独立方法，支持二维 `formulas` 矩阵 |
| 5 | 格式与排版 | `format_cells` / `format_range` / `set_range_format`、`auto_fit_columns` / `autofit_columns`、`freeze_panes`、`add_conditional_formatting` / `list_conditional_formats` / `update_conditional_format` | 含 **9 个 `OPPO_*` 命名样式模板**；`freeze_panes` 走 `freezeAt` / `freezeRows` / `freezeColumns`，**先 `unfreeze` 再冻结**以规避 `InternalError`；条件格式支持 `cellValue` / `colorScale` / `dataBar`（`dataBar` 可设正格式填充色、边框色、渐变开关） |
| 6 | 排序与筛选 | `sort_range` / `apply_filter` / `set_filter_and_sort` | `autoFilter.apply(..., {filterOn: values, values})`；排序用 `range.sort.apply([{key, ascending, sortOn}], hasHeader)` |
| 7 | 结构化表格 Table | `create_table`、`update_table` | **`tables.add` / `style` / `showTotals` / `resize` 只有这一侧有**（WPS 侧完全没有 ListObject） |
| 8 | 图表 | `get_charts`、`add_chart` / `create_chart`、`delete_chart`、`update_chart`、`export_chart_image` | `update_chart`（改标题 / 图例位置 / 位置尺寸）**只有这一侧有**；定位优先 `chart.setPosition(startCell, endCell)` 单元格锚定，失败退化为 `left/top/width/height` 像素；饼图/环形图可按数据点逐点着色；`chart.getImage()` 可直接导出图表图片 |
| 9 | 数据透视表 | `create_pivot_table`、`update_pivot_table` | `update_pivot_table` 支持 `refresh`——**WPS 侧没有刷新/更新** |
| 10 | 形状与图片 | `insert_image`、`list_shapes`、`update_shape` | `sheet.shapes.addImage(base64)`；列出/移动/删除形状。**WPS 表格侧完全没有对应能力** |
| 11 | 审阅与批注 | `manage_cell_comments` / `add_comment` / `list_comments` / `update_comment` | 仅实现 `add` / `list` / `delete`；`update_comment` 别名虽在 switch 里，但会落进同一处理器且**不匹配任何 action 分支 → 静默返回 `{success:true}`**（源码可证，见 §3.3 注） |
| 12 | 预览截图 | `capture_sheet_preview` | `chart.getImage()` 渲染 |
| 13 | 任意脚本 | `run_script` / `execute_script` | `new Function("context","Excel", script)` 在 `Excel.run` 内执行 |

> **名义覆盖 ≠ 实际覆盖**：`list_conditional_formats`、`update_conditional_format`、`update_comment` 三个别名都落进"只做一半"的处理器，属于**别名齐全、语义缺失**（源码可证）。

### 1.2 WPS 侧（`wps-addon/src/**`，9 个模块，61 个 RPC 方法）

| 组件 | 方法数 | RPC 方法 |
|---|---|---|
| 表格（`excel.js` 1844 行） | 30 | `get_workspace_summary`、`get_sheet_outline`、`get_style_token`、`create_sheet`、`delete_sheet`、`clear_range`、`read_range`、`get_range_styles`、`search_cells`、`patch_cells`、`format_cells`、`add_conditional_formatting`、`freeze_panes`、`modify_rows_columns`、`manage_rows_and_columns`、`auto_fit_columns`、`rollback_cells`、`insert_dimension`、`capture_sheet_preview`、`add_chart`、`get_charts`、`delete_chart`、`create_pivot_table`、`set_filter_and_sort`、`set_data_validation`、`manage_sheet`、`manage_cell_comments`、`find_and_replace`、`duplicate_sheet`、`save_workbook` |
| 文字（`word.js` 935 行） | 13 | `word_create_document`、`word_save_document`、`word_close_document`、`word_manage_content`、`word_read_document`、`word_write_content`、`word_format_document`、`word_insert_table_of_contents`、`word_manage_table`、`word_review_and_comments`、`word_page_layout_and_watermark`、`word_find_and_replace`、`word_capture_preview` |
| 演示（`ppt.js` 1225 行 + `ppt-layout.js`） | 9 | `ppt_read_presentation`、`ppt_get_slide_shapes`、`ppt_generate_deck`、`ppt_manage_slides`、`ppt_manage_table`、`ppt_add_business_cards`、`ppt_insert_native_chart`、`ppt_manage_shapes_and_media`、`ppt_capture_slide_preview` |
| 平台/脚本/锁 | 9 | `ping`、`lock_target_document`、`unlock_target_document`、`get_locked_status`、`eval` / `eval_code` / `execute_script`、`inspect_api`、`reload` |

WPS 侧的**深度**亮点（源码可证）：

- **表格**：`get_style_token`（把工作簿主题/样式令牌读出来给 AI 当调色板）；`rollback_cells`（宿主侧快照回填，配合 `patch_cells` 的前后快照）；`add_chart` 支持 `dataRanges` **非连续区域联合数据源**、`yAxis`（min/max/step/numberFormat/title）、`seriesSettings`（逐系列颜色与平滑）、`smoothLine`、`seriesColors`；`manage_sheet` 有 `tab_color` / `protect` / `unprotect` / `move`；`manage_cell_comments` 有 `clear_all`。
- **文字**：`word_insert_table_of_contents`（目录域）、`word_page_layout_and_watermark`（节 + 页面设置 + 水印）、`word_review_and_comments` 含**修订开关与全部接受/拒绝**（`enable_track_changes` / `disable_track_changes` / `accept_all_revisions` / `reject_all_revisions`）+ 批注增列、`word_manage_table` 有 `merge_cells`、`word_format_document` 有 `gov_standard` / `business_modern` / `academic` 三套公文/商务/学术排版模板。
- **演示**：`ppt_generate_deck`（结构化 JSON → 真实页面尺寸排版，含 `ppt-layout.js` 的字号容量估算与 `fitGeneratedPptShapes` 缩放换算）、`ppt_add_business_cards`、`ppt_manage_shapes_and_media` 有 `swap_shapes` / `set_z_order` / `align_shapes`、`ppt_manage_slides` 有 `set_background`、`ppt_manage_table` 有 `merge_cells`（推定，见 §5）。
- **平台**：`inspect_api` 运行时反射探测任意对象（WPS 独有，MS 侧无对应物）；`execute_script` 是带 `app` / `doc` / `wb` / `pres` / `wps` / `params` / `console` 七参数注入的 `AsyncFunction` 沙箱，**比 Office.js 的 `new Function("context","Excel")` 上下文更宽**。

### 1.3 只有一边实现的操作（我方差集）

#### A. 只有 Office.js（MS）侧有、WPS 侧无 —— 14 项

| 能力 | 依据 |
|---|---|
| 命名区域增删查（`list_named_items` / `update_named_item`） | 源码可证：WPS `dispatch.js` 全文无 `named` 分支 |
| 文档属性读写（`get_document_properties` / `update_document_properties`） | 同上 |
| 结构化表格 Table（`create_table` / `update_table`，含样式、汇总行、resize） | 同上 |
| 区域复制（`copy_range`） | 同上 |
| 超链接（`set_hyperlink`） | 同上 |
| 图片插入与形状管理（`insert_image` / `list_shapes` / `update_shape`） | 同上（WPS 的形状能力只在 PPT 侧） |
| 图表更新（`update_chart`：改标题/图例位置/位置尺寸） | 源码可证 + `HOST_IMPLEMENTATION_GAPS.wps = ['update_chart']` |
| 图表图片导出（`export_chart_image`） | 同上（WPS 只有 `capture_sheet_preview` 剪贴板截图路径） |
| 透视表刷新（`update_pivot_table`） | 同上 |
| 重算（`calculate`） | 同上（WPS 在 `addChart` 内部隐式 `app.Calculate()`，无对外方法） |
| 工作表重命名独立方法（`update_sheet` / `rename_sheet`） | WPS 用 `manage_sheet` 的 `rename` action 覆盖 |
| 条件格式 `dataBar` 的填充色/边框色/渐变开关 | 源码可证：WPS 的 dataBar 只设 `BarColor.Color` |
| 条件格式列表/更新（`list_conditional_formats` / `update_conditional_format`） | **仅别名存在**，Office.js 处理器未实现语义 |
| 排序支持多键（`sort_range` 走 `range.sort.apply` 数组） | WPS 的 `setFilterAndSort` 也支持多规则（`sortRules`），此项实际两边都有，**不计入差集** |

#### B. 只有 WPS 侧有、MS 侧无 —— 结构化 Word/PPT 全量 + 5 项 Excel 独有

| 能力 | 依据 |
|---|---|
| **Word 全部 13 个方法** | 源码可证：`office-addon` 只有 `excel/manifest.xml`；MS 侧 Word 只能 `office_execute_script` |
| **PPT 全部 9 个方法** | 同上 |
| `get_style_token`（主题/样式令牌） | 源码可证：Office.js `dispatchExcelTool` 无此分支 |
| `rollback_cells`（宿主侧快照回填） | 同上。**Office.js 侧无此分支**（Windows COM 侧有）；MS 侧回滚在 Office.js 通道上会抛"暂未映射该工具"，在 macOS 上硬失败、在 Windows 上才可能回退到 COM |
| `inspect_api`（任意表达式反射探测） | 同上 |
| `manage_sheet` 的 `tab_color` / `protect` / `unprotect` | 源码可证：Office.js `handleManageSheet` 无 protect/tabColor 分支 |
| `add_chart` 的 `dataRanges` 非连续区域联合源、`yAxis` 坐标轴细调、`seriesSettings` 逐系列 | 源码可证：Office.js `handleCreateChart` 无 yAxis/seriesSettings 参数 |
| `format_cells` 的列宽（`columnWidth`） | Office.js 侧 `handleFormatRange` 也支持 `columnWidth`，**不计入差集** |
| 跨组件文档锁（`lock_target_document` / `unlock_target_document` / `get_locked_status` 在加载项内） | 源码可证；MS 侧的锁在桥接层 `MsOfficeDriver.lockTarget`，不在加载项内 |

#### C. 两边都"有名字"但深度差距明显（Excel 域）

| 操作 | WPS 侧深度 | Office.js 侧深度 |
|---|---|---|
| `add_chart` 参数面 | `dataRanges`、`yAxis{min,max,step,numberFormat,title}`、`seriesSettings[]`、`smoothLine`、`seriesColors[]`、`hasDataLabels`、`position.leftCell` **单元格锚定** | `seriesBy`、`seriesColors[]`、饼图逐点着色、`setPosition(startCell,endCell)` **单元格锚定**、`legend.visible`；**无 yAxis、无逐系列设置、无数据标签开关** |
| `add_conditional_formatting` | `cell_value` / `data_bar`（可设色）/ `color_scale`（可设 min/max 色） | `cellValue`（可设填充/字色）/ `colorScale`（**只 `add`，不设任何颜色**）/ `dataBar`（可设填充/边框色/渐变） |
| 读回校验 | `format_cells` 的 `merge` 失败会**读回校验并抛错**（ISS-42 修复） | `range.merge()` 后不读回 |
| 隐藏行/列 | 逐行 `Rows.Item(n).Hidden` 写 + 读回校验（ISS-60 修复），列用 `Columns.Item` | `update_range_structure` 走 `Range` 整块操作 |
| 截图 | `CopyPicture` → 临时 ChartObject → `Chart.Export` PNG（**像素级工作表渲染**） | `chart.getImage()`（**只能导出图表**，非整表渲染）→ 见 §3.3 待核 |

### 1.4 桥接网关的宿主差异声明（已有的"官方口径"）

`capabilities()`（[catalog.ts](../../../../../src/bridge/catalog.ts)）已经在向外声明宿主差异，本轮分析与之一致：

| 声明项 | 内容 |
|---|---|
| `hosts.wps.excel.unimplementedOnHost` | `['update_chart']` |
| `hosts.microsoft.excel.unimplementedOnHost` | `[]` |
| `hosts.wps.word` | "有限支持：读取、编辑、保存；按实际工具清单使用" |
| `hosts.microsoft.word` | "**未提供 Microsoft 结构化通道**：Word 工具为 WPS 专用；需改用 `office_execute_script`" |
| `hosts.microsoft.ppt` | 同上 |
| `hosts.microsoft.excel.validation` | "代码路径已实现，但**本候选版本尚未在真实 Microsoft Excel 上做实机验收**；静态与模拟通道不代表实机通过" |
| `HOST_IMPLEMENTATION_GAPS` 注释 | "`wps-addon/addon-core.js` 的 RPC 分发 switch 只有 add_chart / get_charts / delete_chart，没有 update_chart；未识别方法会走 `default:` 抛错" |

**对账结论（源码可证）**：该注释描述的 **WPS 缺 `update_chart`** 与当前 `dispatch.js` 一致（§1.2 的 30 个表格方法里确实没有 `update_chart`）。`declaredNotCallable` 由"路由表 ∩ 已注册工具"派生，避免回显路由表——这一点与 [compatibility-diff.md](../../compatibility-diff.md) §3 的记录一致。

> ⚠️ **口径差异（本轮新发现）**：`HOST_IMPLEMENTATION_GAPS.microsoft = []` 只覆盖**表格**。从 §1.3 B 可见，Microsoft 侧的**真实缺口是 Word 与 PowerPoint 两个组件完全没有结构化通道**——该缺口目前只写在 `capabilities().hosts.microsoft.word/ppt` 的**自然语言字符串**里，**没有进入机读的结构化缺口台账**（没有 `unimplementedOnHost` 之类的字段）。AI 若只读 `hosts.microsoft.excel` 会得到"MS 侧无缺口"的错误印象。

**三条通道对同一张路由表（`EXCEL_METHODS`，30 项）的实际覆盖率**（脚本核对，源码可证）：

| 通道 | 覆盖 | 缺哪几项 |
|---|---|---|
| WPS 加载项（`wps-addon/src/dispatch.js`，61 个 RPC 标签） | **29 / 30** | `update_chart` |
| Office.js 加载项（`office-addon/src/rpc.js`，77 个标签） | **29 / 30** | `rollback_cells` |
| Windows COM（`resources/office/excel.ps1`） | **28 / 30** | `update_chart`、`save_workbook` |

> 三边互补但仍然**没有一项是三边全通的**：`update_chart` 在 WPS 与 COM 都缺（幸好两边的 `HOST_IMPLEMENTATION_GAPS` 只声明了 WPS 那一条），`rollback_cells` 在 Office.js 缺（但 COM 有，Windows 上可回退）。

---

## 2. 官方 API 面对照（文档依据）

> 本节全部为**文档可证**（官方文档明确写了），不含实测。所有条目附官方 URL。

### 2.1 Office.js Excel / Word / PowerPoint API 面

**版本梯度（截至查证日）**

| 产品 | 要求集上限 | 来源 |
|---|---|---|
| Excel | **ExcelApi 1.21** + `ExcelApiDesktop 1.1`（仅 Windows 2509+/Mac 16.102+）+ `ExcelApiOnline` | [Excel 要求集](https://learn.microsoft.com/zh-cn/javascript/api/requirement-sets/excel/excel-api-requirement-sets) |
| Word | **WordApi 1.9** + `WordApiDesktop 1.5` + `WordApiHiddenDocument 1.5` | [Word 要求集](https://learn.microsoft.com/zh-cn/javascript/api/requirement-sets/word/word-api-requirement-sets) |
| PowerPoint | **PowerPointApi 1.10**（Mac 支持到 1.10；**iPad 只到 1.1**） | [PowerPoint 要求集](https://learn.microsoft.com/zh-cn/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets) |

**关键节点能力（对"精细操控"直接相关）**

| 要求集 | 新增的高价值能力 | 来源 |
|---|---|---|
| **ExcelApi 1.9**（一次性 500+ API） | `Shape`/`ShapeCollection`/`GeometricShape`/`Image`；`AutoFilter`；**`RangeAreas` 非连续区域一等类型**；`Range.getSpecialCells`；`Worksheet.findAll` / `Range.find`；`Range.copyFrom`；`Application.suspendScreenUpdatingUntilNextSync` / `iterativeCalculation` / `calculationState`；**新图表类型：地图、箱形图、瀑布图、旭日图、排列图、漏斗图**；`RangeFormat`；**`PageLayout`（打印区域 / 打印标题行列 / 页边距 / 纸张 / 方向 / 页眉页脚 / 网格线 / 标题 / 批注 / 错误 / 缩放）**；**`PageBreakCollection`**；`HeaderFooter` / `HeaderFooterGroup`；`CellProperties` / `CellPropertiesFormat`（`indentLevel` / `shrinkToFit` / `textOrientation` / `readingOrder` / `autoIndent` / `protection`）；`ConditionalFormat.getRanges`；`DataValidation.getInvalidCells`；`Line`（`connectBeginShape` / `connectEndShape` + 箭头长度/样式/宽度）；`ChartSeries` 误差线 / 趋势线 / 箱形图选项 / 渐变 | [ExcelApi 1.9](https://learn.microsoft.com/zh-cn/javascript/api/requirement-sets/excel/excel-api-1-9-requirement-set) |
| **ExcelApi 1.17** | **条件格式规则"改造"**：`changeRuleToCellValue` / `ColorScale` / `ContainsText` / `Custom` / `DataBar` / `IconSet` / `PresetCriteria` / `TopBottom` + `setRanges`；**工作表事件**：`onNameChanged` / `onVisibilityChanged` / `onMoved` | [ExcelApi 1.17](https://learn.microsoft.com/zh-cn/javascript/api/requirement-sets/excel/excel-api-1-17-requirement-set) |
| **WordApi 1.9** | **下拉列表 / 组合框内容控件**（`DropDownListContentControl` / `ComboBoxContentControl` / `ContentControlListItem`：`addListItem` / `select` / `deleteAllListItems`） | [WordApi 1.9](https://learn.microsoft.com/zh-cn/javascript/api/requirement-sets/word/word-api-1-9-requirement-set) |
| **PowerPointApi 1.1→1.10** | 1.2 插入/删除幻灯片；1.3 加幻灯片 + 版式/母版 + `Tag`；**1.4 形状的添加/移动/缩放/格式化/删除**；1.5 超链接；1.6 选择；1.7 自定义属性与文档属性；1.8 绑定/形状/**表格**；1.9 表格格式与管理；1.10 辅助功能/**幻灯片背景**/超链接 | [PowerPoint 要求集](https://learn.microsoft.com/zh-cn/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets) |

**Excel 形状的真实能力面（官方专题页）** —— 这是我方 MS 侧最大的**实现**缺口：

| 能做什么 | API | 我方是否实现 |
|---|---|---|
| 几何形状（全 `GeometricShapeType` 枚举） | `shapes.addGeometricShape(type)` | ❌ |
| 图片（JPEG/PNG base64） | `shapes.addImage(base64)` | ✅ |
| 线条 / 连接符（可连到形状的连接点） | `shapes.addLine(...)`、`line.connectBeginShape(shape, site)`、`connectionSiteCount` | ❌ |
| **SVG 矢量图** | `shapes.addSvg(xml)` | ❌ |
| 文本框 | `shapes.addTextBox(text)` | ❌ |
| 形状内文字与字体 | `shape.textFrame.textRange.text/font`、`textFrame.horizontalAlignment` | ❌ |
| 填充色 | `shape.fill.setSolidColor(color)` | ❌ |
| 旋转 | `shape.rotation` / `incrementRotation(deg)` | ❌ |
| 缩放与锁定比例 | `shape.scaleHeight/scaleWidth(type, from)`、`lockAspectRatio` | ❌ |
| 层级 | `shape.setZOrder(ShapeZOrder)` | ❌ |
| **分组** | `shapes.addGroup([...])`、`ShapeGroup.shapes` | ❌ |
| **形状导出为图片** | `shape.getAsImage(PictureFormat.png)` | ❌ |
| 读取"用户当前选中的形状" | `workbook.getActiveShape()` | ❌ |

来源：[在 Excel 加载项中创建和管理形状](https://learn.microsoft.com/zh-cn/office/dev/add-ins/excel/excel-add-ins-shapes)

### 2.2 VBA / COM 对象模型（Windows 原生驱动）

我方 Windows 原生驱动 `resources/office/excel.ps1`（251 行）**只覆盖 Excel 的 28 个结构化方法**（与加载项侧同名同形，但**少了 `update_chart` 与 `save_workbook`**），另加 `script` action 执行任意 PowerShell 脚本；Word/PPT 在 Windows 侧**只有 `export`（PPT 导页图）与 `script`**，没有任何结构化方法（源码可证，见 `runner.ps1` 第 27–41 行）。

> ⚠️ **白名单过度声明（本轮新发现，源码可证）**：`adapter.ts` 的 `fallbackToNative` 把 **`EXCEL_METHODS`（30 项）整体**当作"替代通道支持的方法白名单"传给 `routeOfficeFailure`（`errors.ts` 第 107 行）。但 COM 侧实际只有 28 项：
> - **`save_workbook`**：Office.js 有 `save`、COM **没有** → 在 Windows 上若 Office.js 保存失败且判定可回退，回退会落到 `Invoke-ExcelTool` 的 `default { throw "Unsupported Excel method: save_workbook" }`；
> - **`update_chart`**：Office.js 有、COM **没有** → 同样问题。
>
> 二者都属"白名单说有、实现没有"，会让回退路径在最需要的时候失效（好在是响亮报错，不是静默失败）。另注：`routeOfficeFailure` 第 106 行在非 Windows 平台**一律拒绝回退**，因此 macOS 上 Office.js 失败即硬失败——这是设计意图，不是缺陷。

从 COM 对象模型看，**VBA 是能力上限**，其中 Office.js 完全没有的高价值几类：

| 类别 | COM/VBA 能力 | Office.js |
|---|---|---|
| 文件级 | `Workbook.SaveAs` / `Open` / `Close` / `ExportAsFixedFormat`（PDF/XPS）/ `PrintOut` | **无** |
| 形状扩展 | `AddSmartArt` / `AddTextEffect`（艺术字）/ `AddCallout` / `AddOLEObject` / `BuildFreeform` / `Shape.ThreeD` / `Shadow` / `Glow` / `Reflection` / `SoftEdge` / `Adjustments` / `Nodes` | **无** |
| 宏与求值 | `Application.Run` / `Evaluate` / `ExecuteExcel4Macro` / `RegisterXLL` | **无** |
| 事件 | 45+ Application 级事件 + Worksheet/Workbook 级事件 | 只有 ExcelApi 1.17 的 3 个工作表事件 |
| Word | `MailMerge` / `PrintOut` / `Fields.Update` / 浮动 `Shape` / `ContentControls` | WordApi 有 `ContentControl`，**无** MailMerge/PrintOut |
| PPT | `PageSetup`（页面尺寸）/ 动画 / 切换 / `Export` / `SaveAs` | PowerPointApi **无**动画、切换、导出、页面尺寸 |

> ⚠️ 本机的 `resources/office/*.ps1` 是**我方实现**而非官方 API；上表的"官方能力"部分来自 Microsoft VBA 参考与要求集文档，属**文档可证**；但"我方驱动是否真的调用了这些"是**源码可证**（结论：基本没调用，只有 Excel 30 个方法）。

### 2.3 WPS JS API（加载项 JSAPI）

**关键结论：WPS 表格的对象模型是 Excel COM/VBA 的近乎全量克隆，能力面**远大于**Office.js。**

**`wps` 全局对象**（[wps 对象成员](https://qn.cache.wpscdn.cn/encs/doc/office_v5/topics/WPS%20%e5%8a%a0%e8%bd%bd%e9%a1%b9%e5%bc%80%e5%8f%91/%e5%8a%a0%e8%bd%bd%e9%a1%b9%20API%20%e5%8f%82%e8%80%83/Office%20%e5%85%a8%e5%b1%80%e5%af%b9%e8%b1%a1/wps%20%e5%af%b9%e8%b1%a1%e6%88%90%e5%91%98.htm)）：

- 方法：`EtApplication()`（表格）/ `WpsApplication()`（文字）/ `WppApplication()`（演示）、`CreateTaskpane` / `GetTaskpane` / `ShowDialog` / `alert` / `confirm`
- 属性：**`ApiEvent`（应用程序事件对象）**、`Env`、**`FileSystem`**、**`PluginStorage`（加载项持久化存储）**

**WPS 表格 `Application`**（[Application 对象成员](https://qn.cache.wpscdn.cn/encs/doc/office_v5/topics/WPS%20%e5%8a%a0%e8%bd%bd%e9%a1%b9%e5%bc%80%e5%8f%91/%e8%a1%a8%e6%a0%bc%20API%20%e5%8f%82%e8%80%83/Application/Application%20%e5%af%b9%e8%b1%a1%e6%88%90%e5%91%98.htm)）：约 50 个方法（`Calculate` / `CalculateFull` / `CalculateFullRebuild` / `Union` / `Intersect` / `Evaluate` / **`Run`（运行宏）** / `Undo` / `SendKeys` / `Goto` / `Quit` / `OnTime` / `OnKey` / `ConvertFormula` / `RegisterXLL` / `ExecuteExcel4Macro` / `SaveWorkspace` …）+ 约 150 个属性（**`Names`（命名区域）** / `Workbooks` / `Sheets` / `Windows` / `WorksheetFunction` / `Selection` / `ActiveChart` / `RecentFiles` / `VBE` / `CommandBars` / `SmartArtLayouts` / `SmartArtColors` …）+ **约 45 个事件**（`SheetChange` / `SheetSelectionChange` / `SheetCalculate` / `WorkbookOpen` / `WorkbookBeforeSave` / `WorkbookAfterSave` / `WorkbookNewSheet` / `WorkbookNewChart` / `WindowActivate` / `SheetBeforeDoubleClick` / `SheetBeforeRightClick` / `SheetFollowHyperlink` / 6 个数据透视表事件 …）。

**WPS 表格 `Shapes`**（[Shapes 对象成员](https://qn.cache.wpscdn.cn/encs/doc/office_v5/topics/WPS%20%e5%8a%a0%e8%bd%bd%e9%a1%b9%e5%bc%80%e5%8f%91/%e8%a1%a8%e6%a0%bc%20API%20%e5%8f%82%e8%80%83/Shapes/Shapes%20%e5%af%b9%e8%b1%a1%e6%88%90%e5%91%98.htm)）：`AddShape` / `AddTextbox` / `AddLine` / `AddConnector` / `AddCurve` / `AddPolyline` / `AddCallout` / `AddLabel` / `AddPicture` / **`AddTextEffect`（艺术字）** / **`AddSmartArt`** / **`BuildFreeform`（任意多边形）** / `AddFormControl` / `AddOLEObject` / `SelectAll`。

**WPS 表格 `Shape`**（[Shape 对象成员](https://qn.cache.wpscdn.cn/encs/doc/office_v5/topics/WPS%20%e5%8a%a0%e8%bd%bd%e9%a1%b9%e5%bc%80%e5%8f%91/%e8%a1%a8%e6%a0%bc%20API%20%e5%8f%82%e8%80%83/Shape/Shape%20%e5%af%b9%e8%b1%a1%e6%88%90%e5%91%98.htm)）：约 18 个方法（`PickUp` + `Apply`（**格式刷**）/ `Flip` / `Duplicate` / `ZOrder` / `ScaleHeight` / `ScaleWidth` / `Ungroup` / `RerouteConnections` / `IncrementRotation` / `SetShapesDefaultProperties` …）+ 约 55 个属性（**`ThreeD` / `Shadow` / `Glow` / `Reflection` / `SoftEdge` / `PictureFormat` / `Callout` / `TextEffect` / `SmartArt` / `Adjustments` / `Nodes` / `Vertices` / `ConnectorFormat` / `Hyperlink` / `OnAction` / `LockAspectRatio` / `Placement` / `ShapeStyle` / `TextFrame` / `TextFrame2` / `GroupItems`** …）。

**WPS 表格条件格式**：`FormatConditions` 有 `AddIconSetCondition`（[示例](https://qn.cache.wpscdn.cn/encs/doc/office_v5/topics/WPS%20%e5%8a%a0%e8%bd%bd%e9%a1%b9%e5%bc%80%e5%8f%91/%e8%a1%a8%e6%a0%bc%20API%20%e5%8f%82%e8%80%83/FormatConditions/%e6%96%b9%e6%b3%95/AddIconSetCondition%20%e6%96%b9%e6%b3%95.htm)），另有 `Top10` 等对象。

**WPS 文字 / 演示 API 参考**同样存在于官方文档站（`文字 API 参考` / `演示 API 参考` 目录），对象模型对齐 VBA Word / PowerPoint。

### 2.4 一句话结论

> **不是"WPS 比 MS 弱"，而是"WPS 宿主 API 面（COM 克隆）≫ Office.js 面，但我们的 WPS 加载项只用了很小一撮；MS 侧 Office.js 面本身较窄，而我们的 Office.js 加载项又只用了一小撮中的一小撮。"**
>
> - WPS 宿主：**能力最全**（含 SmartArt、艺术字、三维、事件、宏调用、文件级操作），我们的加载项**只暴露了 30 个表格方法 + 13 个文字 + 9 个演示**。
> - MS Office.js：**能力较窄**（无文件级、无 SmartArt/艺术字/OLE、无动画切换、无 PPT 图表），但**我们有 43 个处理器只覆盖 Excel**，且**形状域只实现了 `addImage`**。
> - MS COM/JXA：只在 Windows 提供 Excel 结构化方法；Word/PPT **只能靠任意脚本**。

---

## 3. 四类能力差集

> **重要的口径纠正**：本轮查证后必须区分三种"差"——
> **(a) 宿主就不支持**；(b) **宿主支持、我们的加载项没实现**；(c) **两边加载项都没实现**。
> 上一轮容易把 (b) 误当成 (a)。本节每一条都标注属于哪一种。

### 3.1 ① MS 有、WPS 没有（或 WPS 明显弱）

| # | 能力 | 类型 | 依据 |
|---|---|---|---|
| ①-1 | **平台覆盖**：Office.js 加载项同一份代码可跑 Excel on the web / Windows / Mac / iPad；WPS 加载项只能在 WPS 客户端内运行 | (a) 宿主 | 文档可证（[Excel 要求集表](https://learn.microsoft.com/zh-cn/javascript/api/requirement-sets/excel/excel-api-requirement-sets)逐平台列出支持情况） |
| ①-2 | **`RangeAreas` 非连续区域一等类型**（一次对多个不连续矩形读值/设格式） | (a) 宿主（**推测**） | Office.js 文档可证有 `RangeAreas`（[ExcelApi 1.9](https://learn.microsoft.com/zh-cn/javascript/api/requirement-sets/excel/excel-api-1-9-requirement-set)）；WPS 官方参考中**只检索到 `Application.Union`/`Range.Areas`，未检索到 RangeAreas 类型** → 标**推测**，需实测 |
| ①-3 | **SVG 矢量图插入**（`shapes.addSvg(xml)`） | (a) 宿主（**推测**） | Office.js 文档可证有；WPS `Shapes` 方法清单中**无 AddSvg** |
| ①-4 | **现代图表类型**：地图 / 箱形图 / 瀑布图 / 旭日图 / 排列图 / 漏斗图 | (a) 宿主（**推测**） | Office.js `Excel.ChartType` 文档可证；WPS 走 `XlChartType` 常量，官方枚举页未检索到这些新类型 → 标**推测** |
| ①-5 | **图表细粒度**：`ChartSeries` 误差线 / 趋势线 / 箱形图选项 / 渐变；`ChartAxis.linkNumberFormat`；`ChartDataLabels.linkNumberFormat`；`ChartAreaFormat.colorScheme`/`roundedCorners`；`Chart.pivotOptions`；`Chart.activate()` | (a) 宿主（**推测**） | Office.js 文档可证；WPS `Chart`/`Series` 是否有对应属性**未查证** → 标**推测** |
| ①-6 | **`DataValidation.getInvalidCells()`**：直接列出"当前哪些单元格的值不合法" | (a) 宿主（**推测**） | Office.js 文档可证；WPS 官方参考未检索到对应方法 |
| ①-7 | **批量读回单元格属性 `Range.getCellProperties`**（一次取回多格的字体/填充/边框/对齐/保护/缩进等） | (c) 粒度 | Office.js 文档可证（`CellProperties` 系列，ExcelApi 1.9）；WPS 需逐格读属性 → 语义差异而非能力缺失 |
| ①-8 | **批处理事务语义 `Excel.run` + `context.sync()`**：多次属性写入合并为一次宿主往返；`Application.suspendScreenUpdatingUntilNextSync()` 可整段关闭屏幕刷新 | (a) 宿主（**推测**） | Office.js 文档可证；WPS 是浏览器内核直连 COM，**未检索到 batch/sync 概念**，每次属性赋值即时生效 → 标**推测**，但 WPS 有 `Application.ScreenUpdating` 可部分替代 |
| ①-9 | **冻结窗格的对象化 API**：`freezePanes.freezeAt(cell)` / `freezeRows(n)` / `freezeColumns(n)`，**不需要激活工作表** | (b) 我方实现差异 | **源码可证**：Office.js `handleFreezePanes` 用 `sheet.freezePanes.*`；WPS `freezePanes` 必须 `sheet.Activate()` + `app.ActiveWindow.SplitRow/SplitColumn` → **WPS 在多窗口/非活动簿上会打错目标** |
| ①-10 | **合并单元格的 `across` 语义**：`range.merge(across)` 支持"按行合并" | (b) 我方实现差异 | 源码可证：Office.js 传 `params.across`；WPS 只调 `range.Merge()` |
| ①-11 | **`update_chart`（改标题 / 图例位置 / 位置尺寸）** | (b) 我方实现缺口 | 源码可证 + `HOST_IMPLEMENTATION_GAPS.wps = ['update_chart']` |

### 3.2 ② WPS 有、MS 没有（我们是否只在一边用得上）

| # | 能力 | 类型 | 我们是否只在一边用得上 | 依据 |
|---|---|---|---|---|
| ②-1 | **Word 全部 13 个结构化工具** | (a) 宿主（我方 MS 侧无 Office.js 通道） | **是**：`wps_word_*` 是 WPS 专用；MS Word 只能 `office_execute_script` | 源码可证（`office-addon/excel/manifest.xml` 只有 Excel；`capabilities().hosts.microsoft.word`） |
| ②-2 | **PPT 全部 9 个结构化工具** | 同上 | **是**：MS PowerPoint 只能 `office_execute_script` | 同上 |
| ②-3 | **PowerPoint 原生图表**（`ppt_insert_native_chart`） | (a) 宿主 | **是**：PowerPointApi 1.1–1.10 官方文档中**没有 Chart 类**，MS 侧即使写 Office.js 也做不了原生图表；只能走 COM/JXA 脚本 | 文档可证（[PowerPoint 要求集](https://learn.microsoft.com/zh-cn/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets) 逐版本能力清单无图表） |
| ②-4 | **SmartArt 创建**（`Shapes.AddSmartArt`） | (a) 宿主 | 否（两边都能做，只是 MS 要绕道 COM） | 文档可证 WPS 有；Office.js 形状 API 清单无 SmartArt |
| ②-5 | **艺术字 TextEffect**（`Shapes.AddTextEffect`） | (a) 宿主 | 否（同上） | 同上 |
| ②-6 | **标注 Callout / 标签 Label / 表单控件 FormControl / OLE 对象 / 任意多边形 BuildFreeform / 贝塞尔曲线 AddCurve / 折线 AddPolyline** | (a) 宿主 | 否（同上） | 同上 |
| ②-7 | **形状的三维 / 阴影 / 发光 / 映像 / 柔化边缘 / 调整值 / 节点编辑**（`ThreeD`/`Shadow`/`Glow`/`Reflection`/`SoftEdge`/`Adjustments`/`Nodes`/`Vertices`） | (a) 宿主 | 否（同上） | 文档可证 WPS `Shape` 有；Office.js `Shape` 无 |
| ②-8 | **格式刷**（`Shape.PickUp()` + `Apply()`） | (a) 宿主 | 否 | 同上 |
| ②-9 | **宏调用与求值**（`Application.Run` / `Evaluate` / `ExecuteExcel4Macro` / `RegisterXLL`） | (a) 宿主 | 部分：我们已有 `wps_execute_script`（等价逃生舱），`Application.Run` 反而多余 | 文档可证 WPS 有；Office.js 无 |
| ②-10 | **文件级操作**：`SaveAs` / `ExportAsFixedFormat`（PDF/XPS）/ `Open` / `Close` / `PrintOut` | (a) 宿主 | **是**：Excel 导出 PDF 这种高频需求，**加载项层两边都不完整**（MS 侧完全没有；WPS 侧加载项也没实现，只有 Word 侧做了 `word_save_document` 的 PDF） | 文档可证（[Word 侧我方源码](../../../../../wps-addon/src/word.js) `doc.ExportAsFixedFormat(filePath, 17)` 已用；WPS 表格侧未实现） |
| ②-11 | **约 45 个 Application 级事件**（`SheetChange` / `WorkbookBeforeSave` / `SheetSelectionChange` …）+ `wps.ApiEvent` | (a) 宿主 | **是**：Office.js 只有 ExcelApi 1.17 的 3 个工作表事件（改名/可见性/移动），**没有单元格变更事件** | 文档可证（WPS Application 事件表；Office.js 1.17 列表） |
| ②-12 | **加载项文件系统访问**（`wps.FileSystem`）与 **加载项持久化存储**（`wps.PluginStorage`） | (a) 宿主 | 部分：我们是"网关侧落盘"，不需要加载项写文件；但 `PluginStorage` 可用于跨会话记住目标文档 | 文档可证 WPS 有 `FileSystem` / `PluginStorage` |
| ②-13 | **`get_style_token`**（把工作簿主题/样式令牌交给 AI 当调色板）、**`rollback_cells`**（宿主侧快照回填）、**`inspect_api`**（任意表达式反射探测） | (b) 我方实现独有 | **是**：这三项在 Office.js 侧没有对应分支。其中 `inspect_api` 在 MS 侧无替代物（只有 `office_execute_script`） | 源码可证 |
| ②-14 | **透视表字段编排**（`PivotFields().Orientation` + `AddDataField`） | (b) 我方实现差异 | **是**：WPS 侧能排行/列/数据字段；Office.js 侧我们只 `add(name, source, dest)` + `refresh` | 源码可证 |
| ②-15 | **工作表保护 / 标签色**（`manage_sheet` 的 `protect`/`unprotect`/`tab_color`） | (b) 我方实现缺口 | **是**：Office.js `handleManageSheet` 无这两个分支（虽然 Office.js 有 `worksheet.protection` 与 `tabColor`） | 源码可证 |
| ②-16 | **`add_chart` 的 `dataRanges` 非连续联合数据源 / `yAxis` 坐标轴细调 / `seriesSettings` 逐系列设置** | (b) 我方实现缺口 | **是**：Office.js 侧 `handleCreateChart` 无这些参数（虽然 Office.js 有 `ChartAxis`/`ChartSeries` 可做） | 源码可证 + 文档可证 |

### 3.3 ③ 两边都有但粒度/语义不同

| # | 主题 | WPS 侧语义 | Office.js 侧语义 | 依据 |
|---|---|---|---|---|
| ③-1 | **图表定位** | `Shapes.AddChart2(type, left, top, w, h)` 用**磅值像素**；`position.leftCell` 只是把该单元格的 `Left/Top` 取出来当像素初值 | `chart.setPosition(startCell, endCell)` 是**真正的单元格锚定**（跟随行列尺寸），失败才退化到像素 | **源码可证**（`wps-addon/src/excel.js` addChart 第 1147–1170 行；`office-addon/src/excel/chart.js` applyChartPosition） |
| ③-2 | **`capture_sheet_preview` 的真实性** | `Range.CopyPicture` → 临时 `ChartObject` → `Chart.Export` PNG，**像素级真实工作表渲染** | 有图表时 `chart.getImage()`（真图）；**无图表时用 Canvas 按 `cellW=110 / rowH=28` 硬编码合成一张"看起来像表格"的图** | **源码可证**（`office-addon/src/excel/chart.js` `handleCaptureSheetPreview` 第 2 段）。⚠️ 这意味着 **MS 侧"视觉自检"会给出与真实文件不一致的结论** |
| ③-3 | **条件格式的规则类型与配色** | `cell_value`（less/greater/equal/between）+ `data_bar`（只设 `BarColor.Color`）+ `color_scale`（**可设两端色**） | `cellValue`（含 **not_equal**）+ `colorScale`（**只 `add`，不设任何颜色**）+ `dataBar`（可设正格式填充色/边框色/渐变开关） | **源码可证**（`wps-addon/src/excel.js` addConditionalFormatting；`office-addon/src/excel/format.js` handleAddConditionalFormatting） |
| ③-4 | **合并单元格的写后校验** | 合并后**读回 `MergeCells` 并比对**，不符即抛错（ISS-42 修复） | `range.merge()` 后**不读回** | 源码可证 |
| ③-5 | **隐藏行/列** | 逐行 `Rows.Item(n).Hidden` 写 + **逐行读回校验**（ISS-60 修复），列走 `Columns.Item` | 走 `update_range_structure` 整块 `Range` 操作 | 源码可证 |
| ③-6 | **批注 vs 备注** | `Range.Comment`（**传统备注 Note**），有 `Author` / `Text()` | `sheet.comments`（**现代线程化批注 Comment**），有 `id` / `authorName` / `creationDate`；MS 的备注是另一个对象 `Range.note`（**我们没实现**） | 文档可证 + 源码可证。⚠️ 同一句"给 A1 加批注"在两侧落到**不同对象**，跨宿主语义不等价 |
| ③-7 | **`update_comment` 名义存在但无实现** | `manage_cell_comments` 的 action 有 `add` / `read` / `delete` / `clear_all`（**含清空全部**，无 update） | `update_comment` 别名落进 `handleManageComments`，但**没有任何 action 分支匹配 → 静默返回 `{success:true}`** | 源码可证 |
| ③-8 | **`list_conditional_formats` / `update_conditional_format` 名义存在但无实现** | 无对应方法 | 两个别名都落进 `handleAddConditionalFormatting`，**只做 add/clearAll → 调用"列表"会改成新增规则、调用"更新"同样只新增** | 源码可证。属"别名齐全、语义错位"的**危险**差异 |
| ③-9 | **命名区域** | 宿主有 `Workbook.Names`（文档可证），但**我方加载项没实现** | 我方实现了 `list_named_items` / `update_named_item`（`workbook.names`，含 `type` / `visible`） | 文档可证 + 源码可证 |
| ③-10 | **文档属性** | 宿主有（VBA 对齐），我方加载项**没实现** | 我方实现了 `get/update_document_properties`（8 个内置字段） | 同上 |
| ③-11 | **结构化表格 Table（ListObject）** | 宿主大概率有（VBA 对齐），我方加载项**没实现** → 排查性问题 | 我方实现了 `create_table` / `update_table`（含 `style` / `showTotals` / `resize`） | 源码可证 |
| ③-12 | **区域复制** | 未实现 | `copy_range` 已实现 | 源码可证 |
| ③-13 | **超链接** | 未实现（宿主有 `Hyperlink` 对象，Shape 也有 `Hyperlink` 属性） | `set_hyperlink` 已实现 | 源码可证 + 文档可证 |
| ③-14 | **图片与形状管理** | 表格侧**完全没实现**（宿主 `Shapes` 全套能力闲置） | 只实现了 `addImage` + `list` + 移动/缩放/删除 | 源码可证 |
| ③-15 | **任意脚本沙箱** | `AsyncFunction("app","doc","wb","pres","wps","params","console")` **七参注入** + `safeSerialize` 深度 2 防循环引用 + 捕获 `console` 输出 | `new Function("context","Excel", script)` **两参注入**，无 console 捕获、无返回值序列化保护 | 源码可证 |
| ③-16 | **筛选** | `setFilterAndSort` 只做 `AutoFilterMode` 开关 + `Sort` 多规则（**没有列级筛选条件**） | `autoFilter.apply(range, col, {filterOn: values, values})` 能按列 + 值列表筛选（**只支持 `values` 一种 `FilterOn`**） | 源码可证 |
| ③-17 | **排序** | `Sort.SortFields.Add(col, SortOn, Order)`（`Order` 只有 1/2 = 升/降，无自定义序列） | `range.sort.apply([{key, ascending, sortOn}], hasHeader)` | 源码可证 |
| ③-18 | **数据验证** | `Validation.Add(type, alertStyle, operator, f1, f2)`，读 `validationType` / `listItems` / `operator` / `minVal` / `maxVal` / 提示与报错文案 —— **与 MCP schema 完全对齐** | `range.dataValidation.rule = params.rule` —— **只认一个原始的 `Excel.DataValidationRule` 对象，Schema 里的 9 个字段一个都不读** | **源码可证**，见 §3.5-M1 |
| ③-19 | **冻结窗格目标** | 依赖 `ActiveWindow`，必须先 `Activate` 工作表 → **跨簿/多窗口不安全** | 直接对 `sheet.freezePanes` 操作，**与活动窗口无关** | 源码可证 |
| ③-20 | **Word 分节与页面设置** | `word_page_layout_and_watermark` 直接操作节与 `PageSetup`，**水印是一等能力** | WordApi 有 `Section`，但**没有 `Watermark` API** → MS 侧水印要靠 OOXML/页眉形状 | 源码可证 + 文档可证 |
| ③-21 | **Word 修订** | `enable/disable_track_changes` + `accept_all_revisions` / `reject_all_revisions`（**只有全量，没有逐条**） | WordApi 有 `changeTrackingMode` 与 `Comment` 对象（可 reply / resolved） | 文档可证 + 源码可证 |
| ③-22 | **PDF 导出** | 仅 **Word** 侧有（`doc.ExportAsFixedFormat(filePath, 17)`）；**Excel / PPT 侧没有** | **加载项层完全没有**（只有 Windows COM `script` 逃生舱） | 源码可证 |

### 3.4 ④ 两边都没暴露的高价值能力（对当前 MCP 工具面的空白）

| # | 空白能力 | 宿主是否支持 | 依据等级 | 对 AI 的影响 |
|---|---|---|---|---|
| ④-1 | **Excel 分页与打印设置**：打印区域、打印标题行列、纸张、方向、页边距、页眉页脚、网格线/标题/批注打印、手动分页符 | **两边都支持**（Office.js `PageLayout` + `PageBreakCollection`；WPS VBA 对齐的 `PageSetup`） | 文档可证（MS 侧）/ **推测**（WPS 侧未逐条核对） | AI 做不出"可直接打印/装订"的报表 |
| ④-2 | **Excel 导出 PDF / XPS** | WPS 支持（`ExportAsFixedFormat`）；Office.js **不支持**，需 COM/JXA | 文档可证 | 交付链路断在最后一步，"生成报表"变成"生成 xlsx 请你自己导" |
| ④-3 | **MS 侧矢量绘图**（几何形状 / 连接符 / SVG / 文本框 / 分组 / 层级 / 旋转 / 缩放 / 形状导图） | **支持**（ExcelApi 1.9 `Shape` 全套） | 文档可证 | MS 用户完全用不上 P5 里那套"134 个矢量元素信息图"，只能靠图片贴图 |
| ④-4 | **WPS 侧矢量绘图与 SmartArt**（`Shapes.AddShape` / `AddTextEffect` / `AddSmartArt` / `BuildFreeform` / `ThreeD` / `Shadow` / 格式刷） | **支持**（宿主对象模型） | 文档可证 | WPS 侧明明能画，我们却只会 `add_chart` + `addImage` 缺失 |
| ④-5 | **单元格内局部格式（富文本）**：一句话里只把某段加粗/变红，或部分字符改字号 | WPS 有 `Range.Characters`（VBA 对齐，文档可证宿主有）；Office.js 侧需 `Range.format` 整格或 OOXML | 文档可证（WPS）/ **推测**（Office.js 路径） | 这是"精细操控文档"的典型诉求，目前只能整格统一 |
| ④-6 | **数据透视表的字段编排与刷新（MS 侧）** | 支持（Office.js `PivotTable` 可加行/列/数据字段；WPS `PivotFields`） | 文档可证 | MS 侧"建透视表"只能建空壳 |
| ④-7 | **条件格式图标集（IconSet）与完整规则族**：红黄绿灯、Top10、重复值、文本包含、公式规则、规则列表与更新 | **两边都支持**（WPS `AddIconSetCondition` / `Top10`；Office.js 1.17 `changeRuleToIconSet` 等） | 文档可证 | 报表缺"红黄绿灯"这类最常见的业务信号 |
| ④-8 | **Word 内容控件全类型**：日期、复选框、下拉、组合框、富文本、纯文本，含 `listItems` | 支持（WordApi 1.9 下拉/组合框；VBA `ContentControls` 全类型） | 文档可证 | AI 做不出"可填模板/合同填空位/表单" |
| ④-9 | **Word 域的更新与交叉引用**（`Fields.Update` / `Field.update()`）：目录页码、交叉引用、页数 | **两边都支持** | 文档可证（WordApi `Field`/`Fields`；WPS VBA 对齐） | 我们只做了"插入目录"，目录页码不会随内容更新 → 交付物一改就错 |
| ④-10 | **Excel 线程化批注的完整语义**：回复、@提及、标记已解决 | 支持（Office.js `Comment` 有 reply/resolved 语义） | 文档可证 | AI 只能"贴一条"，不能参与讨论 |
| ④-11 | **Excel 切片器 / 工作表视图 / 链接数据类型（富值）** | MS 支持（ExcelApi 1.9+ / 较新要求集）；WPS **未查证** | 文档可证（MS）/ **未查证**（WPS） | 交互式仪表盘与"从数据源自动带出公司/股票信息" |
| ④-12 | **PPT 页面尺寸与母版/版式控制、导出**（`PageSetup` / `SlideMaster` / `SlideLayout` / 导出图片或 PDF） | MS 支持 `SlideMaster`/`SlideLayout`（PowerPointApi 1.3，**无页面尺寸、无导出**）；WPS 有 `PageSetup`（文档可证宿主） | 文档可证 | 无法把 4:3 改成 16:9、无法套用公司母版、无法批量导图 |
| ④-13 | **事件驱动的"变更感知"**：用户刚改了哪个单元格、保存前拦一手 | WPS 支持（45+ Application 事件 + `wps.ApiEvent`）；Office.js 只有改名/可见性/移动 3 个事件 | 文档可证 | AI 只能"盲写"，无法基于用户刚做的修改做增量协作 |
| ④-14 | **数据验证的读回与违规定位** | MS 支持（`getInvalidCells`）；WPS 支持读 `Validation` 属性 | 文档可证（MS）/ **推测**（WPS） | AI 写完下拉/范围校验后无法自检"哪些值越界" |
| ④-15 | **命名区域（WPS 侧）/ 文档属性（WPS 侧）/ 结构化表格（WPS 侧）/ 区域复制 / 超链接 / 图片 / 图表更新** | **WPS 宿主都支持** | 文档可证（宿主）；源码可证（我方未实现） | WPS 用户在 MS 侧能用的 7 项能力，回到 WPS 就消失 |

### 3.5 跨宿主参数契约不一致（本轮新发现，全部**源码可证**）

> 这一节是 ③ 类里**最危险**的一类：**MCP 工具 schema、gateway 转发、normalizer 转换三层都是按 WPS 语义写的；Office.js 侧读的是另一套字段名**。结果是同一个 `excel_*` 工具在 `host:'microsoft'` 下**要么静默 no-op，要么报错，要么给出错误结果**。
>
> 三层的证据链：`src/bridge/tools/definitions/excel.ts`（schema 字段名）→ `src/bridge/gateway/excel.ts`（原样转发）→ `src/bridge/office/normalizer.ts`（`default:` 分支只去前缀、不改字段）→ `office-addon/src/excel/*.js`（读的字段名）。

| # | 工具 | Schema 实际发送 | Office.js 实际读取 | 后果 | 等级 |
|---|---|---|---|---|---|
| **M1** | `excel_set_data_validation` | `validationType` / `listItems` / `operator` / `minVal` / `maxVal` / `prompt*` / `error*` | **只有 `params.rule`** | `rule` 永远是 `undefined` → 先 `dataValidation.clear()` 再什么都不设 → **返回 `{success:true}`，下拉/校验静默消失**（还会**清掉区域上原有的校验**） | ⚠️ **静默 no-op + 破坏性** |
| **M2** | `excel_find_and_replace` | `searchQuery` / `replaceText` / `searchRange` / `maxResults` | `params.text \|\| params.query \|\| params.findText \|\| ""`（**不读 `searchQuery`**） | 搜索路径：`queryStr=""` → `String.includes("")` 恒为 `true` → **每个非空单元格都报成命中**；替换路径：`range.replaceAll("", replaceText, …)` → **搜索串为空却调用全局替换，属破坏性调用** | ⚠️ **结果错误 + 破坏性** |
| **M3** | `excel_create_pivot_table` | `sourceRange` / `destCell` / `sourceSheetName` / `destSheetName` / `rowFields` / `columnFields` / `dataFields` | `params.sourceAddress` / `params.destinationAddress` / `params.name` | `sheet.getRange(undefined)` → **报错**（好在是"响亮失败"）；且即便修正字段名，MS 侧也没有**字段编排**能力 | 报错（安全） |
| **M4** | `excel_set_filter_and_sort` | `range` / `enableAutoFilter` / `sortRules` | `params.address` / `params.filterColumn` / `params.criteria` / `params.sortColumn` / `params.hasHeader` | `sheet.getRange(undefined)` → **报错**（安全） | 报错（安全） |
| **M5** | `excel_manage_cell_comments` | `action ∈ {add, read, delete, clear_all}` | 只识别 `add` / `list` / `read` / `delete` | 恰好 schema 用的是 `read` 而不是 `list` → **读批注能通**；但 `clear_all` 走到末尾 `return {success:true}` → **清空全部批注会静默失败**（且 MS 侧也没有 `clear_all` 实现） | ⚠️ **静默 no-op** |
| **M6** | `excel_manage_sheet` | `action ∈ {rename, move, tab_color, protect, unprotect}` + `color` / `password` / `targetIndex` | `handleManageSheet` 只匹配 `copy` / `duplicate` / `rename` / `hide` / `show` / `color`（且读的是 `params.tabColor`，不是 `color`） | **四个 action 全部走到末尾 `return {success:true}`**：`move`、`tab_color`、`protect`、`unprotect` 在 MS 侧**静默 no-op**（只有 `rename` 真通） | ⚠️ **静默 no-op ×4** |
| **M7** | `excel_get_charts` | `shapeName` / `chartIndex` / `chartTitle` / `detail` | 只读 `sheetName` + `detail`，**忽略全部选择器** | 传了 `chartTitle` 也会**返回该表全部图表** → AI 以为筛过了 | ⚠️ **语义静默降级** |
| **M8** | `excel_add_chart` | 含 `yAxis` / `seriesSettings` / `smoothLine` / `dataRanges` / `hasDataLabels` | 只读 `seriesBy` / `seriesColors` / `hasLegend` / `title` / `position.*` + `setPosition` 锚定 | 5 个 WPS 专属参数被**静默忽略**（不报错）→ AI 以为轴范围、逐系列配色、数据标签都设上了 | ⚠️ **静默降级** |
| **M9** | `excel_update_chart` | 完整 schema 存在 | Office.js **已实现**；WPS **未实现** | 在 `host:'wps'` 下必然抛"未知的 RPC 方法"（已由 `HOST_IMPLEMENTATION_GAPS.wps` 声明） | 已在契约中声明 |

> **反面对照（这三条是通的，可作为改造模板）**：`excel_search_cells` 在 normalizer 里有 `case 'search_cells'` → 把 `query` 映射成 `text`，**MS 侧可用**；`excel_format_cells` 有一条极厚的映射分支（同时构造 `font` / `fill` / `alignment` 嵌套对象与扁平字段）；`excel_add_chart` 的 `CHART_TYPE_MAP` 把 `column` → `ColumnClustered`。**M1–M8 缺的正是这类逐字段映射。**

> **口径提醒**：M1–M8 是**我方实现的跨宿主不一致**，不是宿主能力缺陷。WPS 宿主本身提供 `Validation.Add`、`FormatConditions`、`ListObjects`、`Tab.Color`、`Protect` 等全部对象（§2.3 文档可证）。

---

## 4. 补强建议（按价值排序，共 10 条）

> 排序原则：**先修"会让 AI 相信了假结果"的正确性问题，再补新能力**。
> 每条标注：补在哪一侧 / 大概怎么补 / 补了以后 AI 能多做什么 / 依据等级。

---

### S1. 【桥接网关 + Office.js 加载项】修 M1–M8 的跨宿主字段契约（§3.5）

- **怎么补**：在 `src/bridge/office/normalizer.ts` 为 `set_data_validation` / `create_pivot_table` / `set_filter_and_sort` / `find_and_replace` / `manage_sheet` / `manage_cell_comments` 各加一条**显式映射分支**（照抄现有 `format_cells`、`search_cells`、`add_chart` 的样板）；MS 侧 `handleSetDataValidation` 补"用 schema 字段构造 `Excel.DataValidationRule`"；`handleManageSheet` 补 `tab_color`（读 `params.color`）/ `protect` / `unprotect` / `move` 四个 action。
- **顺带修两处台账**：① 把 `routeOfficeFailure` 的 `supportedMethods` 从"整张 `EXCEL_METHODS`"换成"**COM 侧真实实现的方法集**"（去掉 `update_chart`、`save_workbook`），避免白名单说能回退、实际抛出 `Unsupported Excel method`（§2.2）；② 给 `rollback_cells` 在 Office.js 侧补分支，或在 `HOST_IMPLEMENTATION_GAPS.microsoft` 里如实登记。
- **AI 能多做什么**：`host:'microsoft'` 下的下拉校验、筛选排序、透视表、批量替换、工作表保护**才会真的生效**，而不是"返回 success、文件里什么都没有"。
- **依据**：源码可证（三层证据链 + 覆盖率核对脚本已列出）。
- **立即止血**：若短期不改，至少让 MS 侧**未知 action 抛错**而不是 `return {success:true}`。

### S2. 【Office.js 加载项】补全 Excel 形状与矢量绘图域

- **怎么补**：把 `office-addon/src/excel/shape.js` 从 3 个处理器扩成一个 `shape` 能力域：`insert_shape`（`shapes.addGeometricShape` / `addTextBox` / `addSvg`）、`insert_connector`（`shapes.addLine` + `line.connectBeginShape/connectEndShape`）、`group_shapes`（`shapes.addGroup`）、`set_shape_z_order`、`update_shape` 扩到 `rotation` / `scaleHeight|Width` / `lockAspectRatio` / `textFrame.textRange`（文字 + 字体 + 对齐）/ `fill.setSolidColor`、`export_shape_image`（`shape.getAsImage`）、`get_active_shape`（`workbook.getActiveShape`）。
- **AI 能多做什么**：**Microsoft Excel 上也能画出 P5 里那种"134 个矢量元素信息图"**——流程图、架构图、带连接符的关系图、带文字的标注块，而不只是贴一张 base64 图片。这是"精细操控文档"最直接的增益，且 WPS 侧已有对应能力，可做到两侧对齐。
- **依据**：**文档可证**（[Excel 形状专题页](https://learn.microsoft.com/zh-cn/office/dev/add-ins/excel/excel-add-ins-shapes)明确列出全部方法）+ 源码可证（我方只实现了 `addImage`）。

### S3. 【Office.js 加载项】把 `capture_sheet_preview` 从"伪渲染"改成真实渲染

- **现状**：MS 侧无图表时用 Canvas 按 `cellW=110 / rowH=28` **硬编码合成一张图**（源码可证）。
- **怎么补**：三条路（推荐第 1 条）——
  1. 无图表目标时**改走 Windows COM**：`resources/office/excel.ps1` 的 `capture_sheet_preview` **已经用 `Range.CopyPicture` + 临时 `ChartObject` + `Chart.Export` 做真实渲染**，只需在 `capture_sheet_preview` 分支允许这条回退（它属于"不改文档内容"的可重放方法，已在 `REPLAY_SAFE_METHODS` 中）。
  2. macOS 无 COM，退化为"数据摘要 + `chart.getImage()`"，并在返回体里显式标 `synthetic: true` / `renderMode: 'synthetic'`。
  3. 用 `sheet.getUsedRange()` + 逐行 `format` 读回真实值，至少不做像素级伪装。
- **AI 能多做什么**：视觉自检结论与真实文件一致——不会再出现"AI 看图说排版没问题，用户打开发现全错"。
- **依据**：源码可证。**这是本轮最该优先修的正确性问题之一**。

### S4. 【WPS 加载项 + Office.js 加载项】补"分页 / 打印 / 导出 PDF"能力域

- **怎么补**：
  - **Office.js 侧**：`worksheet.pageLayout`（`setPrintArea` / `setPrintTitleRows|Columns` / `orientation` / `paperSize` / `setPrintMargins` / `headersFooters` / `printGridlines` / `printHeadings` / `printComments` / `zoom`）+ `worksheet.pageBreaks`（`add` / `removePageBreaks`）→ 新增 `excel_page_layout` 工具。
  - **WPS 侧**：`PageSetup`（VBA 对齐，文档可证宿主有）+ `Workbook.ExportAsFixedFormat`（PDF/XPS）→ 新增 `excel_export_pdf`；加载项层已有 Word 侧 PDF 样板可复用。
  - **网关**：`excel_export_pdf` 在 `host:'microsoft'` 时路由到 COM/JXA 脚本路径。
- **AI 能多做什么**：从"生成一个 xlsx 请你自己排版打印"升级为**交付可直接打印/装订的报表与 PDF**——这是办公场景最高频的最后一公里。
- **依据**：**文档可证**（MS `PageLayout`/`PageBreakCollection` 在 ExcelApi 1.9 列表中）+ 源码可证（两边都没实现）。

### S5. 【WPS 加载项】补齐"MS 侧有、WPS 宿主也支持、但我们没做"的 7 项

- **怎么补**（全部走 WPS 已有的 COM 克隆对象模型，无需新宿主能力）：
  1. 命名区域 `Workbook.Names`（增删查）
  2. 文档属性 `Workbook.BuiltinDocumentProperties` / `CustomDocumentProperties`
  3. 结构化表格 ListObject（建表 / 样式 / 汇总行 / resize）
  4. 区域复制 `Range.Copy` + `PasteSpecial`
  5. 超链接 `Range.Hyperlinks.Add`
  6. 图片与形状 `Shapes.AddPicture` / `AddShape` / `AddTextbox`
  7. `update_chart`（改标题 / 图例位置 / 位置尺寸）——**唯一被 `HOST_IMPLEMENTATION_GAPS` 显式声明为缺失的表格方法**
- **AI 能多做什么**：消除"同一句指令在 WPS 上失败、在 Microsoft 上成功"的**能力漂移**；WPS 用户也能用上结构化表格、命名区域、超链接与图表微调。
- **依据**：源码可证（我方无分支）+ 文档可证（WPS 宿主有对应对象）。

### S6. 【WPS 加载项 + Office.js 加载项】补条件格式完整规则族，并修"别名错位"

- **怎么补**：
  - **修错位（优先）**：`list_conditional_formats` / `update_conditional_format` 在 Office.js 侧目前都落进 `handleAddConditionalFormatting` → **列规则会变成加规则**。要么实现真语义（`range.conditionalFormats.load("items/...")`；ExcelApi 1.17 的 `changeRuleTo*` + `setRanges` 做"改规则"），要么**先撤掉这两个别名**避免误用。
  - **补规则族**：WPS `FormatConditions` 的 `AddIconSetCondition`（红黄绿灯）、`Top10`、`AddUniqueValues`、`AddContainsText`、`Add(type=xlExpression, formula)`；Office.js 侧的 `iconSet` / `presetCriteria`（重复值、空值）/ `topBottom` / `containsText` / `custom`(公式)。
  - 顺手修 ③-3 的互补缺口：Office.js 的 `colorScale` 分支当前**只 add 不设色**。
- **AI 能多做什么**：报表能出**红黄绿灯 / Top-N / 重复值 / 文本包含 / 公式驱动**这类业务信号，而不是只有最基础的大小比较。
- **依据**：**文档可证**（WPS `AddIconSetCondition`；Office.js 1.17 `changeRuleToIconSet` 等）+ 源码可证（我方实现面窄）。

### S7. 【Office.js 加载项 + WPS 加载项】补 Word 内容控件、域更新与逐条修订

- **怎么补**：
  - **内容控件**：WordApi 1.9 的 `DropDownListContentControl` / `ComboBoxContentControl` + `listItems`；VBA 侧另有日期/复选框/富文本。做成 `word_manage_content_controls`（增删改 + 填值）。
  - **域更新**：`Field.update()` / `Fields` 集合（Office.js）；WPS 侧 `Fields.Update`。补 `word_update_fields`（目录页码、交叉引用、页数）。
  - **逐条修订**：把现有的 `accept_all_revisions` / `reject_all_revisions`（**只有全量**）扩成按范围/按作者逐条，Office.js 侧用 `changeTrackingMode` + `Revision`。
- **AI 能多做什么**：能生成**可填模板 / 合同填空位 / 公文表单**，能保证"改完正文后目录页码是对的"，能**只接受自己改的那几条**修订而不是全盘接受。
- **依据**：**文档可证**（[WordApi 1.9](https://learn.microsoft.com/zh-cn/javascript/api/requirement-sets/word/word-api-1-9-requirement-set)）+ 源码可证（我方只有全量修订与"插入目录"）。

### S8. 【Office.js 加载项 + WPS 加载项】补数据透视表的字段编排与刷新

- **怎么补**：WPS 侧已有 `rowFields/columnFields/dataFields` 实现（源码可证）；**MS 侧要补的是字段名映射 + 字段编排**：Office.js `pivotTables.add(name, source, dest)` 只能建骨架，需再对 `pivotTable.rowHierarchies` / `dataHierarchies` 做 add，或至少把 `refresh()` 与字段读取做全。
- **AI 能多做什么**：从"建一张空透视表"升级为**"按产品线 × 季度 × 销售额一键出汇总透视"**，MS 侧能力追平 WPS。
- **依据**：源码可证（M3 字段名不通）+ 文档可证（Office.js 有 `PivotTable` 对象）。

### S9. 【WPS 加载项】补"单元格内局部格式（富文本）"

- **怎么补**：WPS 表格有 VBA 对齐的 `Range.Characters(start, length)`（可对单元格内某一段设 `Font.Bold` / `Color` / `Size`），把它包成 `patch_rich_text`（`address` + `[ {start, length, bold, color, size} ]`）。Office.js 侧无对应 API，**明确标注为 WPS 专属**，或在 MS 侧走 OOXML（成本高，建议后置）。
- **AI 能多做什么**：**"把 E5 里的 '-12%' 标红加粗"** 这种真正的"精细操控"，而不是整格统一样式——财经/经营分析里极高频。
- **依据**：**文档可证**（WPS `Range` 有 `Characters`，COM 克隆）+ **推测**（Office.js 无对应，需实测确认退化路径）。

### S10. 【桥接网关】把"宿主能力矩阵"做成机读契约

- **怎么补**：`HOST_IMPLEMENTATION_GAPS` 目前只有 Excel 且只到方法级；把 `capabilities().hosts.microsoft.word/ppt` 的自然语言（"未提供 Microsoft 结构化通道"）提升为**结构化字段**，例如 `{ structured: false, fallback: 'office_execute_script', notes: '...' }`，并让 Word/PPT 的 22 个工具在 `host:'microsoft'` 下**提前拒绝并给出正确替代路径**，而不是到宿主层才抛"未知的 RPC 方法"。
- **AI 能多做什么**：**提前选路**——拿到工具清单就知道"Microsoft Word 只能走 `office_execute_script`"，不会浪费 3 轮试错；也让 `bridge_get_capabilities` 成为可信的路由依据。
- **依据**：源码可证（`src/bridge/contracts/host-methods.ts` + `catalog.ts`）。

---

### 次优先（本轮未列入 10 条，但值得记一笔）

| 项 | 理由 |
|---|---|
| **WPS 事件驱动变更感知**（`wps.ApiEvent` + `SheetChange` / `WorkbookBeforeSave` 等 45+ 事件） | 能让 AI 从"盲写"变"协同编辑"（知道用户刚改了哪格、保存前拦一手）。当前只在 `hookWpsEvents` / `onSelectionChange` 里上报选区，未做成工具。依据：文档可证 + 源码可证 |
| **Excel 线程化批注的回复 / 已解决** | Office.js `Comment` 有 reply / resolved 语义；我们只有 add/list/delete。且要注意 ③-6 的"批注 vs 备注"跨宿主不等价 |
| **切片器 / 工作表视图 / 链接数据类型** | MS 侧文档可证支持；WPS 侧**未查证**。做之前需先实测确认两边行为 |
| **PPT 页面尺寸与母版/版式** | Office.js 有 `SlideMaster` / `SlideLayout`（PowerPointApi 1.3）但**无页面尺寸**；WPS 有 `PageSetup`。P5 报告已暴露 WPS 侧坐标换算的坑，母版化是根治方向 |

---

## 5. 依据等级与未实测声明

### 5.1 本机实测限制（影响面很大，务必先读）

| 限制 | 影响 |
|---|---|
| **Microsoft Excel 加载项未连接** | MS 侧（Office.js）**全部结论均无实测**：§1.1 的能力盘点来自源码阅读，§2.1 来自官方文档，§3.5 的 M1–M8 是**静态代码阅读推断的运行时后果**，未在真实 Excel 中复现 |
| **Mac 上无 COM 通道** | `resources/office/excel.ps1`（Windows COM，251 行）**未在本机执行过**；`fallbackToNative` 的 Windows 分支属"文档可证 + 源码可证，未实测" |
| WPS 侧 WPS 12.1.28496 已连接，但本轮**未做任何宿主调用** | WPS 侧结论全部为**源码可证**或**文档可证**；§3 中标注为"推测"的 WPS 能力未实测 |
| 未运行任何 `npm test` / `build` / `dist`（遵守任务约束） | 本轮**不含**回归验证；§3.5 的 M1–M8 需要实机复现才能定级为"已复现缺陷" |

### 5.2 依据等级逐项声明

| 等级 | 含义 | 本文中覆盖的结论 |
|---|---|---|
| **源码可证** | 直接读我方仓库源码得出，可复现 | §1 全部；§3 中带"源码可证"标记的条目；**§3.5 M1–M8 的字段不匹配事实**；§2.4 的"我方暴露面"数字 |
| **文档可证** | 官方文档（learn.microsoft.com / open.wps.cn / wpscdn 官方文档站）明确写了 | §2 全部；§3 中标"文档可证"的宿主能力条目（WPS `Shapes`/`Shape`/`Application` 对象、Office.js `PageLayout`/`Shape`/`RangeAreas`/1.17 条件格式与事件/WordApi 1.9 内容控件/PowerPointApi 各版本能力） |
| **推测** | 依据不足，**必须实测才能定论** | ①-2 `RangeAreas` WPS 无；①-3 WPS 无 SVG；①-4 WPS 无新图表类型；①-5 WPS 图表细粒度弱于 Office.js；①-8 WPS 无批处理；②-* 中"WPS 支持但需确认加载项内可调用"；④ 中 WPS 侧的 `PageSetup` / 图标集 / 富文本路径；M6 修正前的不确定表述已改为源码可证 |

### 5.3 本轮**没有**验证的事情（不要当成结论）

1. **MS 侧任何运行时行为**：M1–M8 的"静默 no-op / 报错 / 结果错误"是**代码路径推断**。M1（`params.rule` 为空 → 清空校验后不设规则）、M2（`text=""` → `includes("")` 恒真）的推断链条清晰，但仍需在真实 Excel 中复现确认。
2. **WPS 宿主是否真能执行文档里列的 API**：`Shapes.AddShape` / `AddSmartArt` / `Range.Characters` / 45+ 事件 / `PageSetup` —— 文档存在 ≠ 当前 WPS 12.1.28496 版本已实现 ≠ 加载项沙箱内可调用。**三层都要实测**。
3. **Office.js 要求集与真实宿主的对应**：本机没有 Microsoft Excel，无法确认用户实际 Office 版本支持到哪个要求集（`ExcelApi 1.9` 需要 1808+；1.17 需要 2302+）。若用户 Office 较旧，§4 的 S2/S4/S6/S7 方案会直接不可用——**落地前必须先做 `Office.context.requirements.isSetSupported` 探测**。
4. **Windows COM 路径**：`excel.ps1` 未实测；S3/S4 中"改走 COM"的建议需要 Windows 实机验收。
5. **`wps.FileSystem` / `wps.PluginStorage` / `wps.ApiEvent`**：文档可证存在，但本轮未确认加载项沙箱中的可用性与权限边界。

### 5.4 与既有材料的一致性

| 既有材料 | 本轮结论 |
|---|---|
| [compatibility-diff.md](../../compatibility-diff.md) §3 | 一致：宿主缺口台账按"可调用工具 ∩ 宿主实现"派生，不复现路由表 |
| `HOST_IMPLEMENTATION_GAPS.wps = ['update_chart']` 注释 | **核对通过**：当前 `wps-addon/src/dispatch.js` 的 61 个 RPC 方法中确实没有 `update_chart`（§1.2） |
| [mcp-sweep/README.md](README.md) 的"返回 success 但什么也没发生" | **本轮在跨宿主维度上新增了同类条目 M1 / M5 / M6 / M7 / M8**，且这批是**因宿主差异产生**，与版本漂移无关（不依赖 WPS 重启复测） |
| `capabilities().hosts.microsoft.word/ppt` | **发现口径缺口**：MS 侧"无结构化 Word/PPT 通道"只写在自然语言里，没进机读缺口台账（见 §1.4 与 S10） |
| `adapter.ts` → `routeOfficeFailure(supportedMethods: EXCEL_METHODS)` | **发现白名单过度声明**：路由表 30 项被整体当作"COM 支持列表"，但 COM 只有 28 项（缺 `update_chart`、`save_workbook`）；见 §2.2 |
| §1.4 三通道覆盖率表 | 新增：WPS 29/30、Office.js 29/30、COM 28/30，**没有任何一项三边全通** |
