# P5 · Microsoft（Office.js）通道修复：ISS-96 / ISS-93

执行者：Microsoft Office 侧修复员（子代理） · 起始时间：2026-09-22 晚
写区限制：只改 `office-addon/src/**`（必要时 `office-addon/public/taskpane.html`）。`src/**`、`wps-addon/**`、`skills/**`、`docs/**` 不改（本文档为任务指定的计划/证据落点）。

背景：本机 Microsoft Excel 加载项**首次可用**；桥接服务凭据 `~/.wps-bridge/token`，端点 `http://127.0.0.1:19890/api/v1/tool/call`，MS 通道需 `host: 'microsoft'`。
**运行中的任务窗格 = 磁盘构建物**（桥接服务 `office-addon/public/taskpane.js`）→ 改源码后必须 `npm run build:office-addon`，Excel 里需重新加载任务窗格。

---

## 0. 只读诊断结论（开工前，源码可证）

| 编号 | 现状（源码） | 根因 |
|---|---|---|
| ISS-96 | `excel/chart.js:handleCaptureSheetPreview` 在有原生图表时用 `chart.getImage()`（**真实**）；无图表时走 Canvas，`cellW=110` / `rowH=28` / `bannerH=42` 硬编码合成，还按标题关键词猜图表网格位置 | Office.js 无"区域截图"API，旧实现用画布"补"出一张像表格的图，`success:true` 返回 → AI 视觉自查结论与真实文件不符 |
| ISS-93-a | `excel/range.js:handleSetDataValidation` 只读 `params.rule`；网关（`gateway/excel.ts:setDataValidation`）传的是 `validationType/listItems/operator/minVal/maxVal/promptTitle/promptMessage/errorTitle/errorMessage` | 先 `dataValidation.clear()`，再因 `rule` 为空什么都不设 → **清空校验且不报错** |
| ISS-93-b | `excel/sheets.js:handleManageSheet` 只认 `action ∈ {copy,duplicate,rename,hide,show,color}`；网关只发 `{rename,move,tab_color,protect,unprotect}`，`tabColor` 取 `params.tabColor` 而网关传 `color` | `move`/`tab_color`/`protect`/`unprotect` 落到无分支 → 末尾 `return {success:true}` **静默 no-op** |
| ISS-93-c | `excel/range.js:handleFindReplace` 已兼容 `searchQuery` 且空串抛错（ISS-98，已修）；但网关的 `searchRange` / `maxResults` **未消费** | 局部（部分修复） |
| ISS-93-d | `excel/chart.js:handleGetCharts` 不读 `shapeName/chartIndex/chartTitle`，一律返回全表 | 选择器参数被忽略 |
| ISS-93-e | `excel/chart.js:handleCreateChart` 忽略 `dataRanges`（多段数据源）、`hasDataLabels`、`smoothLine`、`yAxis`、`seriesSettings`；`chartType:'pareto'` 无映射 | 5 个参数静默忽略 |
| ISS-93-f | `excel/range.js:handleUpdateRangeStructure` 用 `params.dimension`，网关传 `targetType`；`dimension` 缺省 `"rows"` | 想插列却插行；`set_size` 动作完全未实现 |

调用链确认（`src/bridge/office/adapter.ts`）：MS 通道把网关 `params` **原样**交给加载项（仅方法名映射），字段错位确定发生在 Office.js 侧。

---

## 1. 修复清单（逐条：编号 → 文件 → 改什么）

### ISS-96（高，伪渲染）→ 选**修法②：真实渲染能力缺失即显式报错，绝不返回合成图**

- **文件**：`office-addon/src/excel/chart.js`
- **改什么**：
  1. 删除整个 Canvas 合成分支（`cellW=110/rowH=28/bannerH=42` 及"按标题关键词猜图表网格"的 BIZ 特判），不再有第二段返回 `success:true` 的兜底。
  2. 保留**唯一真实渲染路径**：`chartName/name` 命中图表（或 `mode:'chart'`）时用 `chart.getImage()` 导出原生 PNG（Office.js 真实渲染）。
  3. 区域截图（含未指定图表、`mode:'sheet'`、只给 `address` 的情况）→ **抛错**，错误文案写清：① 本通道不支持真实工作表区域渲染；② 替代路径 = 改用 `host:'wps'` 的 `wps_capture_sheet_preview`，或在 Excel 里自行截图后把图给 AI；③ 说明旧版合成图是伪造的、已移除。
  4. 附带：图表命名支持 `Chart 1` / `图表 1` 归一化匹配，取不到图时也抛错并列出可用图表名。

### ISS-93-a（高，数据有效性清空后什么都不设）

- **文件**：`office-addon/src/excel/range.js`
- **改什么**：
  1. 新增 `buildDataValidationRule(params)`：**优先** WPS 语义（`validationType` + `listItems/operator/minVal/maxVal`），兼容旧字段 `rule`（直通）。
  2. `list` → `{list:{inCellDropDown:true, source:"a,b,c"}}`；`number_range` → `wholeNumber/decimal` + `Operator`（`between/greater_than/less_than/equal` → `Between/GreaterThan/LessThan/EqualTo`）；`date`/`text_length`/`custom` 亦按 WPS 语义支持。
  3. `promptTitle/promptMessage/errorTitle/errorMessage` → `promptTitle/promptMessage/errorAlert.title/errorAlert.message`（`showPrompt`/`showErrorAlert` 依有无内容设置）。
  4. **顺序修正**：先算规则 → 规则构造失败（无法识别的 `validationType`、`list` 缺 `listItems`、`number_range` 缺阈值）**先抛错且不清空**；规则合法才 `clear()` + 赋值。
  5. 返回值回读真实 `range.dataValidation.rule/type/prompt`，不再只回 `{success:true}`。

### ISS-93-b（高，manage_sheet 四个 action 静默 no-op）

- **文件**：`office-addon/src/excel/sheets.js`
- **改什么**：
  1. `move`（别名 `activate`/`position`）：读 `targetIndex`（兼容 `position`/`index`/`targetPosition`），`targetIndex` 缺失或非数字 → **抛错**；用 `sheet.position = targetIndex - 1` 真实移动并**读回** `position`。
  2. `tab_color`（别名 `color`/`set_tab_color`）：颜色取 `color`（兼容 `tabColor`），缺失 → **抛错**；设置后读回 `tabColor`。
  3. `protect` / `unprotect`：Office.js `WorksheetProtection` 在 Excel on Mac 桌面版**不可靠/未实现** → **显式抛错**，并在错误里给出替代路径（改用 `host:'wps'` 的 `manage_sheet`，或 Excel「审阅 → 保护工作表」手动加保护）。绝不静默成功。
  4. `rename` 缺 `newName` → **抛错**；未知 `action` → **抛错并列出支持值**（原来是静默 `success:true`）。
  5. 顺带修正 `action === "color"` 分支读 `params.color || params.tabColor`。

### ISS-93-c（find_and_replace 选择器）

- **文件**：`office-addon/src/excel/range.js`
- **改什么**：范围改为 `params.searchRange || params.address || params.range`（网关传 `searchRange`）；`maxResults` 生效（默认 50，`0`/负数视为不限）；返回 `truncated` 标记；替换分支返回真实命中数需 `range.load("text")` 前置——保持"找不到就不谎报"。

### ISS-93-d（get_charts 忽略选择器）

- **文件**：`office-addon/src/excel/chart.js`
- **改什么**：`shapeName`（按 `name`）与 `chartTitle`（标题**包含**匹配）本地过滤，`chartIndex`（1 基）取第 N 张；`chartIndex` 越界 → **抛错并给出实际图表数量**；过滤后为空 → 抛错列出可用图表。返回值加 `filtered`/`selector` 便于调用方确认真的按选择器返回。

### ISS-93-e（add_chart 静默忽略 5 个参数）

- **文件**：`office-addon/src/excel/chart.js`
- **改什么**：
  1. `dataRanges`：多段数据源用 `chart.setData(sourceData, seriesBy)` 逐段追加（取不到则回退单段并在 `warnings` 里说明）。
  2. `hasDataLabels` → `chart.dataLabels.showValue`；`smoothLine` → 每个系列 `series.smooth = true`（折线/散点）；两者**读回**校验，与请求不一致写入 `warnings`。
  3. `yAxis`（`min/max/step/numberFormat/title`）→ `chart.axes.valueAxis`；`step` 用 `majorUnit`。
  4. `seriesSettings`（`seriesIndex/color/smooth`）→ 对应系列着色与平滑；`seriesIndex` 越界 → **抛错**（不静默）。
  5. `pareto`：Office.js 无原生 Pareto 图表类型 → **抛错**说明应当用 `column_clustered` + 自行排序/累计占比，绝不静默建成柱状图（ISS-17 同类问题）。
  6. 颜色/标签/轴设置失败写入 `warnings`（保持既有"校验读回"风格），不再整体 `catch` 后无声吞掉。

### ISS-93-f（manage_rows_and_columns 的 targetType 被忽略 → 想插列却插行）

- **文件**：`office-addon/src/excel/range.js`
- **改什么**：
  1. `dimension` 取值改为 `targetType || dimension`（`targetType` 优先，兼容 `'column'/'columns'/'列'`、`'row'/'rows'/'行'`），映射成 `rows|columns`；**两者都无法识别 → 抛错**（绝不再默认 rows）。
  2. `action='set_size'`（网关 `manage_rows_and_columns` 支持）→ `rowHeight` / `columnWidth` 真实设置，`size` 缺失 → 抛错。
  3. `action='hide'/'unhide'` → `range.hidden = true/false`（原来落到"无分支 + success:true"）。
  4. `index` 支持数字行/列号与**列字母**（`'B'`），`count` 越界/非正 → 抛错；未知 `action` → 抛错并列出支持值。
  5. 允许 `address` 直给区域（与 `index` 二选一）。

### 附带（同一写区、同一批）

- `office-addon/src/excel/range.js`：`handleClearRange`/`handleCopyRange`/`handleSetHyperlink` 对缺失必需参数**显式抛错**（现在会因 `getRange(undefined)` 抛出难懂的宿主错误）。
- 所有新增错误文案统一带 `[Office.js 通道]` 前缀，便于与 WPS 侧报错区分。

---

## 2. 构建与检查（每批改完即跑）

```
npm run build:office-addon
node --check office-addon/public/taskpane.js
node scripts/build-office-addon.mjs --check
npm run typecheck
npm test
npm run check:claims
```

## 3. 真实 MS Excel 验证计划

1. 先确认任务窗格跑的是新代码：`curl -sk https://127.0.0.1:19891/office-addon/taskpane.js | grep <标记>`，标记取本次新增的字符串。
2. Excel 里重新加载任务窗格（插入 → 我的加载项 → 刷新，或关闭重开任务窗格）。
3. 用 `host:'microsoft'` 实测（每条都要**独立读回**）：
   - ISS-96：无图表区域调 `capture_sheet_preview` → 期望**报错**且文案含替代路径；再对真实图表调用 → 期望返回 `imageBase64`。
   - ISS-93-a：`set_data_validation` 设下拉 → 用 `execute_script` 读回 `Validation.Type/Formula1`。
   - ISS-93-b：`manage_sheet` 的 `tab_color` / `move` → 读回 `Tab.ColorIndex`/`Index`；`protect` → 期望显式报错。
   - ISS-93-f：`targetType:'column'` 插入列 → 读回 `UsedRange.Columns.Count` 与表头位移，确认列的位移而非行。
   - ISS-93-d：建 2 张图 → `get_charts` 带 `chartIndex/chartTitle` → 期望只回对应图。
   - ISS-93-e：`add_chart` 带 `hasDataLabels/smoothLine/yAxis.min` → 读回 `HasDataLabels`/`Axes(xlValue).MinimumScale`。
4. 结果如实测实记；不能实测的项标"未验"，不得写成通过。

## 4. 待决策点（汇报里给使用者）

- ISS-96 选②（显式报错）。若希望 macOS 也有真实区域渲染，需要改 `src/**`（网关/原生驱动，`CopyPicture` 走 Windows COM 或 macOS 截屏），**超出本次写区**。
- `manage_sheet` 的 `protect/unprotect`：Office.js 侧不实现（报错）。若要在 MS 宿主支持，需评估 `WorksheetProtection` 在 Excel on Mac 的实际可用性。

---

## 5. 执行记录（边做边更）

| 时间 | 动作 | 结果 |
|---|---|---|
| 开工 | 只读诊断（源码 + 网关字段 + 台账） | 6 类错位全部源码可证，见 §0 |
| — | 计划落盘（本文件） | 完成 |
