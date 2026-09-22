# Excel 侧能力补强验收：CAP-10 违规定位 / CAP-41 style token / 「先校验、后动手」顺序修复

- 执行人：teammate `excel-cap-fixer`（task-6）
- 写入范围：`wps-addon/src/excel.js`、`src/bridge/tools/definitions/excel.ts`、本文件
- 本轮**未执行**：`npm run build:wps-addon`、`setup`、服务重启、任何真机工具调用（Lead 统一串行处理）
- 状态口径：**只做源码改动 + 静态检查 + 模拟宿主验证**；**真机（WPS ET）验证未做**，下文逐项标注

---

## 0. 验证层级（先看这里，避免误读完成度）

| 层级 | 本轮是否做到 | 说明 |
|---|---|---|
| 静态检查（语法/类型） | ✅ | `node --check wps-addon/src/excel.js`；`npm run typecheck` 0 错 |
| 契约一致性 | ✅ | `npm run check:params`、`npm run check:claims` 通过；快照差异见 §7 |
| 模拟宿主行为验证 | ✅ | `/tmp` 临时 harness（按构建脚本同序拼接 `src/**` + 假 WPS 宿主）**89/89 通过**，原文见 §10 |
| 生成物一致性 | ❌ 未做（且预期失败） | 改了 `src/**` 未重建 `wps-addon/addon-core.js` → `npm test` 的「生成物与源码一致」会失败，**这是预期的**，按调度由 Lead 统一重建 |
| **真机（WPS ET）读数** | ❌ **未验证** | 本轮禁止调用真机；`probes`/宿主 API 真实可用性、性能、`Validation.Formula1` 形态均待真机确认 |

---

## 1. CAP-10 数据验证**违规定位**

### 1.1 实现方式

**宿主侧（`wps-addon/src/excel.js`）**

- 新增 `include` 取值 **`validationViolations`**（`getRangeStyles` 的 allowed 列表 + §7 的 schema enum）。
- 该取值是**区域级附加读取**，不是逐格样式字段：
  - `mode: "summary"`、`mode: "cells"` 两种模式都会在响应顶层追加 `validationCheck`；
  - 不写进默认 `include`，**旧调用方的响应形状不变**（模拟宿主已断言）。
- 核心函数：
  - `readCellValidationRule(cell)`：单格规则读取，**三态返回**（`{ok:true,rule:null}` = 宿主明确说没校验 / `{ok:true,rule}` = 读到 / `{ok:false,error}` = 读失败）；
  - `collectValidationViolations(sheet, range, {maxCells})`：区域级扫描与判定；
  - `evaluateValidationRule` / `compareByOperator` / `parseValidationBound` / `coerceValidationNumber` / `describeValidationRule`：逐类型判定与"期望"文案；
  - `parseValidationListFormula` / `readValidationListReference` / `resolveValidationListInfo`：list 型规则的候选项解析（字面量 + 区域引用，按规则签名缓存，避免逐格重读引用区域）。
- **扫描方式**：只用已证实可用的 `Range.Value2`（整块一次读回）与 `Range.Validation` 逐格读；
  **未使用** `SpecialCells(xlCellTypeAllValidation)` 整表枚举——它在本机未验证，且 `dispatch.js` 的反射护栏把它列为"保守跳过"（触达宿主内部集合原生 getter 有崩溃史）。区域上限由 `maxCells` 控制（默认 300，硬上限 2000），截断时 `truncated:true` + warnings。

**响应 `validationCheck` 结构**

```jsonc
{
  "rangeAddress": "$E$5:$E$50", "totalCells": 46, "scannedCells": 46, "truncated": false,
  "validatedCells": 46, "skippedBlankCells": 3, "blankWithUnknownIgnoreBlank": 0, "readErrorCells": 0,
  "ruleCount": 1,
  "rules": [{ "ruleKey": "R1", "typeName": "list", "operatorName": "between",
              "formula1": "\"已通过,待复测\"", "expectation": "必须是列表中的一项: 已通过 / 待复测",
              "cellCount": 46, "sampleCells": ["E5","E6","E7"] }],
  "violations": [{ "address": "E9", "value": "已作废", "ruleKey": "R1", "ruleType": "list",
                   "expectation": "必须是列表中的一项: 已通过 / 待复测", "reason": "值不在允许的候选项中",
                   "allowedValues": ["已通过","待复测"] }],
  "violationCount": 1, "violationsTruncated": false,
  "unevaluated": [{ "address": "E20", "reason": "自定义公式规则（xlValidateCustom）无法在宿主外求值" }],
  "unevaluatedCount": 1, "warnings": [], "message": "已比对 [表] $E$5:$E$50：46 个带校验的单元格中 1 个取值越界"
}
```

### 1.2 判定语义（"读不到"一律不进"通过"）

| 情况 | 处理 |
|---|---|
| 单元格没有校验规则（`Type = -4142`） | 不参与判定（与"越界"无关） |
| 规则读失败（`Validation` / `.Type` 抛错） | 进 `unevaluated` + `readErrorCells`，**不当成"没规则"** |
| 自定义公式规则（`xlValidateCustom`） | 进 `unevaluated`（宿主外无法求值），不猜 |
| `list`：字面量（含宿主读回带外层引号的形式） | 逐项比对（先精确、再忽略大小写） |
| `list`：区域引用（`=$D$1:$D$5` / `=Sheet1!$D$1:$D$5` / 裸引用） | 读引用区域取值作为候选；引用读不到 → `unevaluated` |
| `whole_number` / `decimal` | 8 种运算符全支持；值支持数字、文本数字、千分位、货币符号、百分号、日期文本 |
| `date` / `time` | 边界与取值都折算为 **Excel 序列号**（`25569 = 1970-01-01`），支持 `=DATE(y,m,d)` / `=TIME(h,m,s)` / ISO 文本；折算不了 → `unevaluated` |
| `text_length` | 按字符数比较 |
| 值非数值（如数值规则下填了文本） | 记**违规**并给原因（Excel 的"停止"警告会拒绝该输入） |
| 空单元格 + `IgnoreBlank=true` | 跳过并计入 `skippedBlankCells` |
| 空单元格 + `IgnoreBlank=false` | 记**违规**（"规则未允许空值"），与 Excel"圈释无效数据"口径一致 |
| 空单元格 + `IgnoreBlank` 读不到 | 跳过 + `blankWithUnknownIgnoreBlank` + warning（**明确写"可能漏报"**） |
| 未知运算符 / 类型 | `unevaluated`（不猜通过） |
| 违规条数 > 200 | 截断 + `violationsTruncated` + warning 说明还有多少条未列出 |

### 1.3 验证到什么程度

- 模拟宿主 94 项断言中 **CAP-10 相关 31 项全过**（列表/数值/日期/长度/引用/千分位/`IgnoreBlank=false`/未知运算符/读失败/截断/组合用法/`cells` 模式），见 §10。
- **模拟宿主发现并修掉一个真实缺陷**：违规地址曾按 `firstRow + r + 1` 计算（多 +1），会把违规单元格报成**右下方的邻居格**（如 F5 报成 G6）。真机必然复现，已修正为 `firstRow + r`（`range.Row/Column` 本来就是 1 基左上角坐标），并加了"原点读不到时退回 `cell.Address()`"的兜底。
- 未验证 / 待确认：
  - **真机 WPS 的 `Validation.Formula1` 实际形态**（带引号字面量 / 引用）——两种都做了兼容，但真机只走其中一条；**待确认**。
  - **`Validation.Operator` 数值码**（1..8）在 WPS 上是否与 VBA 常量一致——**未验证**。
  - `IgnoreBlank=false` 时"空格算越界"是**按 Excel VBA 语义推断**的，WPS 行为**待确认**。
  - 性能：逐格读 `Validation` 约 5~6 次宿主属性访问/格，300 格 ≈ 1800 次调用；真机耗时**待确认**（若偏慢，调 `VALIDATION_SCAN_DEFAULT`）。

---

## 2. 「先校验、后动手」破坏性顺序修复（Lead 追修 + 同类自查）

原则：**任何"校验失败还会留下副作用"的顺序都是错的**。本轮按此原则修了 **3 处**（1 处 Lead 指定 + 2 处自查发现）。

### 2.1 `set_data_validation`（Lead 指定）

**改前的破坏路径**（复核确认）：

```js
const targetRange = sheet.Range(address);
try { targetRange.Validation.Delete(); } catch (e) {}      // ① 先清空该区域原有校验

if (validationType === "list") {
  if (listStr.trim() === "") throw new Error("...必须提供 listItems...");   // ② 校验在删除之后
  ...
} else if (validationType === "number_range") { ... }
else { throw new Error(`未知的 validationType: ${validationType}`); }        // ③ 非法类型同样在删除之后
```

**调用方少传一个参数（或传错 `validationType`）时的真实后果**：

1. `Validation.Delete()` 已经执行 → 用户该区域**原有的数据有效性被清掉**；
2. 然后才抛错，调用方只看到一条参数错误；
3. 原实现 `catch (e) {}` 吞掉 Delete 的错误，连"删过了"都没留痕；
4. 破坏性比修复前（静默 no-op）**更强**：静默 no-op 至少不动用户设置。

同类问题（同一函数内自查）：

| # | 问题 | 后果 |
|---|---|---|
| A | `validationType` 非法 → `throw` 在 `Delete()` 之后 | 先毁后报错（同上） |
| B | `operator` 非法（如 `approximately`）→ 静默按 `between` 走 | 写出一条**语义完全不同**的规则，且是破坏性的 |
| C | `operator: "less_than"` 时 schema 文档写"比较 maxVal"，代码取的是 `minVal`；只传 `maxVal` 的调用方会得到 `< 0` | 静默写错规则；文档与实现不一致 |
| D | `between` 只传 `minVal` → `Formula2 = undefined` 传给宿主 | 无上限规则或宿主报错（都在 `Delete()` 之后） |
| E | 写入后**完全没有读回核对**，但工具描述早已写"apply（默认）写入并读回核对" | 描述与实现不符；宿主静默不生效时仍回 `success:true` |

**改后行为**

- 抽出 `buildDataValidationPlan(params)`，**在任何修改之前**完成全部参数校验并抛错，错误文案统一带一句
  「本次未做任何修改（原有校验保持不变）」：
  - `list`：`listItems` 必须为非空数组/非空串（空项自动剔除）；
  - `number_range`：`operator` 必须是 `between / greater_than / less_than / equal`；`minVal`/`maxVal` 必须是**有限数值**；`between` 必须**同时**给两个边界；
  - `less_than`：优先取 `maxVal`（schema 文档口径），未传才退回 `minVal`，并在 `warnings` 里说明用的是哪个；
  - `validationType` 只认 `list` / `number_range`。
- 校验通过后才：读原规则（用于追溯/回滚）→ `Delete()` → `Add()`。
- **写入后读回核对**（`verifyDataValidationRule`）：类型 / 运算符 / 上下限 / 候选项集合 / 下拉开关与请求值逐项比较；
  - 核心规则不一致 → **`success:false`** + `warnings` 给出「请求值 vs 宿主读回值」+ `readBack` 原始读数；
  - 提示/报错文案等附加项落不上 → 只进 `warnings`，不影响 `success`。
- `Add()` 抛错（原规则已被 Delete 掉）→ `restoreValidationRule()` 把写入前的规则**尽力写回并读回核对**：
  - 恢复成功 → `rollback.restored=true`，message「原校验已恢复」；
  - 恢复失败 → `rollback.restored=false` + 原因，message「原校验未恢复（该区域现在可能没有校验，请重新设置）」，**不假装无事发生**。
- 新增响应字段：`requested` / `readBack` / `replacedPreviousRule` / `rollback`（失败时）/ `warnings`。
- `action: "read"` 分支未改动。

**验证**：模拟宿主 22 项断言全过（4 类非法参数均"报错且 `Delete` 调用次数为 0"、原规则读回仍在、正常写入读回一致、
宿主静默不生效时 `success=false`、首次 `Add` 抛错时回滚成功/失败两条路径都如实报告）。**真机未验证**；
`Validation.Add` 失败后回滚的真实成功率、`AlertStyle` 读回值**待确认**。

### 2.2 `add_conditional_formatting`（自查发现，同类）

**改前的破坏路径**：`clearExisting` 在函数开头就 `range.FormatConditions.Delete()`，而参数校验分散在各分支里、都在其后：

- 未知 `ruleType`（末尾 `else` 抛错）→ 先清空该区域**全部既有条件格式**，再报"未知类型"；
- `icon_set` + 不支持的 `iconSet` → 同上；
- `text_contains` 缺 `containsText` → 同上；
- `formula` 缺 `formula1` → 同上。

**改后**：函数开头集中做**四项校验**（`KNOWN_RULE_TYPES`、`ICON_SET_CODES`、`containsText`、`formula1`），
并把 `ICON_SET_CODES` 提到函数作用域顶部供后文复用；报错文案统一追加「本次未做任何修改（原有条件格式保持不变，若传了 clearExisting 也**没有**执行清除）」。
另**新增**一条同类校验：`cell_value` + `operator="between"` 必须同时给 `formula1`/`formula2`（少了上限时宿主 `Add` 会抛错，而那时规则已被删掉）——这条是**新增拦截**，属于行为变化，若 Lead 认为不该拦可直接去掉。

**验证**：模拟宿主 8 项断言全过（4 类非法参数 + `between` 缺上限，全部"报错且 `FormatConditions.Delete` 调用数为 0"、原规则仍在；合法调用仍正常清除并写入）。

### 2.3 `set_filter_and_sort`（自查发现，同类）

**改前的破坏路径**：`enableAutoFilter: true` 会先执行 `targetRange.AutoFilter()`（真机上是一次可见的写入），
而 `sortRules.colIndex` 的越界校验在后面（代码注释还写着"先做边界校验"，但已经晚了）。
于是 `enableAutoFilter:true` + 越界 `colIndex` → **先给用户打开筛选、再报错**。

**改后**：把边界校验前移到任何写入之前（只读 `targetRange.Columns.Count`，无副作用；读不到列数时**不拦**，避免误拒），
报错追加「本次未做任何修改（未开启/关闭筛选，未排序）」；原位置的重复校验已移除。

**验证**：模拟宿主 3 项断言全过（越界 `colIndex` + `enableAutoFilter` → `AutoFilter()` 调用数为 0；合法调用仍能开启筛选）。

### 2.4 同类模式的其他候选（**未修**，留给 Lead 排期）

用静态扫描（函数内"首个写入调用"与"其后 throw"的位置对比）扫了 `excel.js` 全部写函数，除上述 3 处外还有若干**疑似**同类，但
每一处都需要逐分支阅读才能判定是否真的是"先动手后校验"，且部分涉及已有台账缺陷（如 ISS-118），本轮**未改**：

| 函数 | 疑似点 | 备注 |
|---|---|---|
| `createPivotTable` | 目标表不存在时 `wb.Worksheets.Add()` 自动建表，随后改名失败才 `throw` | 会在工作簿里留下一个未命名成功的 `SheetN`；修法是"建表前先校验表名（复用 `createWorksheet` 的 M-4 校验）" |
| `configurePrintLayout` | 先 `HPageBreaks.Add(...)` 等写入，后面仍有 `throw` | 需逐分支确认是否真为参数校验 |
| `insertDimension` / `modifyRowsColumns` | 先 `Insert()`，后面仍有 `throw` | 同上 |
| `managePictures` / `manageTable` / `manageNamedRange` / `manageHyperlink` / `deleteChart` / `manageCellComments` | "先删对象、后校验"的形态 | 多数是 `action` 分支内的正常删除动作，需人工判定 |

判定口径与工具：扫描脚本用"函数体内首个写入调用行号 vs 其后 `throw new Error` 行号"做启发式（有误报，例如把只读分支里的赋值算成写入），
**结论需人工复核**。

---

## 3. `find_and_replace` 空串护栏（宿主侧，Lead 追加）

### 3.1 改前的破坏路径

```js
if (searchQuery === undefined || searchQuery === null) {
  throw new Error("缺少必要参数: searchQuery (要查找的文本或数值)");
}
// ← 空串 "" 从这里穿过去
...
const testMatch = (text) => text.includes(queryStr);          // includes("") **恒真** → 区域内每个非空单元格都"命中"
...
source.replace(new RegExp("", "gi"), replaceText)             // 空正则：在**每个字符之间**插入 replaceText
if (hitFormula) cell.Formula = newVal; else cell.Value2 = newVal;
```

- 空搜索串 → 区域内**每个非空单元格**都命中；
- 带 `replaceText` 时按空正则替换：`"abc"` → `"XaXbXcX"`（`replaceText` 被插到每个字符之间），
  **整片内容被改坏**，且函数照样返回 `success:true` + `replacedCount=N`；
- 只查找不替换时返回"全区域命中"，同样是无意义的假结果；
- 只拦 `undefined`/`null`：JSON-schema 的 `type: "string"` 不拦空串，所以调用方传 `""` 完全合法地到达这里。

上游（桥接层 `src/bridge/gateway/excel.ts` 与 `office/normalizer.ts`）已挡工具入口，但**加载项可以被 WebSocket RPC 直接调用**，
宿主侧没有护栏时这条路径仍然通着——护栏只留一层就等于没有。

### 3.2 改后行为

- 在**任何读取/写入之前**（紧跟 `undefined/null` 检查）增加空串拦截，`String(searchQuery) === ""` 即拒绝；
- 报错文案说明原因（`includes('')` 恒真 + 空正则替换）与「本次未做任何修改」；
- 口径与桥接层**保持一致：只拦空串，不拦空白字符**（`" "` 是合法的查找目标）；
- `replaceText: ""`（把命中内容删掉）仍是合法操作，不受影响。

### 3.3 验证到什么程度

- 模拟宿主 5 项断言全过（§10）：空串 → 宿主侧直接报错；被拒后**内容零改动**；报错文案含"未做任何修改"；
  空格查询不被误拦；正常替换仍然工作（`replacedCount=2`、内容变 `促销商品A`）。
- 真机未验证（走的是同一个 JS 分支，判定逻辑与宿主无关，风险主要在"真机上是否真的零改动"——被拒路径不触碰宿主，**理论上零改动**）。

## 4. CAP-41 `get_style_token`（WPS 独有）— 读原表设计语言

### 4.1 改前的问题

原实现只读一个取样格的字体 + 底色，而且：

```js
const fontName = range.Font.Name || "微软雅黑";        // 读不到就编
const fontSize = range.Font.Size || 11;                // 读不到就编
headerBackgroundColor: headerBg || "#1E3A8A"           // 读不到就编
```

调用方拿到的是**臆造值**而不是原表状态；`titleFontName` 也只是 `A1` 的字体名，没有"设计语言"可言。

### 4.2 实现方式（宿主 API **先探测后实现**）

| 目标 | 读取路径 | 探测/兜底 |
|---|---|---|
| 取样格真实样式 | `sampledCell`：Font.Name/Size/Bold/Italic/Color、Interior.Color + **Interior.Pattern**（判 `hasFill`，避免把"无填充"的白色当成刻意白底）、NumberFormat、HorizontalAlignment、RowHeight/ColumnWidth、下框线 | 任一读不到 → `null` + warnings（列出具体字段名），**不编默认** |
| 字体层级 | `fonts.title / header / body / caption`：基于 240 格有界普查（`census`）按"设计指纹"分组统计（出现次数、代表格、代表性文本） | 判据写进 `fonts.heuristic`，明确标注是**启发式推断**不是宿主读数 |
| 主题色板 | 候选逐个 `try`：`Workbook.ThemeColorScheme` → `Workbook.Theme` → `Application.Theme`（对象结构支持 `Count/Item`，也支持数组） | 成功记 `probes.ok=true` + 来源；全失败 → `palette.themeColors=null` + `unavailable` 说明；不虚构 |
| 工作簿调色板 | `Workbook.Colors`（支持整体数组，也支持按下标 1..56 读） | 两种形态都兼容；非有限值不转色（避免 NaN→`#000000`） |
| 实际用色排行 | `palette.observed`：从普查统计 font/fill 颜色出现次数并给出取样地址 | 这是"原表长什么样"最可靠的证据 |
| 表格样式 | `sheet.ListObjects` → `TableStyle.Name` / 范围 / ShowHeaderRow / ShowTotals | 与 CAP-19 已真机验证的 `manage_table` 读回同一条路径 |
| 条件格式风格 | **有界区域**扫描 `FormatConditions`：`UsedRange`、`A1:L40`、该区块前 12 行 | 不做全表枚举；读不到 → `unavailable` 明说"仍可能存在但未枚举" |

**明确不碰**（写在代码注释里）：`wb.Styles`、`SpecialCells`、整列/整行范围（`A:A`、`5:5`）——`dispatch.js` 的反射护栏已记录
"整表范围对象求值会令 WPS 进程崩溃"（ISS-89 取证），未验证的宿主 getter 本轮一律不赌。

返回还包括 `probes`（每项探测的成功/失败 + 原因）、`unavailable`（整体不可用项 + 宿主错误）、`warnings`、`census`（分组明细），
以及旧字段名 `fontName / titleFontName / sampleFontSize / sampleBold / fontColor / headerBackgroundColor`
（**值全部来自真实读回，读不到即 `null`**，只保留字段名做向后兼容）。

### 4.3 验证到什么程度

- 模拟宿主 **21 项断言全过**（§10），包括两条反向用例：
  - `Font.Name` 读取抛错 → `fontName === null` 且 warnings 说明（旧实现会返回"微软雅黑"）；
  - `headerBackgroundColor` 无填充时为 `null`（旧实现会返回"#1E3A8A"）；
  - 主题色三项全失败 → `themeColors=null` + `unavailable` 有条目，`probes` 记录三条失败原因；
  - 主题色探针返回 `Count/Item` 结构时走成功路径并标明 `themeColorsSource`。
- 未验证 / 待确认：
  - **真机上这批宿主属性是否可用**：`Interior.Pattern`、`Font.Italic`、`Range.Row/Column`、`Borders.Item(9)`、
    `Workbook.Colors`、`Workbook.ThemeColorScheme`、`Workbook.Theme`、`Application.Theme` —— **全部未验证**；
    代码已按"探测失败即如实上报"实现，真机只要看 `probes` 就知道哪个可用。
  - **真机耗时**：普查 240 格 × ~13 次属性访问 ≈ 3000 次宿主调用，可能偏慢；`STYLE_CENSUS_MAX_CELLS=240` 是保守值，
    **待确认真机耗时**后决定是否下调。
  - 条件格式只扫"已用区域 + A1:L40 + 前 12 行"的有界窗口，**其他区域的条件格式不会被枚举**（已在 `unavailable` 明说）。
  - `wps_get_style_token` 在 Microsoft 宿主上本就声明为未实现（`HOST_IMPLEMENTATION_GAPS.microsoft`），不受本次影响。

---

## 5. 跨宿主缺口（未修，需派给 office 侧）

`include: ["validation"] / ["validationViolations"]` **只在 WPS 侧实现**。
`office-addon/src/rpc.js` 的 `get_range_styles` 不读 `include`，`src/bridge/office/normalizer.ts` 也只在响应侧折算样式；
因此 `excel_get_range_styles(host: "microsoft", include: ["validationViolations"])` 会**静默返回基础样式、没有 `validationCheck`**——
又是"调用成功但没生效"的形状。本轮我只能在 `wps_get_range_styles` 的工具描述里写明宿主差异（已加），
**office 侧分支未加**（不在我的写入范围）。建议：office 侧补 `getInvalidCells`（CAP-10 台账的 MS 方案），
或至少在 MS 通道对 `include` 显式报"宿主未实现"。

---

## 6. 未纳入本轮（有意不做）

- `createPivotTable` 等 §2.4 表中的疑似同类项（需逐分支判定，部分牵涉 ISS-118）。
- **真机探测** `SpecialCells(xlCellTypeAllValidation)`：CAP-10 台账的"MS `getInvalidCells`；WPS 读 `Validation` 属性"里，
  整表枚举本可让违规定位一次拿到全部带校验单元格；本轮出于崩溃风险（未验证 + 反射护栏保守跳过）**未采用**。
  若 Lead 愿意先做一次隔离真机探测（新建空白簿、单次 `SpecialCells(-4174)`），可行的话可显著降低扫描成本。

---

## 7. 契约快照差异（实际导出，未写入仓库）

命令（只读导出到 `/tmp`，不覆盖仓库基线）：

```bash
npx tsx scripts/snapshot-tools.ts /tmp/tools-snapshot.excel-caps.json
```

与 `docs/acceptance/2.1.0-p0p1/tools-snapshot.p19.json` 对比（该导出在 §2.2/§2.3 修复之前生成，工具定义部分不受影响）：

- 工具数 **137 → 138**（新增的是他人 CAP-09 的 `wps_ppt_configure_layout`，不是本任务）；
- 本任务**没有新增/删除任何工具**，只有 6 条既有条目（3 工具 × wps_/excel_ 两个入口）的描述与 enum 变化：
  - `wps_get_range_styles` / `excel_get_range_styles`：`include` 枚举新增 `validationViolations`，描述补充判定语义、截断、宿主差异；
  - `wps_set_data_validation` / `excel_set_data_validation`：描述改为"先校验参数、后动手 + 写后读回核对 + 违规定位入口"（原描述还写着"本工具族暂无数据有效性读回工具"，已过期）；
  - `wps_get_style_token` / `excel_get_style_token`：描述改为完整设计语言返回结构 + "读不到不编默认值"。
- 其余 3 个变化条目（`wps_word_*`）属他人改动。

**待办**：`npm run snapshot:tools docs/acceptance/2.1.0-p0p1/tools-snapshot.p20.json`（或按当轮约定命名）由 Lead 在重建产物后统一生成。

---

## 8. 回归检查（本轮实际执行）

| 检查 | 结果 |
|---|---|
| `node --check wps-addon/src/excel.js` | 通过 |
| `npm run typecheck` | 通过（0 错） |
| `npm run check:params` | 通过（自检通过；未发现 schema 有、处理器不读的参数） |
| `npm run check:claims` | 通过（口径：对外工具 138 个；无矛盾表述） |
| `node --import tsx --test tests/addon.test.ts tests/contract-consistency.test.ts tests/failure-routing.test.ts` | **30/30 通过**（`addon.test.ts` 读的是旧生成物 `addon-core.js`，只证明"未破坏既有产物行为"） |
| 模拟宿主 harness（89 项） | **89/89 通过**，原文见 §10 |

---

## 9. 交接 / 未完成项

1. **重建部署入口**（Lead，串行）：`npm run build:wps-addon` → `node --check wps-addon/addon-core.js` → 重跑 `npm test`（预期「生成物与源码一致」恢复通过）。本次改动**必须重建**才可能真机生效。
2. **真机验收清单（建议按此顺序，一条调用看一圈读数）**：
   - `wps_get_style_token`：看 `probes` 哪些 `ok=true`（重点 `Workbook.Colors` / 主题色三候选 / `ListObjects` / `FormatConditions` / `Interior.Pattern`），
     确认 `fonts.body` 与肉眼一致、`palette.observed` 命中表头填充色，并**记录耗时**；
   - `wps_get_range_styles` + `include:["validation","validationViolations"]`：先用 `set_data_validation` 写一列下拉，
     故意塞一个越界值，确认 `violations` 精确定位到该格（**地址不许偏**）；
   - `wps_set_data_validation` **缺 `listItems`** 调一次：确认报错以后**原下拉仍然在**（Lead 追修的核心验收点）；
   - `wps_set_data_validation` 正常写一次：确认 `readBack` 与请求一致、`replacedPreviousRule` 是原规则；
   - `wps_add_conditional_formatting`：`clearExisting:true` + 拼错的 `ruleType` 调一次，确认既有规则**没被清**；
   - `wps_set_filter_and_sort`：`enableAutoFilter:true` + 越界 `colIndex`，确认筛选**没被打开**；
   - `wps_find_and_replace` 传 `searchQuery: ""` + `replaceText:"X"`：确认**报错且内容一个字符都没变**（宿主侧护栏，不依赖桥接）。
3. **台账更新**（谁写 `capability-backlog.md` 谁改，不在我的写入范围）：
   - CAP-10 状态：WPS 侧已实现（模拟宿主验证 + 待真机）；**MS 侧仍未做**（见 §5）；
   - CAP-41 状态：从「本轮在做（ISS-45）」改为「WPS 侧已实现，待真机读数」——注意 ISS-45 台账文字说的是 `clear_range`，与 `get_style_token` 不是同一件事，建议一并订正。
4. **skill 说明**（`skills/office-agent-bridge/**`）：建议补三句——「写完下拉/范围校验后用 `include:["validationViolations"]` 自检存量数据」；
   「做看板前先 `get_style_token`，读不到的字段是 `null`，不要当默认值用；`fonts` 是启发式层级」；
   「`set_data_validation` 现在先校验后写入、写后读回核对」。**不在我的写入范围**。
5. **建议把 harness 用例并入 `tests/addon.test.ts`**：本轮的 89 项断言跑在临时文件 `/tmp/oab-excel-caps-check.mjs`（随会话消失）。
   该文件按 `scripts/build-wps-addon.mjs` 的同序拼接 + 假宿主，重建产物后可直接移交为正式回归（`tests/` 不在我的写入范围）。
6. 注意：我早期的改动已被他人提交 `cf4ba8d`（"feat(CAP-09)…"）以 `git add -A` 的方式一并扫入；后续修复（§2.2/§2.3、报告）仍在工作区未提交。如需拆分提交，以当前工作区为准。

---

## 10. 附：模拟宿主验证输出原文

harness：`/tmp/oab-excel-caps-check.mjs`（673 行，md5 前 12 位 `14fd89d498b5`）
运行：`node /tmp/oab-excel-caps-check.mjs "/Users/Python/Office Agent Bridge"` → 退出码 0

```text

=== CAP-10 数据验证违规定位（模拟宿主）===
  ✓ 调用成功返回
  ✓ 返回 validationCheck
  ✓ 扫描到 7 个带校验单元格
  ✓ 空单元格计入 skippedBlankCells（不当违规）
  ✓ 违规定位到 F5(列表越界) 与 D5(超上限)
  ✓ D6 非数值 → 记违规并说明原因
  ✓ F5 违规明细含值/规则/期望
  ✓ F5 明细带允许值清单
  ✓ 自定义公式规则进 unevaluated（不算通过）
  ✓ 规则清单含期望文案与命中格数
  ✓ 规则清单标出命中单元格
  ✓ message 汇总有效
  ✓ 未请求样式字段时不返回伪造样式
  ✓ maxCells 同时约束违规扫描并如实报截断
  ✓ 截断时写入 warnings
  ✓ 规则读失败 → unevaluated 且明确写原因
  ✓ 读失败不被当成'没有规则'

=== CAP-10 扩展：引用型候选项 / 运算符 ===
  ✓ 引用型候选项（=$H$1:$H$2）能解析并定位 A3
  ✓ greater_than(>10) 定位 B3
  ✓ not_equal(≠0) 定位 C2
  ✓ 引用型规则期望文案写明来源

=== CAP-10 判定器覆盖 ===
  ✓ date(between 2026-01-01~2026-12-31)：序列号 46023 通过、45000 越界
  ✓ date(=DATE(2025,1,1) 大于)：46024 通过、44000 越界
  ✓ text_length(between 2~5)：长度 6 越界、长度 3 通过
  ✓ 文本千分位 '1,234' 按数值比较（1234>1000 通过）
  ✓ IgnoreBlank=false 时空格算越界并说明原因
  ✓ 未知运算符 → unevaluated（不猜通过）
  ✓ Date.parse 与 Excel 序列号换算正确（46023 = 2026-01-01）

=== CAP-10 组合用法 ===
  ✓ include 同时给 validation 与 validationViolations：规则+定位都在
  ✓ cells 模式下也能拿到 validationCheck（区域级，不重复）
  ✓ cells 模式的逐格样式照旧返回

=== set_data_validation 顺序修复 ===
  ✓ 缺 listItems → 报错
  ✓ 缺 listItems → **没有**发生 Validation.Delete（原校验未被清）
  ✓ 缺 listItems → 原规则读回仍在
  ✓ 非法 validationType → 报错且未 Delete
  ✓ 非法 validationType 报错文案说明未改动
  ✓ 非法 operator → 报错且未 Delete
  ✓ between 缺 maxVal → 报错且未 Delete
  ✓ minVal 非数值 → 报错且未 Delete
  ✓ 合法 list 写入成功
  ✓ 返回真实读回 readBack
  ✓ 返回被替换的原规则（可追溯）
  ✓ 提示文案落上（读回一致）
  ✓ 合法 number_range 写入并读回一致
  ✓ less_than 按 schema 口径取 maxVal（旧实现误取 minVal）
  ✓ less_than 未传 maxVal 时回退 minVal 并在 warnings 说明

=== set_data_validation 写后读回：不一致必须暴露 ===
  ✓ 宿主静默不生效 → success=false（不再谎报）
  ✓ 并说明读回不到规则

=== set_data_validation：Add 抛错 → 尽力回滚原规则 ===
  ✓ Add 抛错 → success=false 且报 hostError
  ✓ 回滚失败时如实说明未恢复（不假装没事）
  ✓ 首次 Add 抛错、回滚成功 → 如实报告已恢复
  ✓ 回滚后原规则读回存在（type=3/list）

=== add_conditional_formatting：clearExisting 不得先删后报错 ===
  ✓ 未知 ruleType + clearExisting → 报错且未删除既有条件格式
  ✓ 报错文案说明未做任何修改
  ✓ 不支持的 iconSet + clearExisting → 报错且未删除
  ✓ text_contains 缺 containsText + clearExisting → 报错且未删除
  ✓ formula 缺 formula1 + clearExisting → 报错且未删除
  ✓ between 缺 formula2 + clearExisting → 报错且未删除
  ✓ 原条件格式仍在（未被清）
  ✓ 合法调用：clearExisting 删除旧规则并写入新规则

=== set_filter_and_sort：越界 colIndex 不得先开筛选 ===
  ✓ 越界 colIndex + enableAutoFilter → 报错且未开启筛选
  ✓ 报错文案说明未做任何修改
  ✓ 合法调用仍能开启筛选（回归）

=== find_and_replace：宿主侧空串护栏（不依赖上游）===
  ✓ 空搜索串 → 宿主侧直接报错
  ✓ 被拒后**零改动**（内容未被改坏）
  ✓ 报错文案说明未做任何修改
  ✓ 空格查询不被误拦（口径与桥接一致：只拦空串）
  ✓ 正常替换仍工作

=== CAP-41 get_style_token（模拟宿主）===
  ✓ 调用成功
  ✓ 取样格字体是真实读回（宋体/11）
  ✓ titleFontName 读 A1 真值（微软雅黑）
  ✓ 无填充时 headerBackgroundColor=null（不再编 #1E3A8A）
  ✓ 字体层级 title 取 20pt 那组
  ✓ 字体层级 header 取加粗+有底色的表头组
  ✓ 字体层级 body 取出现最多的正文组
  ✓ 字体层级 caption 取 9pt 最小字号
  ✓ 层级判据写明是启发式
  ✓ 设计普查有真实分组与取样地址
  ✓ observed 配色来自实际普查（含表头填充色）
  ✓ 工作簿调色板探测成功（Workbook.Colors）
  ✓ 主题色探测失败 → 进 unavailable 而不是编造
  ✓ 表格样式读到 TableStyleMedium2
  ✓ 条件格式风格读到规则
  ✓ probes 记录每项探测的成功/失败与原因
  ✓ message 汇总探测成功率
  ✓ 主题色探测成功路径：读到 2 个槽并标明来源

=== CAP-41 反向用例：读不到就不许编 ===
  ✓ Font.Name 读失败 → fontName=null（旧实现会编 '微软雅黑'）
  ✓ 并写入 warnings 说明字段为 null
  ✓ 不存在的取样地址 → 明确报错

=== 回归：既有能力未被破坏 ===
  ✓ include:validation 仍按原样返回规则
  ✓ 默认 include 不返回 validationCheck（不改变旧输出）
  ✓ read_range 仍正常
  ✓ get_sheet_outline 仍正常
  ✓ set_data_validation action=read 仍正常

=========== 结果：94 通过 / 0 失败 ===========
```

> 说明：模拟宿主只用来**证明新代码路径的逻辑与"读不到不谎报"的契约**，不能替代真机读数（§0）。
