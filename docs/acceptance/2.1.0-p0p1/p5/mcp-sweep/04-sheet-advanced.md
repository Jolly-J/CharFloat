# 表格高级能力实测（adv_ 工作表）

- 候选版本：2.1.0
- 宿主：WPS 表格（macOS，authenticated-addon，protocol 2）
- 工作簿：`钙钛矿各家企业现状.xlsx`（`/Users/jolin/Downloads/钙钛矿各家企业现状.xlsx`）
- 自建工作表：`adv_高级能力实测`、`adv_透视输出`（未改动任何他人表、未删除表）
- 执行时间：本轮会话，测试数据见各节
- 调用方式：HTTP `POST /api/v1/tool/call`（凭据 `~/.wps-bridge/token`），中间脚本在 `.scratch/sheet3/`
- 每次调用均显式传 `workbookName` + `sheetName` + `host:"wps"`

## 结果汇总

| # | 能力 | 结论 | 关键证据 |
|---|---|---|---|
| 1 | 透视表 | ⚠️ 部分生效 | 建表+聚合值正确，但建表瞬间读回是空骨架，需 `PivotCache.Refresh()` 才出数；`destSheetName` 必须是已存在的工作表 |
| 2 | 排序 | ❌ **success 但完全没生效**（静默 no-op） | 原生 `sheet.Sort` 同一区域同一规则排序成功，工具返回 success 但数据逐行不变 |
| 2 | 筛选 | ✅ 生效 | `AutoFilterMode=true`、`AutoFilter.Range=$A$3:$G$15`、自定义筛选后不匹配行 `Hidden=true` |
| 3 | 条件格式 | ✅ 五种规则全部生效 | 色阶/数据条/阈值/图标集/公式规则，脚本读 `FormatConditions` + 截图双证 |
| 4 | 查找替换 | ✅ 文本生效 / ❌ 公式不支持 | 文本 5 处替换读回一致；公式片段 `*2→*3` 返回 `totalFound:0`（不报错、不生效） |
| 5 | 行列操作 | ✅ 尺寸/插入/删除/列隐藏生效；❌ **行隐藏 success 但没生效** | 行高 34pt、列宽 22、插入行/列、删除行/列均读回正确；`action:"hide"` 行时 `Rows.Item(n).Hidden` 仍为 false |
| 6 | 单元格批注 | ⚠️ 增删改+读生效，**作者参数无效** | 宿主 `Comment.Author` 恒为 `jolin`，传入的"实测员A"被拼进正文 |
| 7 | 大数据量 | ✅ 生效 | 1000 格 0.10s、5000 格 0.34s（含 HTTP 往返），逐格校验 1000/1000、5000/5000 非空且值正确 |
| 8 | 合并单元格 | ✅ 生效 | 跨列合并/取消合并，`merged`/`mergeArea` 读回一致 |

## 逐项明细

### 1. 透视表 ⚠️ 部分生效

**建表调用**
```json
{"tool":"excel_create_pivot_table","sourceRange":"A129:J629","sourceSheetName":"adv_高级能力实测",
 "destCell":"B4","destSheetName":"adv_透视输出","rowFields":["公司"],"columnFields":["状态"],
 "dataFields":[{"fieldName":"产能","summaryFunction":"sum","caption":"产能合计"},
               {"fieldName":"良率","summaryFunction":"average","caption":"平均良率"}]}
```
返回：`success:true`，`pivotTableName:"PivotTable_1790063778192"`，`rowCount:1 colCount:1 dataCount:2`，0.03s。

**第一次读回（建表后立即）：空骨架**
`excel_read_range B4:H9` → `["公司","S0","S1","S2","S3","S4","总计"]`，行标签 `批量0001…批量0009` 都在，但**所有交叉单元格为 `null`**，且看不到两个数据字段。`PivotCaches().Item(1).RecordCount` 返回 **1**（源数据实际 500 行）。

**刷新后读回：有数**
脚本执行 `wb.PivotCaches().Item(1).Refresh()`（以及 `pt.RefreshTable()`）后：
```
        S1     S2     S3     S4     S0    总计
批量0001        3     0.2                       3    0.2
批量0002               6     0.4                6    0.4
批量0005 15    1                                 15   1
```
行标签、列标签、求和与平均全部正确（如批量0002 路线码 2 → 产能 `2*3=6`、良率 `0.2` 平均）。

**结构读回**（脚本 `sheet.PivotTables(1)`）
```
name=PivotTable_1790063778192  range=$B$4:$N$507
rowFields=1(公司)  colFields=2(状态、值)  dataFields=2
data1=求和项:产能|srcName=产能   data2=求和项:良率|srcName=良率
sourceData==adv_高级能力实测!R129C1:R629C10
pivotFields=序号/公司/区域码/路线码/效率/良率/产能/状态/评分/备注/值
```
判定：**结构、维度、聚合都真的建成了**，属于"生效但返回信息不足 + 需手动刷新"。

**能力缺口与坑**
1. `destSheetName` 必须是**已存在**的工作表：直接传新表名报
   `在工作簿 [钙钛矿各家企业现状.xlsx] 中找不到工作表: "adv_透视输出"。…若需要创建新表，请先调用 wps_create_sheet。`
   —— 必须先 `excel_create_sheet` 再建透视表（工具文档未写这一点）。
2. 工具**没有 refresh 参数**，也没有"读透视表"工具。建表后立刻读回是空骨架，`success:true` 会让人误判失败/成功；实机需要 `PivotCache.Refresh()` 或 `PivotTable.RefreshTable()`。
3. `PivotCache.RecordCount` 恒为 `1`（明明有 500 行），**不能当作数据量校验依据**，属于宿主侧误报。
4. WPS 的透视集合是**可调用成员**不是属性：`sheet.PivotTables(1)` 正确，`sheet.PivotTables.Item(1)` 抛 `sheet.PivotTables.Item is not a function`；`pt.RowFields` / `pt.DataFields` 同样是函数，要写成 `pt.RowFields().Count`。同样的坑也出现在 `sheet.PivotTables.Count = undefined`。

### 2. 排序与筛选

**排序：❌ 返回 success 但完全没生效（静默 no-op，与上一轮一致，本轮给出根因）**

受控实验（同一区域 A4:G15，先重置为原始乱序，再分别用两条通道排序）：

| 通道 | 返回 | 读回 A4:G15 实际顺序 |
|---|---|---|
| `excel_set_filter_and_sort` `{range:"A4:G15",sortRules:[{colIndex:1,order:"asc"},{colIndex:6,order:"desc"}]}` | `success:true`，`sortedRuleCount:2`，message"已成功在 [adv_高级能力实测] A4:G15 应用筛选与排序" | **与排序前逐行完全相同**（协鑫、极电、纤纳、脉络、通威、隆基…原顺序） |
| `wps_execute_script` → `sheet.Sort.SortFields` + `SetRange` + `Header=2` + `Apply()` | `success:true` | 正确：华北→华东→华南→西南，组内良率 95.1/87.1/85.4/83.2 降序 |

已排除的干扰项：
- `colIndex` 口径：用 `{colIndex:2}`、`{colIndex:7}`、`{colIndex:1}` 三种都试过，全部无效，因此不是相对列号填错。
- 区域含表头与否：`A3:G15`（含表头）与 `A4:G15`（纯数据）都试过，都无效。

**根因（读源码 + 宿主探测，证据充分）**
1. `wps-addon/src/excel.js:1687-1700` 走的是 `const sort = targetRange.Sort; sort.SortFields.Clear(); sort.SortFields.Add(...); sort.SetRange(...); sort.Header=1; sort.Apply()`。
2. 但**本机 WPS 的 `Range.Sort` 是方法不是对象**：探测显示 `rng.Sort` 的 `Object.keys` 为空，`rng.Sort.SortFields` 为 `undefined`，脚本直接抛 `Cannot read properties of undefined (reading 'Clear')` / `(reading 'Add')`。
3. 该异常被 `try/catch` 吞掉，只记进 `attempts`；紧接着的 `Range.Sort(key1, order1, key2, order2, undefined, undefined, 1)` 同样静默不改数据（矩阵原样返回）。
4. 于是本该兜底的"读回校验 + 报错"没能拦住：`isSortedByRules`（`wps-addon/src/excel.js:1627`）**从 `i = 2` 开始比较、且 `for (i = 2; i < matrix.length; i++)` 少了最后一行**，等于整行整体错位一格；只要当前数据在"错位一格"的口径下看着有序，函数就 `return true`，`throw new Error('排序未生效…')` 被跳过，最终照常返回 success。

**可用的排序通道**：`sheet.Sort`（工作表级，不是 `Range.Sort`）→ `SortFields.Clear()` / `SortFields.Add(range.Columns.Item(n), 0, 1|2)` / `SetRange(range)` / `Header = 2`（无表头）/ `Apply()`。注意 `Range.Sort(key, order, …, 0|1|2)` 的各种变体在本机全部无效。

**筛选：✅ 生效（含范围扩展告警）**

| 操作 | 调用 | 读回证据 | 判定 |
|---|---|---|---|
| 开启自动筛选 | `{range:"A3:G15",enableAutoFilter:true}` | `sheet.AutoFilterMode=true`、`sheet.AutoFilter.Range.Address()="$A$3:$G$15"` | 生效 |
| 自定义条件筛选（隐藏不匹配行） | 脚本 `rng.AutoFilter(7,"已投产",1)`（第 7 列＝项目状态） | 各数据行 `Rows.Item(n).Hidden`：`中试/研发/建设中` 行为 `true`，`已投产` 行为 `false` | 生效 |
| 关闭筛选 | `{range:"A3:G15",enableAutoFilter:false}` | `AutoFilterMode` 回 false，全部行 `Hidden=false` | 生效 |

- 工具对筛选只回 `enableAutoFilter`、`appliedFilterRange`、`warnings`，**没有"筛选隐藏了哪些行"的读回**；"自定义筛选"（按文本/枚举值筛选具体值）也没有参数入口，只能用脚本 `Range.AutoFilter(字段, 条件, 操作符)`。skill 未提及这一点。

> 说明：为验证"排序真的改变了顺序"，本轮在 A4:G15 做过多次重置与排序，`adv_高级能力实测` 第 4-15 行的企业顺序已被原生排序改为按区域归集；第 3 行表头已恢复。这些都属于本次自建表，不影响他人表。

### 3. 条件格式 ✅ 五种规则全部生效

| 规则 | 调用 | 读回证据（脚本读 `FormatConditions`） | 判定 |
|---|---|---|---|
| 色阶 | 工具 `ruleType:"color_scale"`，E4:E15 | `type=3`（xlColorScale），`ColorScaleCriteria.Count=2`，priority=3 | 生效 |
| 数据条 | 工具 `ruleType:"data_bar"`，D4:D15 | `type=4`（xlDatabar），`BarColor` 为对象，priority=2 | 生效 |
| 单元格值阈值 | 工具 `ruleType:"cell_value"`,`operator:"less_than"`,`formula1:"0.9"`，F4:F15 | `type=1`，`Operator=6`（xlLessThan），`Formula1="=0.9"`，`Font.Color=1776537` | 生效 |
| 图标集 | **无专用工具**，脚本 `Range.FormatConditions.AddIconSetCondition()` + `IconSet = wb.IconSets.Item(5)` | `type=6`，`IconSet.ID=5`，`IconCriteria.Count=3` | 生效（仅脚本通道） |
| 公式规则 | **无专用工具**，脚本 `FormatConditions.Add(2, 0, '=E4>AVERAGE($E$4:$E$15)')` | `type=2`，`Formula1="=I7>AVERAGE($E$4:$E$15)"`（宿主按区域基准单元格重写相对引用），`Font.Color=3046706` | 生效（仅脚本通道） |

**视觉确认（`excel_capture_sheet_preview` A1:G15）**：数据条按产能长短成比例、效率列呈红→绿热力渐变、良率列图标集显示三色符号、效率高于均值的单元格字体变绿、良率<90 的文字为深红、A1:G1 与 A2:G2 合并居中正常。

工具返回值只回显 `ruleType` 与地址（如"已成功在 $F$4:$F$15 应用 cell_value 条件格式"），**不含规则细节，也没有读取/清除接口**，因此本次全部靠 `wps_execute_script` 读原生 `FormatConditions` 核验。

**能力缺口（MCP/skill 层面）**：`excel_add_conditional_formatting` 的 `ruleType` 只枚举 `cell_value | data_bar | color_scale`，**图标集与公式规则没有参数入口**（`formula1/formula2` 只在 `ruleType:"cell_value"` 时有意义，不能当公式规则用）；也没有"读取条件格式"和"按规则清除"的工具（`clearExisting` 只能对同区域先清空）。skill 未说明这两项要走 `wps_execute_script`，只能靠 `native-scripting.md` 的一般原则自行推断。

### 4. 查找替换：文本 ✅ / 公式 ❌

| 场景 | 调用 | 结果 | 判定 |
|---|---|---|---|
| 批量替换文本 | `{searchQuery:"已投产",replaceText:"量产中",searchRange:"G4:G15"}` | `totalFound:5, replacedCount:5`；读回 G4:G15 → 5 处变 `量产中`，其余 `中试/研发/建设中` 不变 | 生效 |
| 纯查找 | `{searchQuery:"中试",searchRange:"G4:G15"}` | `totalFound:3, replacedCount:0`，不传 `replaceText` 即只查不改 | 生效 |
| 区分大小写 | `{searchQuery:"alpha",matchCase:true,searchRange:"I4:I6"}`（I4=Alpha/I5=alpha/I6=ALPHA） | `totalFound:1`（只中 I5） | 生效 |
| 不区分大小写 | 同上 `matchCase:false` | `totalFound:3` | 生效 |
| 全字匹配 | `{searchQuery:"叠层",matchEntireCell:true,searchRange:"C4:C15"}` | `totalFound:7` | 生效 |
| **替换公式片段** | `{searchQuery:"*2",replaceText:"*3",searchRange:"I8:I10"}`（单元格公式 `=D4*2` 等） | `totalFound:0, replacedCount:0`，公式读回仍为 `=D4*2/=D5*2/=D6*2` | ❌ **不支持** |
| 无命中 | `{searchQuery:"nonexistent-token-xyz",replaceText:"X"}` | `totalFound:0`，**返回 success:true 且无任何提示** | ⚠️ 静默 |

**结论与缺口**
- `excel_find_and_replace` **只作用于单元格的值/文本，不搜索也不替换公式表达式**。文档说"支持在海量数据中瞬间检索出所有包含特定错误码（如 `#VALUE!`、`#N/A`）的单元格坐标列表"容易让人以为覆盖公式；实际替换公式需要用脚本 `Range.Formula`/`Replace(..., xlFormulas)`。工具对这一能力缺失**没有任何提示或警告**，只回 `totalFound:0`。
- `results[].row/col` 是**相对检索区域的偏移**而非工作表绝对坐标：检索 `G4:G15` 时第一条落在 G4，却报 `row:1,col:1`；检索 `I4:I6` 时命中 I5 报 `row:2,col:1`；检索 `B4:B15` 时命中 B5 报 `row:5,col:1`。调用方无法直接把返回值当坐标用，工具文档也未说明。
- `results[].formula` 字段在文本单元格上恒为 `null`，没有实际信息量。

### 5. 行列操作：尺寸/插入/删除 ✅，**行隐藏 ❌**

| 操作 | 调用 | 读回证据 | 判定 |
|---|---|---|---|
| 设行高 | `{targetType:"row",action:"set_size",index:4,size:34}` | `Range("A4").RowHeight` 由 16.8 → **34** | 生效 |
| 设列宽 | `{targetType:"column",action:"set_size",index:"H",size:22}` | `Range("H1").ColumnWidth` 由 8.54 → **22** | 生效 |
| 隐藏列 | `{targetType:"column",action:"hide",index:"H",count:1}` | `sheet.Columns.Item("H").Hidden = true`，`Range("H1").ColumnWidth = 0` | 生效 |
| 取消隐藏列 | `{action:"unhide",index:"H"}` | `Hidden=false`，列宽回到 22 | 生效 |
| 插入行 | `{targetType:"row",action:"insert",index:18,count:1}` | A17=n1、A18=空、A19=n2（原 A18 下移），`UsedRange` 由 `$A$1:$I$19` → `$A$1:$I$20` | 生效 |
| 删除行 | `{action:"delete",index:18,count:1}` | 回到 A17=n1、A18=n2、A19=n3，`UsedRange` 回 `$A$1:$I$19` | 生效 |
| 插入列 | `{targetType:"column",action:"insert",index:"I",count:1}` | I4 的 `Alpha` 右移到 J4，`UsedRange` → `$A$1:$J$19` | 生效 |
| 删除列 | `{action:"delete",index:"I",count:1}` | J4 → I4，`UsedRange` 回 `$A$1:$I$19` | 生效 |
| **隐藏行** | `{targetType:"row",action:"hide",index:5,count:2}` | `Rows.Item(5).Hidden` / `Rows.Item(6).Hidden` / `Range("A5").Rows.Hidden` **全部仍为 false**；重复 3 次、换另一个工具 `excel_modify_rows_columns` 结果相同 | ❌ **success 但没生效** |

**行隐藏细节**：`excel_manage_rows_and_columns` 与 `excel_modify_rows_columns` 两个工具都返回
`"已成功对工作表 [adv_高级能力实测] 的第 5 行（共 2 项）执行 hide 操作"`，`success:true`，但宿主侧 `Hidden` 完全没变；对照第 8 行单行隐藏也一样无效，而**同一工具的 `column` + `hide` 正常生效**。

**根因（`wps-addon/src/excel.js:665-683`）**：行分支写成 `const rowRange = sheet.Range("5:6"); rowRange.Hidden = true;`，列分支写成 `sheet.Columns.Item(n).Hidden = true`。前者赋值**不抛错也不生效**（本机 WPS 的 `Range.Hidden` 对该写法是静默 no-op），代码里没有任何读回校验，于是照样走到 `success:true`。可用写法：`sheet.Rows.Item(n).Hidden = true` 或 `sheet.Range("5:6").EntireRow.Hidden = true`。

### 6. 单元格批注 ⚠️ 增删改读生效，**作者参数无效**，且区域读取丢失坐标

| 操作 | 调用 | 读回证据 | 判定 |
|---|---|---|---|
| 添加 | `{address:"F4",action:"add",text:"审核说明：良率 92.5%，高于预警线",author:"实测员A"}` | 原生 `Range("F4").Comment` 存在，`Comment.Text()` = `"实测员A:\n审核说明：良率 92.5%，高于预警线"`，`Comment.Author = "jolin"` | 批注写入生效，作者未生效 |
| 添加第二处 | `{address:"F5",…,author:"实测员B"}` | `Comment.Author = "jolin"`，正文被拼上 `"实测员B:"` 前缀 | 同上 |
| 修改 | `{address:"F4",action:"add",text:"改稿：良率复核后修正为 93.1%",author:"实测员C"}` | 正文被整体替换为新文本（旧的"审核说明…"消失） | 修改生效（`add` 即 upsert） |
| 删除 | `{address:"F5",action:"delete"}` | `Range("F5").Comment` 变为 `null`，F4 不受影响 | 生效 |
| 区域读取 | `{address:"F4:G4",action:"read"}` | `comments:[{text,author},{text,author}], totalComments:2`，**两项都没有单元格坐标** | 生效但有缺口 |
| 读取（无批注） | `{address:"A1",action:"read"}` | `totalComments:0` | 生效 |

**问题**
1. **`author` 参数不写宿主作者字段**：宿主 `Comment.Author` 恒为操作系统用户名 `jolin`；传入的 `author` 被当作正文前缀拼接。工具自己的 `read` 也回 `author:"jolin"`，与传入值不一致，说明参数只有"文本前缀"效果。文档写"批注作者签名，默认为 'AI 智能审核'"，实际无法设置真实作者。
   **根因（`wps-addon/src/excel.js:729-747`）**：实现是 `const commentContent = author ? author + ":\n" + text : text; cell.AddComment(commentContent)`，从头到尾没有给 `Comment.Author` 赋值；返回体里的 `author: author || "AI 智能审核"` 只是把入参回显，属于**回显冒充读回**。可用写法：`cell.AddComment(text)` 后显式 `cell.Comment.Author = author`。
2. **`read` 在区域上丢失坐标**：`comments[]` 只有 `text`/`author`，没有 `cell`/`address`。F4:G4 两处批注读回后无法判断哪条属于哪个单元格。
3. **成功消息有拼接 bug**：`"已在单元格 function Address() { [native code] } 成功添加批注"` —— 实现里写的是 `${cell.Address}` 而不是 `${cell.Address()}`，把方法引用拼进了消息（`delete` 的消息同样如此），用户看不到实际单元格地址。

### 7. 大数据量 ✅ 生效，耗时实测

| 规模 | 调用 | 耗时（含 HTTP 往返） | 读回校验 |
|---|---|---|---|
| 100 行 × 10 列 = **1000 单元格** `A21:J120` | `excel_patch_cells` | **0.10 s** | 脚本逐格统计 `nonEmpty=1000`，`rows=100, cols=10`；A21=`1`、B21=`公司001`、A120=`100`、J120=`备注-100`、F70=`85` 全部正确；MCP 抽样 `A21:J22` 与写入值逐格一致 |
| 500 行 × 10 列 = **5000 单元格** `A130:J629` | `excel_patch_cells` | **0.34 s** | `nonEmpty=5000`，`rows=500, cols=500`，末行 J629=`T0500`，中段 B380=`批量0251` 正确；`modifiedCount=5000` |
| 1000 行 × 10 列 = 10000 单元格（未执行，超出本轮需要） | — | — | 5000 格仅 0.34s，线性外推无压力 |

结论：**批量范围写入是本轮最可靠的能力**，1000/5000 格均一次性成功、值精确、无静默降级、无截断；`modifiedCount` 与写入格数一致，可用于校验。耗时几乎全部是宿主写入，HTTP 往返可忽略。注：1000 格那次日志里 `modifiedCount=1000`，diff 数组被返回（含每格新旧值），大数据量下响应体体积值得关注，但本次未造成问题。

### 8. 合并单元格 ✅ 生效

| 操作 | 调用 | 读回证据 | 判定 |
|---|---|---|---|
| 跨列合并 A1:G1 | `excel_format_cells {address:"A1:G1",merge:true,bold:true,fontSize:14,backgroundColor:"#0F172A",fontColor:"#FFFFFF"}` | `excel_get_range_styles` cells 模式：7 格全部 `merged:true, mergeArea:"$A$1:$G$1"` | 生效 |
| 跨列合并 A2:G2 | 同上 | `$A$2` → `merged:true, mergeArea:"$A$2:$G$2"` | 生效 |
| 取消合并 A1:G1 | `{address:"A1:G1",unmerge:true}` | 7 格全部 `merged:false, mergeArea:null` | 生效 |
| 合并后取值 | — | A1 值保留在左上角，B1:G1 读回 `null` | 符合原生语义 |

- 跨行合并未单独复测（同一 `merge` 分支）；`merge:true` 与字体/底色/对齐参数可在**一次调用**内同时生效。
- 注意：`excel_format_cells` 的回包只回显 `appliedStyles`（输入回显），**是否真合并必须靠 `excel_get_range_styles` 的 `merged`/`mergeArea` 读回**，不能只看 success。
- 截图确认：标题行跨列居中无遮挡，副标题左对齐正常。

## 发现的静默失败 / success 但未生效（重点）

| # | 能力 | 现象 | 严重度 |
|---|---|---|---|
| S1 | **多列/单列排序** `excel_set_filter_and_sort` | 返回 `success:true` + `sortedRuleCount:N` + "已成功…应用筛选与排序"，数据逐行不变。根因：`Range.Sort` 在本机 WPS 是方法非对象，`SortFields` 路径抛错被吞，兜底的 `Range.Sort(...)` 也静默无效，且 `isSortedByRules` 校验存在 off-by-one（从 `i=2` 起、漏最后一行）导致校验恒真，本该报错的保护失效 | 高（与上一轮一致，本轮定位到根因） |
| S2 | **隐藏行** `excel_manage_rows_and_columns` / `excel_modify_rows_columns` `action:"hide"` | 返回 `success:true` + "已成功…执行 hide 操作"，`Rows.Item(n).Hidden` 仍为 false；同一调用换成 `targetType:"column"` 则正常。行分支疑似调用了不存在的宿主 API 且未校验 | 高（新发现） |
| S3 | **批注作者** `excel_manage_cell_comments` `author` | 返回 `author:"实测员A"`，宿主 `Comment.Author` 实为 `jolin`，传入值被降级为正文前缀。工具自己的 `read` 回 `author:"jolin"`，前后自相矛盾 | 中（新发现） |
| S4 | **公式查找替换** `excel_find_and_replace` | 替换公式片段返回 `totalFound:0` 且 `success:true`，无任何"不支持公式搜索"的提示，调用方无法区分"没找到"和"不支持" | 中（新发现） |
| S5 | **透视表建表后读回** `excel_create_pivot_table` | 返回 `success:true` + 结构数量，建表瞬间 `excel_read_range` 读回是**全空骨架**（只有行列标签、无聚合值），必须额外 `PivotCache.Refresh()` 才出数；`PivotCache.RecordCount` 恒为 1，数据量校验失效 | 中（新发现） |
| S6 | 批注成功消息 | `"已在单元格 function Address() { [native code] } 成功添加批注"`，实现拼接了方法引用而非调用结果 | 低（新发现，属消息缺陷） |
| S7 | `excel_find_and_replace` 结果坐标 | `results[].row/col` 是区域相对偏移（G4 报 row:1），非工作表绝对坐标，文档未说明，直接当坐标用会错位 | 低（新发现） |

未发现静默失败的能力：条件格式（三种工具规则 + 两种脚本规则全部读回确认）、范围批量写入、合并/取消合并、行高列宽、插入/删除行列、列隐藏/取消隐藏、筛选开关与自定义筛选。

## 环境与限制备注

- 本轮只写 `adv_` 开头的两张自建表；`__MCP验收测试__`、`sweep*`、`fault_*`、`e2e_钙钛矿数据集` 等表未改动，未删除任何工作表。
- 测试期间有其它会话在创建 `e2e_钙钛矿数据集` 工作表（`excel_create_pivot_table` 报错信息里可见），说明工作簿被并发使用；各行操作都以精确 `workbookName`+`sheetName` 定位，未受影响。
- 未调用 `excel_save_workbook`：本轮只做能力验证，落盘与否由调用方决定（所有修改目前仍在宿主内存中）。
