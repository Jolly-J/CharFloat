# 中间产物：三通道路由覆盖率与跨宿主字段契约核对

生成时间：本轮 MS-vs-WPS 对照分析（只读，未改任何源码/既有文档）
用途：`docs/acceptance/2.1.0-p0p1/p5/mcp-sweep/08-ms-vs-wps.md` 的 §1.4 / §3.5 证据底稿。

## 1. 三通道对 `EXCEL_METHODS`（`src/bridge/contracts/host-methods.ts`，30 项）的覆盖

复现命令（仓库根运行，只读）：

```bash
python3 - <<'PY'
import re
hm=open('src/bridge/contracts/host-methods.ts').read()
block=hm[hm.index('EXCEL_METHODS = ['):hm.index('] as const')]
methods=re.findall(r"'([a-z_]+)'", block)
ps=open('resources/office/excel.ps1').read()
com=set(re.findall(r"'([a-z_]+)'", ps[ps.index('switch($method)'):]))
com |= set(re.findall(r"\$method -eq '([a-z_]+)'", ps))
print("EXCEL_METHODS:", len(methods))
print("COM implements:", len([m for m in methods if m in com]))
print("COM missing   :", [m for m in methods if m not in com])
wd=open('wps-addon/src/dispatch.js').read()
wps=set(re.findall(r'case "([a-z_]+)"', wd))
print("WPS missing   :", [m for m in methods if m not in wps])
oa=open('office-addon/src/rpc.js').read()
off=set(re.findall(r'case "([a-z_]+)"', oa))
print("OfficeJS missing:", [m for m in methods if m not in off])
PY
```

实际输出：

```
EXCEL_METHODS: 30
COM implements: 28
COM missing   : ['update_chart', 'save_workbook']
WPS missing   : ['update_chart']
OfficeJS missing: ['rollback_cells']
```

## 2. RPC 标签计数

```
WPS dispatch.js  : 61 个 case 标签，全部唯一
  ├─ 表格域 30（含 modify_rows_columns / manage_rows_and_columns 两个别名指向同一处理器）
  ├─ 文字 13、演示 9
  └─ 平台/脚本/锁 9（ping、lock×3、eval/eval_code/execute_script、inspect_api、reload）

Office.js rpc.js : 77 个 case 标签，全部唯一 → 43 个 handleXxx 处理器（全部为 Excel）
```

## 3. 跨宿主字段契约不匹配（§3.5 M1–M8）核对路径

三层证据链，逐层读即可复现：

1. **Schema 字段名**：`src/bridge/tools/definitions/excel.ts`
2. **Gateway 原样转发**：`src/bridge/gateway/excel.ts`（各 Handler，把 `args` 字段逐个搬进 `callOffice(method, {...})`）
3. **Normalizer**：`src/bridge/office/normalizer.ts` —— 有 `case` 的走显式映射，**没有 `case` 的走 `default:` 只去 `wps_`/`excel_` 前缀、字段原样透传**
4. **Office.js 读取的字段名**：`office-addon/src/excel/*.js`

抽 schema 属性名的命令：

```bash
cd src/bridge/tools/definitions && python3 - <<'PY'
import re
src=open('excel.ts').read()
for b in re.split(r'\n\s*\{\s*\n\s*type: "function"', src):
    m=re.search(r'name: "([a-z_]+)"', b)
    if not m: continue
    print(m.group(1), "=>", ",".join(re.findall(r'^\s{12}([a-zA-Z_][a-zA-Z0-9_]*):\s*\{', b, re.M)))
PY
```

关键对照（详细后果见报告 §3.5）：

| 工具 | Schema 发送 | Office.js 读取 | 类别 |
|---|---|---|---|
| `set_data_validation` | `validationType`/`listItems`/`operator`/`minVal`/`maxVal`/提示文案 | **只有 `params.rule`** | 静默 no-op + 破坏性（先 `clear()`） |
| `find_and_replace` | `searchQuery`/`replaceText`/`searchRange`/`maxResults` | `text`/`query`/`findText` | 结果错误 + 破坏性（`replaceAll("", …)`） |
| `create_pivot_table` | `sourceRange`/`destCell`/… | `sourceAddress`/`destinationAddress`/`name` | 报错（`getRange(undefined)`） |
| `set_filter_and_sort` | `range`/`enableAutoFilter`/`sortRules` | `address`/`filterColumn`/`criteria`/`sortColumn`/`hasHeader` | 报错（`getRange(undefined)`） |
| `manage_cell_comments` | `add`/`read`/`delete`/`clear_all` | `add`/`list`/`read`/`delete` | `clear_all` 静默 no-op |
| `manage_sheet` | `rename`/`move`/`tab_color`/`protect`/`unprotect` + `color`/`password` | `copy`/`rename`/`hide`/`show`/`color`(读 `tabColor`) | **4 个 action 静默 no-op** |
| `get_charts` | `shapeName`/`chartIndex`/`chartTitle`/`detail` | 只读 `sheetName`+`detail` | 选择器被静默忽略 |
| `add_chart` | 含 `yAxis`/`seriesSettings`/`smoothLine`/`dataRanges`/`hasDataLabels` | 只读 `seriesColors`/`hasLegend`/`position.*` 等 | 5 个参数静默忽略 |

**反面对照（这三条是通的，可作改造模板）**：`search_cells`（normalizer 有显式 `case`，把 `query` 映射成 `text`）、`format_cells`（厚映射分支，同时构造扁平与嵌套字段）、`add_chart` 的 `CHART_TYPE_MAP`。

## 4. 未做的事

- 未运行任何宿主调用（WPS 虽已连接，但本轮为静态分析）。
- 未运行 `npm test` / `build` / `dist` / git 操作（任务约束）。
- 未改写仓库任何源码或既有文档。
- 未产出中间 JSON/日志文件——本轮全部结论均由上述只读命令的 stdout 直接支撑，命令已内联在本文件中以便复现。
