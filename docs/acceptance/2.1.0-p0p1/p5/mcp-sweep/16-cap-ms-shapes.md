# 16 · CAP-08 / CAP-52：MS 侧矢量绘图补强

- 开始时间：2026-09-22（本轮）
- 工作目录：`/Users/Python/Office Agent Bridge`
- 写区：`office-addon/src/**`、`office-addon/public/taskpane.html`（仅必要时）；本文件
- 只读参考：`src/bridge/**`、`docs/**`、`skills/**`（不改）
- 目标宿主：本机 Microsoft Excel（`工作簿1.xlsx`），Office.js 任务窗格已连接

## 台账原文（`docs/acceptance/2.1.0-p0p1/capability-backlog.md`）

> **CAP-08** | **MS 侧矢量绘图**：几何形状、连接符、SVG、文本框、分组、层级、旋转缩放、形状导图
> 依据：Office.js ExcelApi 1.9 `Shape` 全套（**文档可证，未实测**）
> **CAP-52** | MS 侧**矢量绘图域**整体补齐（与 CAP-08 同源）

完成定义（台账末尾统一口径，缺一不算）：宿主实现 + 工具定义 + 读回能力 + 契约快照差异 + skill 说明 + 真机读数 + 回归。

## 进度

| # | 步骤 | 状态 |
|---|---|---|
| 1 | 静态盘点：现状实现与工具面缺口 | 已完成 |
| 2 | 真机探测：Office.js `Shape` 各 API 在本机可用范围 | 进行中 |
| 3 | 宿主实现：按实测结果补齐 `office-addon/src/excel/shape.js` + RPC 分支 | 待开始 |
| 4 | 读回能力：列表 / 位置尺寸 / 填充 / 层级 | 待开始 |
| 5 | 真机读数：逐项写→读回 | 待开始 |
| 6 | 回归：build / node --check / --check / typecheck / check:claims / test | 待开始 |
| 7 | 报告：`gateway.ts` 注册行 + skill 说明 + 契约快照差异 | 待开始 |

## 1. 静态盘点结论（源码可证）

- **宿主实现已存在但极少**：`office-addon/src/excel/shape.js` 只有三个函数
  `handleInsertImage` / `handleListShapes` / `handleUpdateShape`（移动/尺寸/删除）。
- **这三个函数是"死分支"**：`src/rpc.js` 的 `dispatchExcelTool` 有
  `insert_image` / `list_shapes` / `update_shape` 三个 case，但
  `src/bridge/contracts/host-methods.ts` 的 `EXCEL_METHODS` 不含它们，
  因此 `unifiedExcelTools()` 不会派生出 `excel_insert_image` / `excel_list_shapes` / `excel_update_shape`；
  未加 `wps_` 前缀的裸名也不在 `HANDLERS` 里。
  实测：`{"name":"excel_list_shapes","arguments":{"host":"microsoft"}}` → `未知工具 excel_list_shapes`。
- **`handleListShapes` 读回字段严重不足**：只 load `name/id/type/left/top/width/height`，
  没有填充色、线条、旋转、层级、文本；`handleUpdateShape` 没有旋转、填充、层级、文本。
- **`addLine` / `addSvg` / `addTextBox` / `addGroup` / `setZOrder` / `getAsImage` 完全没实现。**
- 形状层面的"读回"在 MCP 层**没有任何入口**：`list_shapes` 未注册，所以 AI 现在连
  "这张表里有哪些形状"都问不到。

## 2. 真机探测（进行中）

探测通道说明（重要）：本机 MCP 工具面里**没有任何入口**能调用 Office.js 任务窗格的
`run_script` 分支（`excel_*` 路由表未含形状方法；`office_execute_script` 走的是 macOS JXA / Windows COM，
不是 Office.js）。探测采用的方式记录在下面「探测方式」小节，结果逐项列出。

（结果见下方追加章节）
