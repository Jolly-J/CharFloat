# P5 全方位 MCP 盘点（第一轮）

执行时间：2026-09-22 15:17–15:24 · 环境：macOS 27.0 arm64 · WPS 12.1.28496 · 加载项 2.1.0 · 后台 pid 34519

## 方法：只给目标，不给步骤

派 4 个并行子代理，每个只收到"**要做出什么成品**"，**不告诉它该用哪个工具**——由它自己从工具清单、`tools/list` 描述与 `skill` 文档判断。这样反推出来的"卡在哪、哪里要猜"，就是说明质量的真实读数。

| 子代理 | 目标成品 | 报告 |
|---|---|---|
| `inventory` | 91 个工具的全景盘点 + 说明质量评估 | [01-inventory.md](01-inventory.md)（507 行，P-01～P-14） |
| `tables` | 一个"拿得出手"的复杂表格 | 见 [issues.md](../../issues.md) §ISS-38～45 |
| `charts` | 一区多类型图表 + 条件格式 | 见 [issues.md](../../issues.md) §ISS-17～26 |
| `shapes` | JS 脚本构建多元素矢量图形 | 见 [issues.md](../../issues.md) §ISS-11～16 |

## 三个真实成品（均已由协调方拉渲染图核对）

| 成品 | 工作表 | 规模 | 截图 |
|---|---|---|---|
| 产业链信息图 | `sweep图形_钙钛矿产业信息图` | **134 个矢量元素**：标题条 + 三栏分组 + 5 段流程带箭头 + 4 应用卡 + 3 技术路线箭形块 + 4 格 KPI 条 + 页脚 | [shapes-infographic-final.png](shapes-infographic-final.png) |
| 产能与效率看板 | `sweep图表_产能与效率看板` | **5 张原生图表**（条形 57 / 折线 4 / 环形 -4120 / 散点 -4169 / 柱状 51）+ 4 KPI 公式卡 + 数据条/色阶/阈值告警 | [charts-dashboard-final.png](charts-dashboard-final.png) |
| 企业跟踪看板 | `sweep表格_钙钛矿企业跟踪看板` | `A1:N34`：双行表头 + 4 KPI 公式卡 + 20 家企业 + 4 列动态公式 + 合计行 + 3 结论卡 + 冻结 + 筛选 + 下拉 + 4 条条件格式 + 4 条批注 | [tables-dashboard-final.png](tables-dashboard-final.png) |

**结论**：把工具说明与 skill 文档交给一个 AI，**它能做出真正可交付的成果**——前提是愿意付出探测成本（`shapes` 走了 6 轮探测 + 2 次全量重绘，`charts` 迭代 6 轮，`tables` 每踩一个静默失败就得换脚本路径）。

## 本轮最典型的失败模式

**"返回 `success` 但什么也没发生"**，命中条目：

| 条目 | 表现 |
|---|---|
| ISS-38 | `set_filter_and_sort` → `success` + `sortedRuleCount:1`，数据没排序 |
| ISS-39 | `auto_fit_columns` → `success` + `results:[]`，列宽不变 |
| ISS-17 | `chartType:'scatter'` → `success`，实际建的是柱状图 |
| ISS-29 | `search_cells` 搜公式 → **0 命中 + `success:true`** |
| ISS-11 | 对齐传字符串 `'center'` → 赋值不报错，静默变左对齐 |
| ISS-42 | `format_cells` 的 `merge` 失败被吞 → 仍返回 `success` |

另有反过来的：**说明写得越绝对，越容易是没实现的**——`add_chart` 的 `cellRange` 被描述为"100% 完美的行级锁定、杜绝像素漂移"，而 WPS 宿主根本不读这个字段（ISS-18）。

## 未验证，不要当成通过

- `inventory` 只读实测 **11 个**工具；**70 个写工具的真实行为仍未验证**（未验 ≠ 可用）。
- 三个成品都**没做关闭重开后的持久化复验**；P5.6 的落盘核验只对 `__MCP验收测试__` 做过一次。
- Microsoft Office 全通道未连接、未实机验收。
- Word 未连接；PPT 已连接但没有打开的演示文稿。

## 第二轮待启动

Word / PPT 矩阵、故障与并发用例、Microsoft 通道——需使用者先打开对应测试文档并在 WPS 文字 / Microsoft Excel 中加载授权加载项。新问题编号接 **ISS-46** 起。
