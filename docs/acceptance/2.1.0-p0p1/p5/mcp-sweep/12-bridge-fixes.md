# 12 桥接侧修复计划（ISS-46/93/97/77/02/48/49/50/51/59/90/16/03）

> 写区：`src/bridge/**`，排除 `src/bridge/tools/definitions/**` 与 `src/bridge/tools/excel.ts`。
> 不在写区内的改动一律不做，改用等价手段（运行时返回字段、能力暴露、错误文案）。
> 状态标记：⬜ 未开始 / 🔄 进行中 / ✅ 完成 / ⚠️ 只改文案 / ❌ 修不动

## 逐条计划

| 编号 | 文件（拟改） | 改什么 | 状态 |
|---|---|---|---|
| ISS-46 | `gateway/audit.ts`、`audit-store.ts`、`gateway/excel.ts`、`gateway/script.ts` | 回滚判定加**身份判据**：① 目标工作表存在性（`get_sheet_outline`）；② 记录一致性（同区域更晚的 applied 记录 → 拒绝）；③ 引入"未审计改动时间戳"台账，脚本/结构类操作标记后，若发生在记录时间之后 → 拒绝并给显式恢复路径 | ⬜ |
| ISS-93 | `office/normalizer.ts` | 为 12 个走 `default:` 的 `excel_*` 补字段映射（set_data_validation / manage_sheet / manage_rows_and_columns / set_filter_and_sort / create_pivot_table / manage_cell_comments / get_charts / add_chart / find_and_replace / create_sheet / delete_sheet / duplicate_sheet / capture_sheet_preview / update_chart / get_range_styles）；映射不到的字段**显式抛错**，不静默 no-op | ✅ |
| ISS-97 | `office/adapter.ts`、`contracts/host-methods.ts` | COM 回退白名单从整张 `EXCEL_METHODS` 换成按 COM 实际覆盖生成（排除 `update_chart` / `save_workbook`），保留方法名常量便于核对 | ⬜ |
| ISS-77 + ISS-02 | `gateway/excel.ts`、`gateway/lock.ts`、`gateway/locks.ts`、`ws-server.ts`、`tools/diagnostics.ts` | `get_workspace_summary` 拆字段：`openWorkbooks` / `hasOpenWorkbook`（真实打开状态）/ `lockTarget`（锁目标是否存在、是否陈旧）；修正文案与回退顺序；`get_locked_status` 补文件列表 | ⬜ |
| ISS-48 | `gateway/excel.ts`、`audit-store.ts`、`types.ts`、`tools/diagnostics.ts` | 给格式/条件格式/冻结/行列/图表/数据校验等写操作补**审计记录**（无值快照，标记 `rollbackable:false`）；回滚对无快照记录拒绝并说明；能力里暴露可回滚范围清单 | ⬜ |
| ISS-49 | `mcp-server.ts` | MCP 会话从 `initialize` 的 `clientInfo` 取名（SDK `getClientVersion()`），审计 `clientName` 可区分 | ⬜ |
| ISS-50 | `service-client.ts`、`cli.ts`、`ws-server.ts` | 后台不可用类错误统一为"服务名 + 地址 + 恢复动作" | ⬜ |
| ISS-51 | `ws-server.ts` | `/mcp` 会话加空闲回收（可配置超时 + 定期清理）；429 文案给出释放方式（DELETE / 等待回收 / 重启后台） | ⬜ |
| ISS-59 | `ws-server.ts`、`tools/diagnostics.ts` | 桥接侧可查"运行中的是哪一版"：登记连接上报版本 + 磁盘加载项构建指纹（sha256/mtime/size/路径）+ 是否与已连接版本一致 | ⬜ |
| ISS-90 | `ws-server.ts`、`tools/diagnostics.ts` | connected→disconnected 且非正常关闭（1006/心跳超时等）时记录"疑似宿主崩溃"信号 + 恢复指引，`/api/v1/status` 与诊断工具可查 | ⬜ |
| ISS-16 | `gateway/excel.ts` | `save_workbook` 返回体加"工作簿级保存"范围警告（schema 在 definitions，写区外） | ⬜ |
| ISS-03 | `catalog.ts` | 参数校验报错列出**允许参数集**；对 `excel_*`/`wps_*` 域统一容忍 `host`/`workbookName`/`sheetName` 并在响应里回报被忽略的参数 | ⬜ |

## 分层与契约约束

- 协议层（`ws-server.ts` / `mcp-server.ts`）不得新增对 `catalog` 的静态引用，只经注入的 `ToolService`。
- 不修改 `ToolService` 签名（`tests/session-isolation.test.ts` 依赖 `dist/` 产物）。
- 契约层 `contracts/` 保持纯定义。

## 验证

1. `npm run typecheck` → 0 错误；
2. `npm test` → 86/86（新增用例后总数增加，原 86 项必须全绿）；
3. `npm run build:main`；
4. 不启动/重启后台服务。

## 进度记录

（边做边追加：每条修复写入"改动 + 回归测试 + 未实测项"。）
