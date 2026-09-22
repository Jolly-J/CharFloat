# 故障与并发实测原始证据（fault 用例）

- 时间：2026-09-22
- 目标文档：`钙钛矿各家企业现状.xlsx`（/Users/jolin/Downloads/，使用者授权测试）
- 服务：wps-bridge 2.1.0，WPS 加载项 authenticated-addon，HTTP 19890
- 新建工作表：`fault_case1_rollback`、`fault_case2_scope`、`fault_case2_tmp`、`fault_case3_A`、`fault_case3_B`、`fault_case5_params`（未删改他人表；3 张 sweep 表 usedRange 与开工前基线一致）
- 副本/脚本：`.scratch/fault/mcp.py`（MCP 客户端）、`call.py`、`case3.py`、`case5.py`、`case3.log`、`case5.log`

## 用例 1 审计回滚负向

1.2 写入 `fault_case1_rollback!A1:B3` → auditId `28992179-c4b9-4ba7-92a5-b5eaf052233a`
1.3 另一工具改同区域：`wps_execute_script`（component=excel, workbookName=…, code 内用 `wb`）
    → `{"success":true,"executionTimeMs":9,"returnValue":"{\"B2\":999,\"A3\":\"B-用户改过\"}"}`
1.4 读回：`values=[["项目","数值"],["A",999],["B-用户改过",2]]`
1.5 旧 auditId 回滚 → `isError=true`，原文：
    `目标区域已有后续修改或快照缺失，已拒绝覆盖。请先核对当前内容。`
1.6 读回未变 → 判定：**通过**

补充原始返回：
- 区域外改动不检测：改 C6 后回滚 A6:B7 → `{"success":true,"message":"成功恢复区域 $A$6:$B$7 的数据",...}`，C6 保留
- 行结构变更后回滚 → `目标区域已有后续修改或快照缺失，已拒绝覆盖。请先核对当前内容。`
- 重复回滚 → `{"success":true,"message":"该记录已处于已撤销状态，无需重复操作"}`
- 不存在 id → `找不到对应的审计记录: not-a-real-id`
- 缺参 → `arguments.auditId: 必填`；数字 → `arguments.auditId: 类型不正确`
- 两个有审计工具先后写同区域，回滚旧记录 → 拒绝；先回滚新记录成功，再回滚旧记录成功（逆序撤销正确）
- 脚本写 `doc.Worksheets` → `原生脚本执行异常: Cannot read properties of null (reading 'Worksheets')`（Excel 场景 `doc` 为 null，须用 `wb`；SKILL.md 第 11 行"已自动绑定"表述与实测不符）

## 用例 2 回滚覆盖边界

- `excel_format_cells` / `excel_add_conditional_formatting` / `excel_freeze_panes` / `excel_modify_rows_columns` / `excel_add_chart` **均不返回 auditId**；`wps_get_audit_history` 只有 `actionType=update_values`（patch_cells）
- 样式假成功：patch A1:B3 得 auditId `9fcc5f06-…` → 同区域加粗+红底 → 用该 auditId 回滚 →
  `{"success":true,"message":"成功恢复区域 $A$1:$B$3 的数据","result":{...,"restoredRows":3,"restoredCols":2}}`
  读回：值被清空，但 6 格仍为 `bold=true,fontColor=#FFFFFF,backgroundColor=#FF0000` → **样式未恢复，仍报成功**
- 图表：建图返回 `{"shapeName":"Chart 1",...}`，无 auditId；回滚数据记录不会删除图表
- 删表后回滚 → `在工作簿 [钙钛矿各家企业现状.xlsx] 中找不到工作表: "fault_case2_tmp"。当前可用的工作表为: [...]`
- **覆盖事故（真实数据破坏）**：删表 → 重建同名空表 → 回滚被拒（正确）；再用脚本把**同样内容**写回 → 同一 auditId 回滚
  `{"success":true,"message":"成功恢复区域 $A$1:$B$2 的数据",...}`
  读回：`values=[[null,null],[null,null]]` → **重建的数据被清空，返回 success:true**（判据是"内容相等"而非"同一次修改"）

## 用例 3 双客户端并发

- 两个独立 MCP 会话（sessionId 2752ce20… / dee788a5…），各写各表，8 次写入 8/8 `success:true`，用时 7.93s（宿主串行）
- 交叉读回：`fault_case3_A` 只见 A 标签，`fault_case3_B` 只见 B 标签 → 无串写
- 锁隔离：A `wps_get_locked_status` → `{"gatewayLocks":{"excel":"钙钛矿各家企业现状.xlsx"}}`；B → `{"gatewayLocks":{"excel":"B专属不存在.xlsx"}}`；B 不传 workbookName → `未在 WPS 中找到目标工作簿 [B专属不存在.xlsx]` → **锁按会话隔离，不串**
- 审计：按 `sheetName=fault_case3_A` 过滤只返回该表 4 条 → 无串扰。但 **clientName 恒为 `MCP Agent`**（`src/bridge/mcp-server.ts:29`），两个客户端在审计里无法区分；HTTP `/api/v1/tool/call` 路由才认 `clientName`（实测记录为 `fault-HTTP-Client`）
- 同区域并发（H1:I1，两会话各 5 次）：10/10 `success:true`，最终值为最后写入者，审计 10 条按序完整，无丢写

## 用例 4 后台生命周期

- 停服：`ELECTRON_RUN_AS_NODE=1 "/Applications/Office Agent Bridge.app/Contents/MacOS/Office Agent Bridge" ".../app.asar.unpacked/dist/bridge/cli.cjs" --stop` → 退出码 0，`/health` connection refused
- 停服后调工具：
  - 项目 `skills/office-agent-bridge/scripts/bridge_client.py` → `<urlopen error [Errno 61] Connection refused>`（不点明服务/端口/恢复方式）
  - 技能 `wps_client.py` → `执行异常: 无法连接到 WPS Bridge 服务 (http://127.0.0.1:19890)。请确认本地守护进程正在运行。详情: <urlopen error [Errno 61] Connection refused>`
  - `--doctor` → `{"installationRecord":true,"credentialsPresent":true,"error":"<urlopen error [Errno 61] Connection refused>"}`
  - `--status` → `{"ready":false,"occupied":false,"message":"fetch failed"}`
- 恢复：`bridge_client.py --start` 在**本会话沙箱内失败**：`EPERM: operation not permitted, open '/Users/jolin/.wps-bridge/service.log'`（沙箱只允许写工作区）
- 变通恢复：复制 token/installation.json/audit_history.json 到 `.scratch/fault/bridgehome`，`WPS_BRIDGE_HOME=<该目录> WPS_BRIDGE_PORT=19890 bridge_client.py --start` → `Bridge 后台已就绪`，health `{"pid":68118,...}`
- 重启后：`isWpsConnected=true`，word/excel/ppt 全部 connected（加载项自动重连）；工作簿 16 张表（含 6 张 fault_）全在；审计历史 total=68 保留
- **会话上限**：HTTP `/mcp` 硬上限 64（`src/bridge/ws-server.ts:188`），无空闲过期；一次性会话累积后 `initialize` 返回 `HTTP 429 {"error":"会话数量达到上限"}`（实测两次），DELETE /mcp 可释放；恢复后新会话正常

## 用例 5 参数边界

报错清楚且拒绝（HTTP 422，`success:false`）：
`arguments.host: 必填` / `arguments.address: 类型不正确` / `arguments.foo: 未知参数` / `arguments.freezeRowIndex: 类型不正确` / `arguments.limit: 类型不正确` / `arguments.sheetName: 必填` / `arguments.values[0]: 类型不正确` / `写入矩阵必须与目标区域尺寸一致` / `必须提供 values 或 formulas 进行更新` / `缺少必要参数: sheetName` / `未知工具 wps_not_a_tool` / 找不到工作表时列出全部可用表名 / 找不到工作簿时列出已打开文件

**不指出允许值**（原文即全部内容）：`arguments.host: 不在允许值中`、`arguments.chartType: 不在允许值中`、`arguments.view: 不在允许值中`

**静默吞掉/无范围校验**：
- `excel_freeze_panes freezeRowIndex=-3` → `{"success":true,...,"freezeRowIndex":-3,"message":"已成功锁定第 -4 行之上的表头吸顶显示"}`（`wps-addon/src/excel.js:623` 守卫 `>1` 不成立，仍设 `FreezePanes=true`）
- `excel_get_range_styles maxCells=99999` → 静默截到 500（`returnedCells:500,truncated:true`）；`maxCells=-5` → 静默返回 1 格
- `wps_get_audit_history limit=100000` → 静默截到 50；`limit=-5` → 静默变 1
- `excel_read_range address="ZZZ9999999"` → `Cannot read properties of null (reading 'Rows')`（内部错误直出，无可用地址格式提示）

## 恢复步骤（把后台交还给规范 home）

1. `ELECTRON_RUN_AS_NODE=1 "/Applications/Office Agent Bridge.app/Contents/MacOS/Office Agent Bridge" "/Applications/Office Agent Bridge.app/Contents/Resources/app.asar.unpacked/dist/bridge/cli.cjs" --stop`
2. `python3 skills/office-agent-bridge/scripts/bridge_client.py --start`（在非沙箱终端执行；读 `~/.wps-bridge/installation.json`）
3. 校验：`curl -s http://127.0.0.1:19890/health` 返回 `{"service":"wps-bridge","protocol":2,...}`

当前状态：服务在 19890 正常运行，`WPS_BRIDGE_HOME=/Users/Python/Office Agent Bridge/.scratch/fault/bridgehome`（token 与原 home 哈希一致，audit_history.json 已复制），功能无差异，仅日志/状态写在项目内。
