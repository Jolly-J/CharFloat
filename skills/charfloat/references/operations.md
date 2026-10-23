# 启动与排障

## 启动入口

本节仅用于连接异常，不是每个文档任务的前置步骤。stdio MCP 自动启动或复用后台；HTTP 需要先启动管理器或已安装 CLI。关闭 GUI 不停止后台，“停止服务”会影响所有 AI。

本技能 `scripts/bridge_client.py` 不依赖第三方 Python 包。从当前技能目录定位脚本：

```sh
python3 <技能目录>/scripts/bridge_client.py --doctor
python3 <技能目录>/scripts/bridge_client.py --start
```

启动使用安装记录中的可执行文件和入口。记录缺失时从已安装应用启动一次；源码环境按项目 README 构建运行，不猜测或下载运行时。

## 安装布局

macOS 默认 `~/.wps-bridge`；Windows 默认 `%LOCALAPPDATA%/WPSBridge`。`WPS_BRIDGE_HOME`、`WPS_BRIDGE_PORT` 可覆盖，必须与客户端安装记录一致。

- `installation.json`：已安装可执行程序、CLI 入口、资源路径。
- `token`：本机连接凭据。不要在聊天、日志或共享配置中打印它。
- `service.json`：上次服务实例记录，仅供线索；不能据此直接杀进程。
- `service.log`：后台日志。
- `audit_history.json`：单元格修改记录。

CLI 支持 `--start`、`--status`、`--doctor`、`--stop`、`--repair-addon`。无参数是 stdio MCP。GUI/stdio 共享后台，退出其中一个不应切断其他连接。

## 故障处理顺序

1. `--doctor` 查看安装路径是否存在、服务身份与协议是否匹配、认证是否通过。
2. 未启动：用安装记录启动，等待服务健康检查，失败查看日志。一次启动失败后先分析，不反复创建后台进程。
3. 端口占用或服务身份不符：报告冲突，提示关闭旧版 Bridge；不要杀掉未知 PID，不自动改用户全部 MCP 配置。
4. 凭据不匹配：检查当前运行目录、端口和安装记录是否一致。不要盲目轮换 token，这会断开已安装加载项。
5. WPS 未连接：打开对应的表格、文字或演示组件，检查加载项是否加载；必要时检查版本并重新部署本项目加载项（会自动备份并解除 WPS 误阻断配置，保留其他插件索引）。修复不关闭 WPS。
6. Office COM 无法附着：确认 Windows 桌面 Office 已启动、文档已打开，并与 Bridge 在同一用户及权限会话下。先返回原始错误，不自动以管理员身份重启。
7. 文档找不到：重新列举已打开文档并精确选择。目标名称不唯一时停止，不按模糊匹配继续。
8. 写入超时：结果未知。读取目标范围或对象确认是否已生效，再决定下一步。

## 有限修复

用户要求修复 Bridge 连接时，可以启动服务、重新部署本项目加载项和修复明确选中的客户端条目。先说明具体故障与拟修改项；已经授权的同范围修复不重复索要授权。解析失败的外部配置保持原样，报告文件位置。

GUI “安装 / 修复加载项”或 CLI `--repair-addon` 会更新本项目文件并合并 XML 索引，保存备份。安装后需用户在合适时机重启对应 WPS 组件；未保存文档不能自动关闭。

停止后台会影响所有连接，只有用户要求停止或为其明确授权的维护操作时使用。不得拿 `--stop` 当作普通诊断步骤。

## HTTP 接入

只连接本机 `http://127.0.0.1:<port>/mcp`。从当前运行目录的 `token` 文件读取凭据，放入 `Authorization: Bearer ...` 请求头。加载项通过单独的认证 WebSocket 通道连接。不要将服务暴露到公网或把本机地址直接交给云端连接器当作可访问地址。

Python helper 默认每次生成独立 sessionId；多步 CLI 任务可显式传 `--session-id` 复用同一任务上下文，不同任务隔离，仍显式传目标文档名。

## 批注：两种批注不是一回事，别用错

宿主上有**两套**批注，名字像、API 像，**但一套能用、一套是空壳**。
用之前先分清楚，否则会得出"批注功能坏了"的错误结论。

| | 传统批注 `Comment` | 线程批注 `CommentThreaded`（新版） |
|---|---|---|
| 长相 | 红三角 + 便签 | 右侧面板 + 回复串 |
| **Excel 侧** | ✅ **完整可用** | ❌ **空壳**：正文不存、回复文本恒为 `default`、`Resolved` 写不进 |
| **Word 侧** | ✅ 可用 | ✅ 可用（`doc.Comments`，是另一套 API） |

### 表格侧怎么用（已真机验证 6/6）

`excel_manage_cell_comments`（`wps_*` 与 `excel_*` 同名二选一）：

```jsonc
// 加/改批注（同一格再调一次就是更新）
{ "action": "add", "address": "D10", "text": "审核说明：8月良率低于 90%，经核实为设备温控报警" }

// 读回（address 可以给一片区域，如 "D10:D20"，不是只能一格）
{ "action": "read", "address": "D10:D20" }

{ "action": "delete", "address": "D10" }        // 删单个
{ "action": "clear_all" }                        // 清空全表
```

**读回是判据**：`action: "read"` 拿到的 `comments[].text` 才是真的写进去了。
别只看 `add` 返回 `success` —— 本仓库整治过一整批"返回成功但没生效"的缺陷。

`author` 字段由**宿主自动填成当前 WPS 用户**，传了也不改作者（只会拼进正文）。

### 文字侧怎么用（已真机验证 4/4）

`wps_word_review_and_comments`：

| action | 说明 |
|---|---|
| `add_comment` | 插入批注（返回值带**读回的正文与作者**） |
| `list_comments` | 列出批注（含 `done` 是否已解决） |
| `resolve` / `reopen` | **标记已解决 / 重新打开**（`commentIndex` 从 1 开始） |
| `reply` | ⚠️ **本机 WPS 文字不可用**：`Replies.Add(Range,文本)` 返回对象但回复数不增加，工具会**如实报失败**并给替代做法 |

### 做不到的两件事（别绕圈子）

- **表格侧的线程批注**：宿主 API 是空壳。要讨论串请改用**传统批注**。
- **两种宿主上的"批注回复"**：都是空壳。需要"回复"语义时，
  把补充内容**追加到批注正文**里（如"【补充】…"），或人工在 WPS 里回复。

### 让批注"看得见"

默认不显示批注，只有**红三角标记**。要让对方一眼看到：
鼠标移到格子上停一下会弹框；或右键单元格 →「显示/隐藏批注」常显。
**演示或验收时先说明这一点**，否则对方会说"我没看到"。

## WPS / Office 侧栏与连接

WPS 宿主入口负责连接，侧栏只是界面；正常收起侧栏不会断线。需要查看状态时，从 WPS“开始”功能区打开“字浮侧栏”。

先启动字浮后台，再打开 Microsoft Excel 工作簿。首次使用每个工作簿时，手动启动一次“字浮 CharFloat”加载项，等待“后台连接已启用”提示。支持共享运行时的 Excel 随后可收起侧栏；该工作簿再次打开时自动连接。其他工作簿仍需首次启用一次。不支持该能力或启用失败时，以侧栏提示为准；不要承诺所有新文档都能零点击自动加载。

读取真实连接状态确认是否就绪；看到侧栏关闭不能直接判断已断线。没有连接时先确认字浮后台和当前工作簿的首次启用状态，不能反复启动多个 CLI 服务。

多个 Excel 工作簿同时后台连接时，用状态中的 `officeWorkbooks` 选定目标，并显式传 `workbookName`；同名文件可传完整文档路径。没有唯一目标时请求会被拒绝，不要去掉目标参数或改走原生通道猜测。

## CLI 按需查询与证据输出

以下命令不启动服务；需要已有 Bridge。保留原有 `工具名 'JSON'` 调用。

```sh
python3 <技能目录>/scripts/bridge_client.py --list-tools wps_ppt_
python3 <技能目录>/scripts/bridge_client.py --schema wps_ppt_get_slide_shapes
python3 <技能目录>/scripts/bridge_client.py wps_ppt_get_slide_shapes --args-file target.json --output shapes.json --session-id task-demo-001
python3 <技能目录>/scripts/bridge_client.py wps_ppt_capture_slide_preview --args-file target.json --output preview.json --preview-output preview.png --session-id task-demo-001
```

参数文件为 UTF-8 JSON 对象，如 `{"presentationName":"示例.pptx","slideIndex":1}`。结果和图片文件拒绝覆盖，父目录需已存在。旧位置参数调用保留原响应及 imageBase64；仅显式 `--preview-output` 才解码 PNG 并省略 base64；预览保存后仍需实际看图。工具列表仅输出名称/描述，`--schema` 返回单个工具真实定义，使用现有 `/api/v1/mcp-tools` 查询端点。

PNG 输出只做 base64 与 PNG 签名检查，不等于完整图像有效性验收。宿主调用返回后（包括失败、超时）本地写图/写JSON失败时，返回 phase=local_output、hostCallReturnedSuccess（成功/失败）、doNotReplay=true，并在可用 stderr 保留原执行状态、部分完成ID及安全布局读回，诊断写到 stderr；不要重放已执行命令。

失败输出保留服务提供的字段（包括 errorDetails 中的 kind/executed，以及经过协议白名单筛选的 partial 对象ID/计数，以及布局失败的数值 applied/before/after 与分类告警）；服务未给出的 executed 标为 unknown，不猜测是否已执行。超时/断线不自动重试，先读回目标。默认每次调用使用独立 sessionId；多步任务可显式传同一个 `--session-id task-demo-001` 复用会话锁，不同任务使用不同标识，勿用全局固定值。查询不创建锁，仍从同一 Bridge 查询 schema；执行须显式携带 presentationName。不能把这些 CLI 选项传给主应用生命周期 CLI。
