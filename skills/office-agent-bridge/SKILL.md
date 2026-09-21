---
name: office-agent-bridge
description: 使用本机 Office Agent Bridge 连接已打开的 WPS 或 Microsoft Office 文档，优先处理 Excel 表格；负责启动检查、宿主与能力确认、连接排障和有限修复。适用于 Bridge/MCP 连接失败、实时表格操作及询问 Word/PPT 支持范围。
---

# Office Agent Bridge

Bridge 是用户本机的办公连接服务。推荐实时表格场景，不能将“有原生脚本接口”解释成所有功能都已支持。

## 接入后的第一步

调用 `bridge_get_capabilities`，再按需要调用 `bridge_diagnose`。确认实际平台、服务版本、已连接宿主和验证程度。MCP 连接成功不代表 WPS 加载项或 Microsoft Office 可访问。

- 统一表格工具使用 `excel_*`，显式选择 `host: "wps"` 或 `"microsoft"`。
- `wps_*` 兼容工具仅操作 WPS。Windows Microsoft Excel 结构化适配已实现，实机验收状态以能力返回及验证报告为准；macOS Microsoft Office 目前是原生脚本通道。
- 多文档必须用工作区返回的精确 `workbookName` / `documentName` / `presentationName`，不要猜路径或按相似名称挑文件。
- 没有连接时先修复连接，不要通过离线改另一份文件掩盖连接故障。用户明确要求离线处理时尊重其选择。

## 服务启动与检查

GUI 窗口与后台进程独立。关闭窗口或退出管理器不停止 Bridge；“停止服务”会断开所有 AI。

stdio MCP 会自动启动或复用后台。HTTP MCP 不会凭地址自动启动，须先启动管理器或运行已安装 CLI。

本技能的 `scripts/bridge_client.py` 不依赖第三方 Python 包。先定位当前技能文件所在目录，再执行该目录下的脚本；不要照抄某个开发者的绝对路径。

```sh
python3 <技能目录>/scripts/bridge_client.py --doctor
python3 <技能目录>/scripts/bridge_client.py --start
python3 <技能目录>/scripts/bridge_client.py bridge_get_capabilities '{}'
```

启动仅使用安装时记录的可执行文件和入口。缺少安装记录时，提示用户从已安装的 Office Agent Bridge 应用启动一次；源码环境按项目 README 构建并运行 CLI。不要猜测 `node` 已安装，也不要下载来历不明的运行时。

具体检查、故障分类和修复步骤见 [启动与排障](references/operations.md)。需要介绍支持范围或选择方案时读取 [能力边界](references/capabilities.md)。

## 操作与交付

先读工作区与目标范围，再执行用户要求的修改。保留已有样式，按需用公式计算、使用原生图表和设置格式，不强制把所有表格改成统一模板。

写入后读回关键值和公式。视觉修改需检查真实预览；预览可能改变剪贴板或激活文档，应避免与并发桌面操作争用。

只有 `patch_cells` 的值与公式进入审计回滚。格式、图表、结构、Word/PPT 和任意脚本没有统一回滚。写入成功不等于已保存文件；报告分别说明修改和保存结果。

超时、断线或审计失败时可能已发生修改。先读回，禁止自动重试写入。明确报告失败或部分完成，不把占位文字框称为原生图表。
