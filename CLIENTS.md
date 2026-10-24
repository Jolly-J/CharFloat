# AI 客户端接入与验证

字浮只合并自己的 MCP 条目，写前备份，写后读回。其他服务、用户参数和停用设置保留。配置无法解析、存在歧义或被其他进程同时修改时停止，不覆盖原文件。界面“已保存”表示配置与当前字浮一致；实际加载和授权仍以 AI 客户端为准。

## 两端路径

`H` 表示当前用户主目录：macOS 的 `$HOME`、Windows 的 `%USERPROFILE%`。`A` 在 macOS 是 `H/Library/Application Support`，在 Windows 是 `%APPDATA%`（缺省 `H/AppData/Roaming`）。下表使用 `/` 便于阅读，Windows 运行时使用 Windows 路径规则。应用缓存目录、安装目录与 MCP 用户配置目录不能混用。

| 客户端与官方依据 | MCP 用户配置 | Skill 入口 |
|---|---|---|
| [WorkBuddy](https://www.codebuddy.cn/docs/workbuddy/From-Beginner-to-Expert-Guide/Function-Description/MCP-Guide) | `H/.workbuddy/mcp.json`，`mcpServers` | `H/.workbuddy/skills`；链接加载须在客户端复核 |
| [Qwen Code](https://github.com/QwenLM/qwen-code/blob/main/docs/users/features/mcp.md) | `H/.qwen/settings.json`，`mcpServers` | `H/.qwen/skills` |
| [千问办公](https://help.aliyun.com/zh/qwenwork/desktop-connectors) | 应用内“扩展 → 连接器 → 添加” | 应用内管理；与 Qwen Code 分开 |
| [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness/blob/master/docs/user/guide/mcp-memory.md) | `$DSH_HOME/cordis.patch.yml`，默认 `H/.dsh`；YAML 插件补丁 | `$DSH_HOME/skills` |
| [Kimi](https://www.kimi.com/code/docs/en/kimi-code-desktop/settings-and-extensions.html) | `$KIMI_CODE_HOME/mcp.json`，默认 `H/.kimi-code` | `$KIMI_CODE_HOME/skills` |
| [TRAE](https://docs.trae.ai/ide/add-mcp-servers) | 设置 → MCP → 手动添加；官方未明确全局文件的稳定绝对路径 | 应用内管理 |
| [通义灵码 / Qoder](https://docs.qoder.com/zh/user-guide/chat/model-context-protocol) | 应用内 MCP 配置 | 应用内管理 |
| 豆包工作 | 应用内本地连接器；历史配置文件不作为连接成功证据 | 应用内管理 |
| [Claude Code](https://code.claude.com/docs/en/mcp) | `H/.claude.json` 顶层 `mcpServers` | [H/.claude/skills](https://code.claude.com/docs/en/skills) |
| [Codex](https://learn.chatgpt.com/docs/extend/mcp?surface=cli) | `$CODEX_HOME/config.toml`，默认 `H/.codex`；`mcp_servers` | [H/.agents/skills](https://learn.chatgpt.com/docs/build-skills) |
| [Claude Desktop](https://py.sdk.modelcontextprotocol.io/get-started/real-host/) | `A/Claude/claude_desktop_config.json` | 不自动写入本地 Skill 目录 |
| [Cursor](https://cursor.com/docs/mcp) | `H/.cursor/mcp.json` | [H/.cursor/skills](https://cursor.com/docs/skills) |
| [Windsurf / Devin](https://docs.devin.ai/cli/extensibility/mcp/configuration) | macOS：`$XDG_CONFIG_HOME/devin/mcp_config.json`，默认 `H/.config`；Windows：`A/devin/mcp_config.json`；保留已存在的 `H/.codeium/windsurf` 入口 | 当前配置目录的 `skills`；仅旧目录存在时保留旧入口 |
| [Gemini CLI](https://geminicli.com/docs/tools/mcp-server/) | `H/.gemini/settings.json`；另只读检查 `mcp-server-enablement.json` | [H/.gemini/skills](https://geminicli.com/docs/cli/skills/) |
| [Antigravity](https://antigravity.google/docs/mcp) | `H/.gemini/config/mcp_config.json`；保留已有旧版入口 | [H/.gemini/config/skills](https://antigravity.google/docs/skills/)；已存在 CLI 时也链接 CLI 入口 |
| [OpenCode](https://opencode.ai/docs/config/) | 两端都使用 `$XDG_CONFIG_HOME/opencode/opencode.json[c]`，默认 `H/.config`；支持 `OPENCODE_CONFIG` | 同一配置目录的 `skills` |
| [VS Code / Copilot](https://code.visualstudio.com/docs/agents/reference/mcp-configuration) | 原生 `A/Code/User/mcp.json` 用 `servers`；已存在的 `$COPILOT_HOME/mcp-config.json`（默认 `H/.copilot`）用 `mcpServers`；两者分别适配 | `H/.agents/skills` |

环境变量须指向绝对路径。配置与客户端环境变量不同、VS Code 自定义 profile/便携 IDE、DSH profile 层覆盖、企业策略或工作区配置都可能改变实际生效结果；不猜测未发现的配置目录，也不覆盖更高优先级策略。TRAE 文档提示部分版本的 `command` 不接受空格；若安装路径含空格，应在 TRAE 内确认其当前版本支持情况，不能仅凭粘贴成功判定可用。

## Skills 使用稳定链接

程序中的 Skill 原稿同步到字浮用户数据下的 `managed-skills/`，各客户端只链接到这个稳定位置，避免 Windows 更新到新版本目录后仍读取旧技能。macOS 使用目录符号链接，Windows 使用目录 junction；不会要求管理员或修改系统安全设置。旧副本或旧链接先移到相邻的 `.charfloat-skill-backups/`，不删除用户修改。源文件冲突、权限错误和链接读回失败都会报告。

已连接的托管源在字浮启动时刷新；更新不会自动给新客户端注入配置。未公开支持本地 Skill 的客户端保持应用内导入。文件系统链接成功与 AI 实际扫描成功分开验收；尤其 WorkBuddy 的官方文档主要描述导入和市场入口，不据此承诺所有版本都扫描链接目录。

## 验证与后续维护

“设置 → 连接检查”分别读取后台版本、办公加载项安装与连接、AI 配置、技能链接。后台关闭时检查仍可运行。它不会启动办公软件、执行用户 MCP 命令或自动修复。

新增或升级适配器时先核实厂商官方文档和当前版本源码，分别核对 Mac/Windows 用户目录、环境变量、格式、停用和授权规则。更新 `client-registry.ts`、相应适配器及本表，不能以路径存在代替配置正确。

```sh
node --import tsx --test tests/client-integration.test.ts tests/platform.test.ts tests/update-install.test.ts
npm run typecheck
```

临时目录回归覆盖损坏/重复字段配置、用户注释与停用保留、YAML 不执行表达式、软链接备份/失败回退、更新路径迁移。实机还需在所选 AI 客户端刷新连接，查看字浮工具列表、读取能力信息，再确认 Skill 被识别；不要用配置文件检查代替这一层验证。Windows 路径用例和 macOS 目录链接测试不等于 Windows 实机通过。

版本更新与程序发布流程见 [UPDATES.md](UPDATES.md)。服务器地址、机器路径、运行日志及验收报告只保留在 Git 忽略的本地资料中。
