# Bridge 2 结构

`src/main` 只负责桌面窗口、托盘、安装与本机服务管理。服务不由窗口持有；UI IPC 通过本机认证 HTTP 请求访问后台。

`src/bridge/cli.ts` 提供管理命令和 stdio 代理。`service-client.ts` 检查服务身份、协议及凭据，在没有服务时分离启动 `--serve`；端口占用时不终止未知进程。后台由绑定端口保证单实例。

`ws-server.ts` 承载 HTTP、Streamable HTTP、SSE 和认证加载项 WebSocket。办公文档、工具、提示词端点均需凭据；health 只返回服务身份、版本与 PID。

`catalog.ts` 是协议层共享工具清单与参数校验入口，提供动态能力与诊断。`excel_*` 统一入口显式指定宿主；旧 `wps_*` 兼容入口保留。

`gateway.ts` 保留既有工具参数映射和单元格审计。`office/adapter.ts` 把宿主调用送到 WPS 加载项或 Windows Excel 适配器。办公调用串行执行，避免剪贴板/目标争用；原生子进程通过异步 stdin/stdout 交互，不拼接 shell 参数。

`resources/office` 为 Windows PowerShell COM 实现。结构化参数通过 JSON stdin 输入，目标精确匹配；找不到、名称不唯一或宿主返回错误时停止。JXA 通道显式对应 macOS Office，不伪装成结构化能力对等。

审计包含宿主标识，按修改后快照进行回滚冲突检测。任意脚本等未覆盖操作的限制在能力资源和 Skill 中说明。

技能、运行时能力描述与验证报告共同维护：实现清单不能自动升级为实机验证状态。后台日志不记录用户工具参数和认证 token。
