# 加载项侧栏与后台连接

## 统一样式与品牌图标

两种侧栏使用 `office-addon/public/taskpane.css` 作为唯一视觉样式来源，构建时同步到 WPS 的 `panel.css`，颜色、品牌区、状态标签、工作区卡片、字号与底部按钮一致；各自宿主的功能差异保留。两种侧栏均显示 App PNG，Office 加载项列表与工具栏分别由清单图标定义控制。

Office 图像响应不得带禁止缓存指令；原图标 URL 的 `/assets/` 映射漏掉实际目录，导致请求落入失败响应；原托管逻辑对 PNG 也返回 no-store，会导致默认图标替代。图片保留 ETag / Last-Modified 供缓存使用，HTML/JS/CSS 继续不缓存；清单与侧栏图标 URL 已更新以避开旧图标缓存。依据：[微软图像规范](https://learn.microsoft.com/en-us/javascript/api/manifest/image?view=word-js-preview)。加载项 UUID、运行时、权限与工具路由保持不变。

## WPS

WPS 表格、文字、演示在已有的“开始”功能区中保留一个使用 App 图标的大号“字浮侧栏”按钮（上图下文），打开右侧任务窗格；不再增加独立的整页功能区。

`index.html` 是宿主加载入口，持有唯一的 WebSocket/RPC 执行器。`panel.html` 通过 `CHARFLOAT_PANEL_VIEW` 标识 UI 上下文，复用操作函数，但不创建 WebSocket、不注册 RPC、不挂载宿主事件。面板仅提供工作区、连接状态、运行日志和收起/重连操作；从认证 HTTP 状态接口显示对应组件的文档与连接状态，关闭或收起面板不会停止宿主连接。

任务窗格按组件保留单实例。支持共享存储时可跨入口重载恢复；缺少该接口时仍复用当前入口中的窗格引用。收起使用创建后取得的实际 ID，不猜测固定 ID；已有任务窗格在构建变化时重新导航到新资源。部署器为两种入口的 JS 和侧栏 CSS 添加构建指纹，指纹涵盖侧栏页面、样式、Ribbon 入口及 App 图标，保证同版本样式修改也触发更新。

侧栏仅保留工作区、状态和日志，底部为收起与重新连接。侧栏使用原始 App PNG；功能区使用构建时从原图直接预缩放的 32px 图标及 64px `@2x` 资源，以独立文件名避免旧图标缓存。资源由 `scripts/build-wps-ribbon-icons.mjs` 生成并纳入部署指纹。

用户实机截图已证实该 WPS 能显示 App PNG，因此此前 [2026-06-15 官方回复](https://bbs.wps.cn/topic/90097) 不应作为当前版本一概不支持的结论。源图和可读取部署副本均为 1024px，原图低分辨率已排除；具体内部缩放链路及是否使用 @2x 资源尚未确认。工具栏专用资源是针对缩放问题的改进候选，实际清晰度仍须宿主截图验证，不以构建/本地图像清晰替代真实验收。


官方接口：[CreateTaskPane](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/Application/member/CreateTaskpane)、[任务窗格](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/TaskPane/task-pane-overview)。具体宿主版本的加载和窗口行为仍须实机验证。

## Microsoft Excel

当前 Office 加载项面向 Microsoft Excel。旧实现仅在任务窗格 WebView 启动后连接 CLI，面板未加载时没有可运行的连接代码。

清单在 `VersionOverrides` 内配置 `SharedRuntime 1.1` 和 `Runtime lifetime="long"`，Ribbon、FunctionFile、SourceLocation 使用同一 `Taskpane.Url`，并移除多任务窗格 ID。字浮后台须已启动，以便 Excel 从本机 HTTPS 地址加载运行时；已加载的运行时在后台服务重启后仍会重试连接。首次激活后调用 `Office.addin.setStartupBehavior(Office.StartupBehavior.load)`，让当前工作簿下次打开时加载运行时；显示/隐藏侧栏不重新创建连接或事件监听器。

更新清单后需重新启用加载项，让 Excel 读取新的运行时配置；旧缓存仍生效时，退出 Excel 后重新启用。

**范围是每个工作簿**：首次必须手动启动一次加载项；已启用的工作簿随后可收起侧栏，重新打开时自动连接。首次打开的其他工作簿仍须单独启用。该机制不意味着“安装一次后所有新工作簿零点击加载”。不修改工作簿业务数据，也不调用保存命令。

多个 Office 工作簿各自保留后台连接，调用按 `workbookName` 或文档路径定位；目标有歧义时在发出请求前拒绝，不切换到原生通道猜测。`bridge_get_capabilities` / 状态接口的 `officeWorkbooks` 可查看连接名单。同名工作簿可用完整文档路径区分。

不支持共享运行时的 Excel 保留基本任务窗格入口，提示需保持侧栏开启；启用自动启动失败时保留当前连接并明确提示下次需手动启动。HTTP 与客户端两条清单部署路径均按内容比较，不能因版本号相同而跳过运行时配置更新。

官方依据：[共享运行时配置](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)、[文档打开时运行](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/run-code-on-document-open)、[支持版本](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)。`OnDocumentOpened` 的应用级激活需要不同的部署条件，且当前不支持 Office on Mac，不用于本机侧载方案。

## 验证边界

`tests/addon-lifecycle.test.ts` 覆盖：WPS 不重复创建侧栏、面板不抢占 RPC 连接、收起后仍响应请求、组件显示隔离、旧窗格更新；Office 首次配置、已配置时不重复写入、隐藏后响应请求、重复初始化防护、版本不支持及设置失败，以及清单关联。

代码与模拟检查通过后，还需分别在 WPS 和 Excel 真机验证：打开侧栏、收起后读取测试文档、重新打开已启用工作簿、不打开侧栏时读取测试文档。浏览器页面检查不能替代这组宿主验证，也不能用模拟结果声明真实宿主已通过。
