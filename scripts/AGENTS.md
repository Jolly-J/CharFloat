# 辅助脚本与副作用

## 文件地图

- [dev.ts](dev.ts)：启动 Vite 和 Electron；启动前依次构建主进程与**加载项部署入口**（`build:main` + `build:addons`）。
- [inspect.ts](inspect.ts)：启动服务并探测宿主。
- [verify-all-mcp-tools.ts](verify-all-mcp-tools.ts)：访问真实 Microsoft 宿主的工具验证。
- [verify_all_tiers.ts](verify_all_tiers.ts)：真实 Bridge 集成操作。
- [build_beautified_sheet.ts](build_beautified_sheet.ts)：真实表格写入与排版。
- [create-icons.mjs](create-icons.mjs)：生成图标文件。
- [check-agents.mjs](check-agents.mjs)：只读检查协作导航。
- [snapshot-tools.ts](snapshot-tools.ts)：只读导出工具契约与能力快照（`npm run snapshot:tools`）。
- [check-capability-claims.ts](check-capability-claims.ts)：只读核对 skill、官网与**加载项页面**中的能力数量表述与元数据源（`npm run check:claims`）。
- [check-param-forwarding.ts](check-param-forwarding.ts)：只读核对"**schema 声明了参数、网关处理器却没读取**"的静默丢参（`npm run check:params`）；带自检，证明能抓到未转发参数且读过的不误报。
- [build-win-icon.mjs](build-win-icon.mjs)：由 `build/icon.png` 生成 Windows 用的**多尺寸** `build/icon.ico`（`npm run build:win-icon`）；纯 Node 实现（自带 PNG 解码/缩放/编码与 DIB 封装），写完自检格式，不通过不落盘。
- [dist-parallel.mjs](dist-parallel.mjs)：**并行**跑两个平台的打包（`npm run dist:all`）。`electron-builder --mac --win` 是串行的，而打包以 I/O 为主、单进程吃不满多核；两平台写入不同目录，可安全并行。`--no-sign` 跳过 macOS 代码签名（`npm run dist:fast`）——实测签名是 mac 侧的主要耗时（mac 204s vs win 43s，且 electron-builder 自身 CPU 只用了 3 秒：`codesign` 在等 Apple 时间戳服务器，属网络等待，所以 CPU 闲置）。
- [build-cli-bytecode.mjs](build-cli-bytecode.mjs)：把 `dist/bridge/cli.cjs` 编译成 **V8 字节码** `cli.jsc`，并把 `cli.cjs` 就地替换为加载器（`npm run build:bytecode`，已串进 `npm run build`）。**必须用 Electron 的 Node 编译**（V8 版本与运行时锁死）。详见 [code-protection.md](../docs/code-protection.md)。
- [check-cli-protection.mjs](check-cli-protection.mjs)：旧命令兼容入口，委托包级检查，不移动开发文件或连接默认用户服务。
- [build-release-assets.mjs](build-release-assets.mjs)：仅生成 `dist/release-app/` 发布副本；输入白名单见 [release-files.json](release-files.json)，保护策略与过期校验见 [lib/release-protection.mjs](lib/release-protection.mjs)。
- [check-release-protection.mjs](check-release-protection.mjs)：校验最终 ZIP/ASAR，底层见 [lib/check-release.mjs](lib/check-release.mjs)。
- [before-pack.cjs](before-pack.cjs)、[after-pack.cjs](after-pack.cjs)、[after-artifact.cjs](after-artifact.cjs)：打包前防过期、包内/解包路径检查、最终 ZIP 哈希证据。
- [probe-release-mcp.mjs](probe-release-mcp.mjs)：隔离运行发布 CLI，比较完整工具契约、提示词、能力资源；只使用临时 HOME 和随机端口，关闭自己的子进程，不控制真实宿主。
- [probe-mcp-tools.mjs](probe-mcp-tools.mjs)：以 stdio 启动 `cli.cjs` 做一次 MCP 握手、**只输出工具数量**；供上面的校验脚本复用。
- [build-wps-addon.mjs](build-wps-addon.mjs)：由 `wps-addon/src/**` 生成部署入口 `wps-addon/addon-core.js`（`npm run build:wps-addon`，`--check` 只校验）。
- [build-office-addon.mjs](build-office-addon.mjs)：由 `office-addon/src/**` 生成部署入口 `office-addon/public/taskpane.js`（`npm run build:office-addon`，`--check` 只校验）。

## 定位与联动

先看脚本入口、目标、连接地址和写操作，再决定能否运行。命令入口见 [package.json](../package.json)。

## 修改边界

文件名带 inspect/verify 不代表只读；实际宿主脚本不能作为普通单元测试执行。不要沿用示例里写死的端口、认证和目标。新增脚本说明输入、输出与副作用，避免打印凭据。

## 验证

导航检查用 `npm run check:agents`；能力声明一致性用 `npm run check:claims`（两者都只读）。加载项构建用 `npm run build:addons`（**会覆写部署入口**，属写操作但只写生成物）。其余按实际副作用选择隔离环境和验证方式。

## 避坑

Windows 包 exe 图标显示空白 → 只看"`RT_GROUP_ICON` 条目数够不够"会得出错误结论：**条目数量正确、图标画面正确，仍然可能不显示** → 已证实原因：Windows 只对 **256×256** 支持 PNG 内嵌条目，小于 256 的条目必须是 **BMP(DIB)**；全部用 PNG 时资源管理器的小/中/大图标视图拿不到可解码的条目 → 正确做法：<256 写 32bpp DIB（`BITMAPINFOHEADER` 40 字节 + 自下而上 BGRA + 1bpp AND 掩码），256 写 PNG，与原版 `electron.exe` 的资源结构一致 → 由 [build-win-icon.mjs](build-win-icon.mjs) 生成并在写完自检，`npm run dist` 已串联该步 → 校验方法：用 `resedit` 读 exe 的 `RT_ICON`，逐条判断是 DIB 还是 PNG（对照 [win-icon.test.ts](../tests/win-icon.test.ts)）。另注意 **PE 版本资源是 UTF-16LE**，`strings -a` 只找 ASCII 会漏判。

直接手改 `wps-addon/addon-core.js` 或 `office-addon/public/taskpane.js` → 下次构建即被覆盖，且源码与部署产物不一致 → 两者都是生成物（顶部有"请勿手改"） → 改 `*/src/**` 后跑对应构建 → `--check` 模式可检出漂移。

能力文案检查只看 `skills/` 与官网 → **加载项页面这类手写文件里的数量文案会漏网** → 已证实：`office-addon/public/taskpane.html` 长期显示"42 项能力已就绪"，与四个口径（路由 30 / 对外工具 91 / WPS 26 / Microsoft 27）都对不上，并随发布包出厂；生成器只产 `taskpane.js`，检查器既不扫该目录、模式也不含"`N 项…能力`" → 扫描范围加入 `office-addon/public` 与 `wps-addon`，模式加入 `(\d+)\s*项[^，。；、\n]{0,16}?能力`；文案里的宿主必须写明（写 `Microsoft Excel`，不能只写"Excel"），否则按"未指明宿主"报错 → 反向验证：把 27 改回 42 即 `exit 1`。

直接执行 verify 脚本 → 可能修改真实文件 → 先检查请求列表和目标 → 使用明确测试文档并读回，不能靠脚本名称判断安全性。

用本机 node 编译 V8 字节码 → 在 Electron 里**加载即失败** → 字节码与 V8 版本锁死（实测本机 V8 14.1 / Electron V8 13.4）→ 必须用**目标运行时的二进制**编译（`ELECTRON_RUN_AS_NODE=1 node_modules/electron/dist/...`），并把编译时版本写进加载器；版本不符时报可操作错误而不是 undefined → [build-cli-bytecode.mjs](build-cli-bytecode.mjs)。

编译字节码前不删旧文件 → bytenode 按**输入文件名**产出（`cli-full.cjs` → `cli-full.jsc`），产物缺失时会被上一轮的陈旧文件蒙混过关，**打出来的包跑的是旧代码** → 编译前先 unlink 目标，编译后校验再改名。

以为"加了 minify 就没有源码了" → 注释和函数名没了，但**逻辑仍可读**；而且明文产物（`cli-full.cjs`）若没进 `build.files` 的排除清单，等于白做 → 排除清单与保护手段必须同时到位 → [check-release-protection.mjs](check-release-protection.mjs) 的包内检查。

给 esbuild 的 minify 加 `mangleProps` → 连 `args?.workbookName` 这类**属性名**一起改名，双宿主靠属性名对接，**功能全废** → esbuild 默认不重命名属性名，**不要开这个开关**。

只验证"构建成功"就发布代码保护改动 → 字节码不会因构建失败而失败：**它只会在运行时加载不了** → 必须验证「把明文包临时移走仍能跑」，才证明真的走字节码而非悄悄回退 → [probe-release-mcp.mjs](probe-release-mcp.mjs) 只接受无明文回退的发布入口。

只验证 `app.asar` 里的内容就认为打包没问题 → 实际跑不起来、且**全平台都挂**（不只是出问题的那个平台）→ `src/main/index.ts` 会把入口**重定向到 `app.asar.unpacked`**，而 `asarUnpack` 是按**文件名逐个列**的：漏列 `cli.jsc` 或 `bytenode` 就缺文件/缺模块，进程起来即死，主进程等 6 秒后报"后台启动失败" → 新增可执行文件后**必须同步 `asarUnpack`**，并用**安装器实际使用的那条路径**跑一次 → [lib/check-release.mjs](lib/check-release.mjs) 的 unpacked 检查。

**只比 V8 版本号就认为字节码可用** → Windows 上仍报 `Invalid or incompatible cached data (cachedDataRejected)`，界面表现是「后台启动失败」且 `service.log` **为空**（V8 在原生层中止，不抛 JS 异常）→ V8 字节码的缓存头除版本号外还含**平台/编译配置的 flags hash**，macOS 与 Windows 都是 V8 13.4 也**不通用**；加载器必须**真正 try 一次**并在失败时回退，不能只做版本比较 → [build-cli-bytecode.mjs](build-cli-bytecode.mjs) 的加载器模板；兜底策略见 `before-pack.cjs`。

Windows 包没有明文兜底 → 字节码被拒后无处可退，直接起不来 → 在 macOS 上**无法**生成 Windows 的字节码（跑不了 PE），只能让 Windows 包附带 `dist/bridge/cli-full.cjs`，由加载器回退；**代价是 Windows 侧保护降为 minify** → `before-pack.cjs` 按平台增删该文件，校验侧用 `allowPlainCli` 放行（**只放行这一个文件**），macOS 侧仍移除它 → [release-protection.mjs](lib/release-protection.mjs) 的 `checkFiles`。

`cli-full.cjs` 只放进 asar、没解包 → 加载器在 `__dirname`（**解包目录**）里找不到它，回退静默失效 → `asarUnpack` 是按**文件名逐个列**的，必需文件不在清单里就只进 asar → 新增运行文件必须同步 `asarUnpack`，并用安装路径实跑一次 → [check-release.mjs](lib/check-release.mjs) 的"Windows 明文兜底未解包"断点。

重复执行 `build:bytecode` 而不先重建明文包 → 上一次的加载器被当成源码编译，产出 **1 KB 的假字节码**，**构建照样"成功"**，但打出的包必然起不来 → 脚本开头校验输入 `cli.cjs` 是否为明文包（含 `require('bytenode')` 即判为加载器并拒绝） → [build-cli-bytecode.mjs](build-cli-bytecode.mjs)。

发布副本不先清空就重建 → `beforePack` 会按平台往副本里增删文件（Windows 加明文兜底、macOS 移除），上一轮 Windows 留下的兜底文件让本次 `verifyStage()` 直接报「禁止发布」，看起来像源码问题 → 重建前 `rmSync(STAGE)` → [build-release-assets.mjs](build-release-assets.mjs)。

## 同步维护

文件入口、职责、调用关系或验证方式变化时同步更新本页；新增已证实的重复问题时补充原因、处理方式及证据。其余遵循根目录协作规范。
