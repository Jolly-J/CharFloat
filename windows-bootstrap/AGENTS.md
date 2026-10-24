# Windows 当前用户配置引导

- `main.mjs`：仅首次安装的独立 Electron 引导窗口、IPC、桌面快捷方式和显式启动；拒绝旧更新参数；不进入主应用或办公宿主。
- `update-zip.mjs`：已签名完整 ZIP 的有界清单生成、PE 平台/架构、ASAR 应用版本及未知文件验证；由客户端独立助手调用，禁止运行下载 EXE。
- `update.mjs`：供独立助手调用的更新准备；验证签名交接、确认旧目录归属、等待旧 GUI 退出，core 提交前消费 GUI 就绪后授予的一次性授权，并在目录/来源/快捷方式提交点重新核对未撤销；失败旧票据不能因用户后来退出而生效。`core.mjs` 仅允许更高版本，来源记录独立于目标目录首次创建，复用手动失败留下的目录仍捕获实际旧目录；不重跑客户端首次配置。文件锁或快捷方式失败保留旧版与用户数据供新授权重试。通用流程见[更新指南](../UPDATES.md)。
- `core.mjs`：ZIP/逐文件 SHA-256、路径约束、有界解压、独立版本目录与锁；只接受调用方提供的用户目录，测试必须用临时目录。
- `physical-fs.mjs`：Electron 中解压/核验使用 `original-fs` 操作物理 ASAR；普通 Node 保留内置 fs。不能全局设置 `process.noAsar`，引导自身仍从 ASAR 加载界面。
- `zip-stream.mjs`：把锁定 ZIP 库的旧读取流转换为原生流，保留背压、错误和取消传播；Mac 更新、Windows 预检与解压共用。
- `preload.cjs`：沙箱 renderer 的最小 IPC 接口。
- `progress.mjs`：主进程跨文件合并进度，最多每 200ms 刷新同阶段的最新值；阶段切换即时，错误或完成前销毁待发送状态，避免小文件突发或终态后回退。
- `diagnostics.mjs`：仅诊断入口的 runId/用户目录启用，白名单与脱敏 JSONL、2MiB 上限、退出终态；主进程/渲染回执/异常与心跳。普通启动默认关闭，不把诊断变量传给已安装应用。
- `Start-Local-Debug.cmd` / `.ps1`、`local-debug-README.txt`：离线调试入口，先核包内 `debug-package.json` 的 SHA/大小，再以普通用户启动本包并实时合并为一份报告。不改执行策略或全局环境；原生 NSIS 内部和已安装应用不在完整采集范围。
- `diagnose-startup.cmd` / `.ps1`、`diagnose-startup-README.txt`：不启动 EXE 的原生临时目录诊断工具，普通用户运行，策略阻止即停止；预置哈希仅识别早期故障样本，不是当前发布包的验收标准。报告仅本地保存，不纳入 Git。
- `Read-DiagnosticSnapshot.ps1`：共享读写快照，按 runId/sequence 去重并保留未完成行供下次完整重读；暂态读取失败保留已有事件，恢复后清除当前告警，累计恢复次数另报。
- `shortcut.mjs`、`launch.mjs`、`stages.mjs`：安全快捷方式备份/读回、再次校验后启动、最低展示时间。
- `index.html`、`style.css`、`renderer.js`：三个至少五秒的品牌阶段；分别显示包校验、实际写入与安装文件复核进度和当前文件，不按时间编造。`bytes` 是当前阶段实际处理字节，`verifiedBytes` 只计哈希通过的整文件，写入 100% 不代表安装成功。
- 构建入口：`scripts/build-windows-bootstrap.mjs`，生成独立 `dist/windows-bootstrap`，不修改业务发布副本。输入必须为已校验的 Windows ZIP 和固定源码 SHA；品牌图选自 `website/public/character`，三张输入固定保存于 `windows-bootstrap/assets`，避免依赖被 Git 忽略的官网目录。

默认仅写 `%LOCALAPPDATA%/Programs/CharFloat` 与 `app.getPath('desktop')` 当前用户桌面。保留旧版本、用户数据及陌生文件；不能请求管理员权限、结束其他进程、绕过企业策略或静默覆盖未知快捷方式。失败不能显示完成。

验证：`node --import tsx --test tests/windows-bootstrap.test.ts`；实际 Electron UI、Windows Shell 快捷方式与最终 EXE 需分别报告，模拟测试不能替代实机。

## 避坑

真实 ZIP 含 STORE（未压缩）条目时预检/解压停住 → 只用 DEFLATE 小样本无法复现 → yauzl 2 的 fd-slicer 在 EOF 前标记 destroyed，却不发 close，现代 async iterator/pipeline 会等待不来的完成信号 → 通过 data/end/error 转接原生 PassThrough，保留背压与失败传播 → STORE 空文件/普通文件/框架链接回归及真实双平台 ZIP 预检；完整 Electron 解压单独验证。

普通 Node 解压通过，但 Electron 卡在载荷 `app.asar` → 继续沿用 Electron 包装的 fs 会把物理档案当成虚拟目录 → 默认 fs 的 ASAR 拦截影响写入和大小/类型核验 → 安装引擎使用 Electron `original-fs`，保留引导自身的 ASAR 支持 → 必须用完整 Electron 主进程解压真实载荷；`ELECTRON_RUN_AS_NODE=1` 的测试不能覆盖此路径。

外壳不显示独立 splash → 只删 BMP 或编译 define 会漏掉临时目录初始化 → 锁定模板在无 splash 时静默启动，原初始化只放在 `.onGUIInit` → `scripts/lib/windows-portable.mjs` 核对模板指纹后，仅在本次编译输入的 `.onInit` 显式初始化，先于静默模式与解压；不修改依赖文件或系统 TEMP → 编译输入/参数回归、最终 EXE 无 BMP/BgImage；Windows 启动仍须实测。

Windows 写日志期间读取快照 → `ReadAllLines` 的默认共享会与已有写句柄竞争，且陈旧 warning 会误导 → Node/libuv 写句柄本身允许 Read/Write/Delete 共享 → reader 显式共享相同权限，只解析完整行；失败保留事件，恢复清除当前 warning 并另报历史次数 → `tests/windows-bootstrap-sharing.ps1` 与 Windows 条件回归；Mac 跳过不能当实机通过。

解压单个大文件或复核安装 → 只在整文件完成时刷新累计值，界面会长时间停在上一项 → 流仍可正常读写而用户看不到进展 → 每 200ms 读取输出流 `bytesWritten`，并明确区分写入、文件核验和全量复核；哈希未通过不能计入 `verifiedBytes` → 见真实 ZIP 慢写入与错误哈希回归 `tests/windows-bootstrap.test.ts`。

`app-builder-lib` 25.1.8 的 `unpackDirName` 注释与实现相反 → 按注释设 false 会得到同构建共享临时目录 → 实际 `NsisTarget` 在 `true` 时才省略固定目录 define → 打包入口断言没有 `UNPACK_DIR_NAME`，在 `.onInit` 显式初始化每次启动独立插件目录。升级 builder 后必须复核，模板指纹不符停止构建。

## 同步维护

构建入口或产物变化同步根导航、scripts/AGENTS.md 与本地 `docs/release-policy.md`（不随 GitHub clone 提供）；不把旧源码标签误当新引导源码，分别保存应用载荷 SHA 与引导提交。
