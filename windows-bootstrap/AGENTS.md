# Windows 当前用户配置引导

- `main.mjs`：独立 Electron 引导窗口、IPC、桌面快捷方式和显式启动；不进入主应用或办公宿主。
- `core.mjs`：ZIP/逐文件 SHA-256、路径约束、有界解压、独立版本目录与锁；只接受调用方提供的用户目录，测试必须用临时目录。
- `preload.cjs`：沙箱 renderer 的最小 IPC 接口。
- `shortcut.mjs`、`launch.mjs`、`stages.mjs`：安全快捷方式备份/读回、再次校验后启动、最低展示时间。
- `index.html`、`style.css`、`renderer.js`：三个至少五秒的品牌阶段；工作进度来自已核验的实际写入字节，不按时间编造。
- 构建入口：`scripts/build-windows-bootstrap.mjs`，生成独立 `dist/windows-bootstrap`，不修改业务发布副本。输入必须为已校验的 Windows ZIP 和固定源码 SHA；品牌图选自 `website/public/character`，三张输入固定保存于 `windows-bootstrap/assets`，避免依赖被 Git 忽略的官网目录。

默认仅写 `%LOCALAPPDATA%/Programs/CharFloat` 与 `app.getPath('desktop')` 当前用户桌面。保留旧版本、用户数据及陌生文件；不能请求管理员权限、结束其他进程、绕过企业策略或静默覆盖未知快捷方式。失败不能显示完成。

验证：`node --import tsx --test tests/windows-bootstrap.test.ts`；实际 Electron UI、Windows Shell 快捷方式与最终 EXE 需分别报告，模拟测试不能替代实机。

## 避坑

`app-builder-lib` 25.1.8 的 `unpackDirName` 注释与实现相反 → 按注释设 false 会得到同构建共享临时目录 → 实际 `NsisTarget` 在 `true` 时才省略固定目录 define → 由打包入口断言真实编译参数没有 `UNPACK_DIR_NAME`，并用 splash 明确初始化每次启动独立的插件目录。升级 builder 后必须复核，不能盲用此兼容值。

## 同步维护

构建入口或产物变化同步根导航、scripts/AGENTS.md 与 docs/release-policy.md；不把旧源码标签误当新引导源码，分别保存应用载荷 SHA 与引导提交。
