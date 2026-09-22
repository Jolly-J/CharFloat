# 发布包代码保护

## 范围和边界

保护只发生在构建/打包阶段，不改业务源码、MCP 契约、技能原稿或宿主操作行为。
开发者继续维护 `src/`、`wps-addon/src/`、`office-addon/src/`、`skills/`、`prompts/`。

本地运行的软件无法保证不被逆向或原样复制。当前措施降低直接取得源码和阅读实现的便利性，
不是授权系统、加密保险箱或“不可复刻”承诺。ASAR 是归档；SHA-256 清单是构建证据，不是签名。

| 内容 | 发布处理 | 已知边界 |
|---|---|---|
| CLI | Electron V8 字节码 `cli.jsc` + 小加载器；不带明文回退 | 常量可提取，存在逆向路径；不能阻止复用字节码 |
| WPS/Office 加载项 | 仅发布副本做固定种子的标识符/字符串数组混淆 | 仍是可执行 JS，可运行分析；真实宿主兼容性需实测 |
| MCP 工具描述、技能、参考、提示词 | 完整保留 | Agent 必需的调用知识不作为秘密 |
| Windows COM PowerShell | 原样保留 | 运行必需且仍为明文；没有声称原生核心保护已完成 |
| 主进程、界面、预加载 | 保留现有压缩构建产物 | 仍可分析，不声称字节码覆盖全产品 |
| 源码、map、旧入口、备份、开发规范 | 发布白名单排除 | 新增运行资源须更新白名单 |

## 构建流程

```text
npm run build                  开发构建（原有流程）
        ↓
npm run build:release-assets   只生成 dist/release-app/ 发布副本
        ↓
electron-builder              beforePack → afterPack → artifactBuildCompleted
```

`npm run build:release` 连续执行前两步。`dist`、`dist:mac`、`dist:win`、`dist:all`、`dist:fast`
已经接入发布构建，后续正常发包沿用这些命令即可。直接调用 electron-builder 前须先生成最新发布副本。
开发生成物不会被发布混淆器覆盖，已有 `release/` 包不会被 build:release-assets 修改。

- [release-files.json](../scripts/release-files.json)：运行资源显式白名单；界面带哈希资源单独限定文件类型。
- [build-release-assets.mjs](../scripts/build-release-assets.mjs)：从白名单复制到专用目录，混淆两个加载项，输出指纹清单。
- [release-protection.mjs](../scripts/lib/release-protection.mjs)：固定混淆策略、文件检查、过期副本检查。
- [before-pack.cjs](../scripts/before-pack.cjs)：缺文件、哈希不符或输入已变即阻止打包。
- [after-pack.cjs](../scripts/after-pack.cjs)：检查实际 ASAR 与 unpacked，缺失真实 CLI 入口/字节码/Bytenode 即失败。
- [after-artifact.cjs](../scripts/after-artifact.cjs)：重新读取最终 ZIP 校验，旁置 `.zip.protection.json`，绑定实际 ZIP SHA-256。

`dist/release-app/protection-manifest.json` 记录发布资源哈希与构建输入哈希。它不含源码和凭据内容。
构建依赖只用于本机处理，不调用混淆云服务，不上传代码。

## 混淆兼容性约束

不重命名属性或全局接口，不改 MCP 工具名与参数；不使用反调试、自毁、死代码注入、
控制流平坦化或额外 eval 包装。字符串编码仅用于增加静态阅读成本，不是保密密钥。

WPS 的发布指纹同时包含输入与保护配置；头部注释和运行时注册值保持一致，
避免开发副本与保护副本被误判为同一构建。原始构建文件不变。

## 检查命令

```sh
npm run build:release
node --import tsx --test tests/release-protection.test.ts
npm run check:release -- '<实际 ZIP / app.asar / .app 路径>'
```

`check:release` 不执行包内代码；没有包、显式路径不存在、没有保护清单都必须失败。
旧 `check:protection` 入口也使用这套包级检查，不再移动本地明文文件或连接默认用户服务。

MCP 运行验收单独执行：

```sh
node scripts/probe-release-mcp.mjs '<目标 Electron 可执行文件>' '<发布 CLI 路径>' '<发布资源根目录>' '<工具快照 JSON>'
```

探针只接受无明文回退的入口，使用临时运行目录和随机本机端口，测试后停止自己的服务。
验证完整 tools/list（名称、顺序、描述、schema、只读标记）、能力资源、提示词、工具调用与未知工具错误。
可在目标操作系统上调用该系统的包内可执行文件；在 macOS 使用本机 Electron 读取 Windows 包不等于 Windows 真机验证。
真实 WPS/Office 加载和业务操作仍须用独立测试文档验收，不能用构建或模拟测试代替。

## 已证实的打包陷阱

- 从发布副本目录打包 → 沿用根路径 asarUnpack → 匹配使用源路径，CLI 未解包 → 用 `**/dist/bridge/cli.cjs` 等匹配并检查实际 unpacked → afterPack 门禁已实际拦截此错误。
- 只检查本地 dist/ → 发布目录缺文件仍可能漏检 → 校验最终 ASAR 和 ZIP，而非本地代码片段。
- 只检查工具数量 → 描述或参数丢失也可能通过 → 比对完整工具契约，并实跑提示词和资源读取。
- Electron 升级 → 旧字节码与 V8 不兼容 → 用目标 Electron 重编译，目标系统实跑发布 CLI。

本轮范围、基线和验证结果见 [构建保护验收](acceptance/2.1.0-protection/README.md)。
