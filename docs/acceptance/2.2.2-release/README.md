# 2.2.2 多客户端接入与发布验收

日期：2026-10-17。状态：两平台发布包构建、内容校验与包内 CLI 契约探针完成，待提交/打标签/上传；上传后在本文件追加发布记录。

## 元数据

| 字段 | 内容 |
|---|---|
| 产品版本 / 候选标识 | 2.2.2 / `2.2.2-release` |
| 交付范围及升版理由 | 15 客户端分类接入能力与多项接入正确性修复；**版本由项目负责人指定为 PATCH 2.2.2**（新增能力向后兼容，未按语义化版本自动推导 MINOR） |
| 平台 / 系统 / 宿主 / Electron | macOS 27.0.1 (26A434) arm64 构建主机；Electron 35.7.5；未在真实 WPS/Office 宿主验收（见限制） |
| 源码完整提交 SHA | 见本次提交（发布后追加记录） |
| 构建输入指纹 | `dist/release-app/protection-manifest.json` SHA-256 `9232c3db31f3f62c19c10a52d989bbadee2f6a2bc7b2a374842b47f99af5c950` |
| 工具契约快照 | [tools.json](tools.json)，142 个对外工具（与 2.2.1 相同契约，未增删工具） |
| 构建依赖口径 | 构建前按 `package-lock.json` 执行 `npm ci` 恢复依赖（`@modelcontextprotocol/sdk` 1.30.0 / `zod` 3.25.76）；锁文件未改动。此前工作区曾被 `pnpm install` 改成 pnpm 布局并解析出 `@modelcontextprotocol/sdk` 1.32.1（要求 zod ≥3.25），导致构建失败；按锁文件恢复后通过 |

## 本次产物

| 平台 | 位置 | 校验 |
|---|---|---|
| macOS arm64 | `release/2.2.2/mac/mac-arm64/字浮 CharFloat.app` | Info.plist `2.2.2` / `com.antigravity.charfloat`；`Mach-O arm64`；ad-hoc（linker-signed）签名，无开发者证书、未公证；`check:release` 通过（68 个受保护资源） |
| Windows x64 | `release/2.2.2/win/字浮-CharFloat-2.2.2-win-x64.zip` | PE32+ x86-64；exe 图标 7 尺寸（16/24/32/48/64/96/256，<256 为 32bpp DIB、256 为 PNG）与版本资源 `2.2.2.0` 实测命中；`check:release` 通过；未签名 |
| macOS 上传 ZIP | `CharFloat-2.2.2-mac-arm64.zip`（由上述 `.app` 用 `ditto -c -k --keepParent --norsrc --noextattr` 制作） | `check:release` 通过（68 个受保护资源），上传后与远端 `digest` 对账 |

关键哈希（完整值见 [results.json](results.json)）：

| 项 | 值 |
|---|---|
| mac `app.asar` SHA-256 | `18b90a7c3590288b7a488a79af0b8aa0272a571f51e7b972c43a7ab2e2c8e9ab` |
| mac 主程序 SHA-256 | `7baba2344adeeb821ff779b6d0f932365872ade22d6390cfb4ee777eb2f3231a`（与 2.2.1 相同，Electron 运行时未变） |
| mac 上传 ZIP SHA-256 | `34973a5f70638f2ec21509a9ca79e0dedcf0c51afbfce6d02c661f6a90d6e9ef`（115,341,769 字节） |
| Windows ZIP SHA-256 | `3a5fa2937143a6b56af51276cae2830a253b56f3a438cce3af424643449150ae`（125,924,137 字节） |

## 验证记录

| 检查 | 结果 | 证据 |
|---|---|---|
| 版本一致性 | 2.2.2 / Office 清单 2.2.2.0 | [versions.json](versions.json) |
| 导航 / 品牌 / 能力声明 | 全部通过 | 提交说明与命令输出 |
| 类型检查 | 退出码 0 | [typecheck.txt](typecheck.txt) |
| 完整回归 | 164/164，失败和跳过均为 0 | [tests.txt](tests.txt) |
| 完整构建（含加载项与字节码） | 退出码 0 | [build.txt](build.txt) |
| 发布副本 | `dist/release-app` 69 文件 | [release-assets.txt](release-assets.txt) |
| macOS 打包 | 退出码 0，`CSC_IDENTITY_AUTO_DISCOVERY=false`，显式 `--publish never` | [mac-build.txt](mac-build.txt) |
| Windows 打包 | 退出码 0，串行平台构建，`--publish never`，附带明文 CLI 兜底 | [win-build.txt](win-build.txt) |
| mac 最终包内容 | `check:release` 通过，68 个保护资源 | [mac-protection.json](mac-protection.json) |
| Windows 最终 ZIP | `check:release` 通过，68 个保护资源 | [win-protection.json](win-protection.json) |
| mac 上传 ZIP | `check:release` 通过 | [mac-zip-protection.json](mac-zip-protection.json) |
| 包内真实 CLI | 142 项工具契约、能力资源、提示词、工具调用与未知工具错误全部通过；无明文回退 | [mac-runtime-probe.json](mac-runtime-probe.json) |
| exe 图标与版本资源 | 7 条目图标格式合法；UTF-16LE 命中产品名 / `2.2.2.0` / 公司名 / 文件描述 | [win-metadata.json](win-metadata.json) |
| 旧版保留 | 2.2.1 / 2.2.0 产物未改动，2.2.1 ZIP 哈希与上一版记录一致 | [baseline.json](baseline.json) |

命令：`npm ci`、`npm run build:release`、`npm run build:win-icon`、`npx electron-builder --mac --arm64`、`npx electron-builder --win --x64`（均 `--publish never`、隔离 `directories.output`）、`node scripts/check-release-protection.mjs <实际路径>`、`node scripts/probe-release-mcp.mjs <electron> <cli.cjs> <asar 解包根> <tools.json>`。

## 验收限制（未验证项）

- **未做真实 WPS/Office 宿主验收**：本次未改宿主实现、加载项业务逻辑与工具契约，加载项仅同步版本号并重建；因此未在真实宿主验证加载项升级后的读写。不能据此判定宿主侧验收通过。
- **未做 Windows 实机运行验收**：本机无 Windows 桌面；Windows 包仅完成静态内容、保护、图标与版本资源核验，未运行 exe。
- **macOS 未公证、无开发者证书**：签名状态为 ad-hoc（linker-signed），Gatekeeper 首次运行需用户手动放行；不得描述为开发者签名。
- **Windows 未签名**：SmartScreen 会提示未知发布者。
- 交叉构建 Windows 时，rcedit 经 electron-builder 自带 wine 执行，需要工作区外写权限（wine prefix）；按 [2.1.0 P5.2 记录](../2.1.0-p0p1/p5/p5.2-windows-cross-build.md) 的已知方式放行该构建命令，未降级 `signAndEditExecutable`，exe 图标与版本信息完整。

## 回退

- 上一正式版本 2.2.1 包与指纹保留未动；需要回退时下载 GitHub Releases `v2.2.1` 的两平台附件即可。
- 本次升级向后兼容：MCP 工具契约未变化，客户端配置写入仍只改本产品条目。旧版 2.2.x 留下的 `~/.codex/mcp.json`、`~/.qwen/mcp.json` 不会自动清理，可按需手动删除。

## 正式发布记录

待上传后追加：标签、完整提交 SHA、Release URL、附件名/大小/`digest`、对账结果与时间。
