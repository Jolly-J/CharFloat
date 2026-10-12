# 2.2.1 本地构建与版本管理核对

日期：2026-10-12。状态：两平台包构建和内容校验完成，macOS 实际包内 CLI 运行验证完成；未提交、建标签、上传或替换远端发布。2.2.0 产物保留且指纹未变。

## 本次产物

| 平台 | 位置 | 校验 |
|---|---|---|
| macOS arm64 | `release/2.2.1/mac/mac-arm64/字浮 CharFloat.app` | Info.plist 2.2.1、arm64；无开发者证书签名、未公证；实际 ASAR 与 unpacked 内容检查通过 |
| Windows x64 | `release/2.2.1/win/字浮-CharFloat-2.2.1-win-x64.zip` | PE machine=0x8664、便携 ZIP；解包路径中有 CLI 兜底；未在 Windows 实机运行 |

完整指纹与大小见 [results.json](results.json)，旧文件指纹见 [baseline.json](baseline.json)。mac 的 `app.asar` SHA-256 为 `358e58d6040da1b6de94c1f2cdb04d7b1e26ef802eddf252b64fa655342d039d`；Windows ZIP SHA-256 为 `6246812e046c697332608c10eda12e2d6c305f3d288d955a5deee92df4af35a8`。

最终版本目录只保留 mac `.app` 和 Windows ZIP，未复制临时 mac ZIP、blockmap、Windows 解包目录或构建诊断文件。源码并未提交，实际来源是固定提交加工作区差异，不将 HEAD 当作唯一来源。

## 版本与变更

用户确认本次升为 2.2.1并保留旧版。同步了 package、锁文件根版本、Bridge、WPS/Office 运行常量及清单和三处界面兜底；依赖版本未升级。Office 清单对应 2.2.1.0。版本断言读取 package.json，仍核对实际注册及 MCP 版本。

`CHANGELOG.md` 增加本次明确属于 2.2.1 的源码拆分和 CLI 导出修复，没有把较早的 Unreleased 记录全部挪到本版本。未新增完整版本管理系统；已有声明、当前缺口与建议见 [版本管理核对](versioning-audit.md)。

## 验证记录

| 检查 | 结果 | 证据 |
|---|---|---|
| 版本一致性 | 2.2.1 / Office XML 2.2.1.0 | [versions.json](versions.json) |
| 类型检查 | 退出码 0 | [typecheck.log](typecheck.log) |
| 完整回归 | 153/153，失败和跳过均为 0 | [tests.log](tests.log) |
| 开发及发布构建 | 退出码 0，独立保护副本 60 文件 | [build-release.log](build-release.log)、[release-assets.log](release-assets.log) |
| macOS 打包 | 退出码 0，显式关闭发布和开发者证书签名 | [mac-build.log](mac-build.log)、[mac-identity.json](mac-identity.json) |
| Windows 打包 | 退出码 0，串行平台构建；附带明文 CLI 兜底 | [win-build.log](win-build.log) |
| mac 最终包内容 | 59 个保护资源，实际解包运行入口齐全 | [mac-protection.json](mac-protection.json) |
| Windows 最终 ZIP | 59 个保护资源，包及解包入口齐全 | [win-protection.json](win-protection.json) |
| 实际 mac 包 CLI | 字节码入口无明文回退，142 项完整工具契约、提示词、能力资源、工具调用及未知工具错误通过 | [mac-runtime-probe.log](mac-runtime-probe.log) |
| 旧版保留 | 2.2.0 ZIP、mac ASAR 和主程序指纹均不变 | [baseline.json](baseline.json)及 [results.json](results.json) |

串行构建规避当前双平台钩子对同一个发布副本增删 `cli-full.cjs` 的潜在竞态；没有修改整个并行构建器。所有构建先输出到本任务临时目录，两平台结束并通过检查后再复制最终产物，不清理旧版本。

## 来源与复核边界

源码锚点、输入指纹、差异及新增源码副本见 [source-state.json](source-state.json)。固定基线加 `source.patch` 与 `source-untracked/` 用于恢复本次业务源码和构建所需新增代码；它不是包括全部验收产物的整工作区归档。

公开 v2.2.0 的附件信息只读核对，见 [existing-public-release.json](existing-public-release.json)，没有修改远端。

本机 WPS 真实业务证据来自升版前的开发候选，见 [宿主报告](../2.2.0-excel-modularization/host/README.md)。2.2.1 新包内的发布保护资源已完成内容及模拟检查，但未再次部署到真实宿主；Windows 只能据本次构建确认结构与架构，不能据交叉打包宣称实机运行通过。

本次只构建新包，没有替换 `/Applications` 中已安装应用或部署 2.2.1 插件。需要公开发布时再固定源码提交/标签并完成必要平台验收，不把本地包构建直接等同于已正式发布。
