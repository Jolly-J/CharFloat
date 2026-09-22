# 2.1.0 P0.5 / P0.6 环境与发布矩阵盘点

- 盘点对象：改造计划 [refactoring-plan.md](../../refactoring-plan.md) 第 4 节 P0.5（真实验收环境与测试资产）、P0.6（发布架构矩阵与签名要求）。
- 候选源码指纹：`git rev-parse HEAD` = `cac6d38e87791e53e61f3ecfc7286e0f39606f68`（工作区脏，未提交修改见 `git status --short`，本文件不重复记录）。
- 本盘点为**只读**：未安装任何软件，未运行 `npm run dist` / `npm run build` / `npm test` / `electron-builder`，未删除或移动 `release/` 任何产物。本目录下新增文件仅本文件；同目录已有的 `tools-snapshot.baseline.json`（P0.2 产物）未改动。
- 所有"能否"结论只依据本次实际命令输出与仓库内现有产物；找不到依据的写"未找到依据"，不推测。

---

## 1. 主机与工具链

### 1.1 主机与操作系统

```console
$ sw_vers
ProductName:		macOS
ProductVersion:		27.0
BuildVersion:		26A428

$ uname -m
arm64

$ uname -a
Darwin Mac 27.0.0 Darwin Kernel Version 27.0.0: Tue Aug 11 21:02:57 PDT 2026; root:xnu-13432.1.9~1/RELEASE_ARM64_T8132 arm64
```

**结论**：主机为 Apple Silicon（arm64）macOS 27.0（build 26A428）。本机是 macOS arm64 的原生真机环境，可承担 macOS arm64 的构建与实机运行验收。

### 1.2 Node / npm（与 CI 不一致）

```console
$ node -v
v25.9.0

$ npm -v
11.12.1

$ which node
/opt/homebrew/bin/node

$ which npm
/opt/homebrew/bin/npm

$ node -p "process.versions.v8"
14.1.146.11-node.25

$ which -a nvm volta fnm asdf
（均未找到）
```

CI 使用 Node 22：

```yaml
# .github/workflows/verify.yml
- uses: actions/setup-node@v4
  with:
    node-version: 22
    cache: npm
```

**结论（不一致，需登记）**：

- 本机 `node` 为 **v25.9.0**，CI 固定 **Node 22**，两者主版本不同。
- 仓库无 `.nvmrc` / `.node-version`（`ls` 报 No such file or directory），也没有 nvm/volta/fnm/asdf 等版本管理器，因此"CI 与本地同版本"目前**没有机制保证**。
- `package.json` 的 `@types/node` 为 `^22.13.10`，README 写明"需要 Node.js 22 LTS 或更高版本"；按字面 v25 满足"更高"，但不等于与 CI 等价。
- 影响：本地 `npm test` / `npm run build` 的通过不能直接等同于 CI Node 22 下的通过。P1.4/P1.6 与 P5.1 应记录实际使用的 Node 版本，或补 `.nvmrc` 固定到 22（属改造动作，本盘点不做）。

### 1.3 打包工具链版本

```console
$ node -p "require('./node_modules/electron/package.json').version"
35.7.5

$ node -p "require('./node_modules/electron-builder/package.json').version"
25.1.8
```

**结论**：Electron 35.7.5、electron-builder 25.1.8 已安装（`node_modules` 存在）。与 `release/_legacy-build/builder-effective-config.yaml` 中记录的 `electronVersion: 35.7.5` 一致。

### 1.4 PowerShell

```console
$ which pwsh
（未找到，exit 1）
$ which powershell
（未找到，exit 1）
$ ls /usr/local/microsoft
ls: /usr/local/microsoft: No such file or directory
```

**结论**：本机**没有 PowerShell**。因此 `tests/native-contracts.ps1` 和 `tests/windows-office.ps1` 在 macOS 本机**一条都不能执行**（连"仅语法解析"的默认模式也不行）。这两个文件当前唯一可执行环境是 CI 的 `windows-latest`。

### 1.5 Windows 交叉打包能力（仅检查，未安装）

```console
$ which wine
（未找到，exit 1）
$ which wine64
（未找到，exit 1）
$ which mono
（未找到，exit 1）
$ brew list --formula | grep -iE "wine|mono|qemu|virtualbox|vmware|docker"
（无匹配：未通过 brew 安装 wine/mono/qemu/VirtualBox/VMware/docker）
```

但 electron-builder 自带缓存**已经存在**，说明 Windows 打包链在本机真实跑过：

```console
$ ls ~/Library/Caches/electron-builder
nsis  winCodeSign  wine

$ find ~/Library/Caches/electron-builder/wine -maxdepth 2
.../electron-builder/wine/wine-4.0.1-mac/wine-home/...
.../electron-builder/wine/wine-4.0.1-mac/lib64/libwine.1.0.dylib
（electron-builder 自带的 wine 4.0.1 macOS 构建）

$ ls ~/Library/Caches/electron-builder/nsis
nsis-3.0.4.1  nsis-resources-3.4.1

$ ls ~/Library/Caches/electron-builder/winCodeSign/winCodeSign-2.6.0
appxAssets  darwin  linux  openssl-ia32  rcedit-ia32.exe  rcedit-x64.exe  windows-10  windows-6

$ ls ~/Library/Caches/electron
electron-v35.7.5-darwin-arm64.zip
electron-v35.7.5-win32-arm64.zip
electron-v35.7.5-win32-x64.zip
```

**结论（关键）**：

- 系统级 `wine` / `mono` **未安装**；PATH 中不存在，`/usr/local/microsoft` 不存在。
- 但 electron-builder 会自行下载并使用自己的 wine（`wine-4.0.1-mac`）、NSIS 与 winCodeSign（含 `darwin/` 与 `rcedit-x64.exe`）。这三样都已在缓存中，且 Electron 缓存中同时存在 `win32-x64` 与 `win32-arm64` 两个运行时包，`darwin` 只有 `arm64` 一个。
- 因此"macOS 本机能否产出 Windows 包"的答案是**能**，依据不是假设而是本机已有 Windows 目标产物 + 完整 Windows 打包缓存（详见第 4 节）。这是**交叉打包**能力，与"能否在 Windows 上运行验收"是两件事。
- 注意区分：`electron-builder` 的 wine 只是打包工具内部依赖，**不能**当作可运行 Windows 应用的环境。本机**未安装** CrossOver 之外的任何 Windows 兼容/虚拟化方案；CrossOver 26.0 虽存在（`/Applications/CrossOver.app`，Wine 系兼容层），但它不是 Windows 桌面，本次**不作为**验收通道，且未验证其可运行本产品与 Office。

### 1.6 是否可访问 Windows 机器或虚拟机

```console
$ which docker           → 未找到（exit 1）
$ which utm              → 未找到（exit 1）
$ ls -d /Applications/UTM.app          → No such file or directory
$ which vmware vmrun     → 未找到
$ ls -d /Applications/VMware*.app      → No such file or directory
$ which prlctl           → 未找到
$ ls -d /Applications/Parallels*.app   → No such file or directory
$ which qemu-system-x86_64             → 未找到（exit 1）
$ which VBoxManage                     → 未找到（exit 1）
```

补充排查（同样只检查是否存在）：

```console
$ for c in qemu-system-aarch64 qemu-img limactl colima podman nerdctl multipass vagrant; do which $c; done
（全部 not found）
```

**结论**：本机**没有任何可用的 Windows 桌面环境**——无 Docker、无 UTM、无 VMware、无 Parallels、无 QEMU、无 VirtualBox，也无 Lima/Colima/Podman/Multipass/Vagrant。`/Applications` 顶层列表（`ls /Applications`）中也没有任何虚拟化软件。**P5 的 Windows 侧实机验收在本机当前条件下不可执行**，属阻塞项（登记见第 6 节）。

### 1.7 Rosetta 2（影响 macOS x64 判断）

```console
$ ls -d /Library/Apple/usr/libexec/oah
/Library/Apple/usr/libexec/oah

$ arch -x86_64 /usr/bin/true
（exit 0）

$ sysctl -n sysctl.proc_translated
0

$ pgrep -fl oahd
15086 /usr/libexec/rosetta/oahd
```

**结论**：Rosetta 2 已安装且可用，可在本机转译运行 x86-64 macOS 二进制。这只意味着"macOS x64 包或许能在本机跑起来"，**不等于** Intel Mac 实机验收；如纳入 macOS x64，仍应在真实 Intel 机器（或明确接受"Rosetta 转译"作为降级证据）上验收。

### 1.8 签名与公证工具链可用性

```console
$ xcode-select -p
/Applications/Xcode.app/Contents/Developer

$ which codesign xcrun spctl
/usr/bin/codesign
/usr/bin/xcrun
/usr/sbin/spctl

$ xcrun --find notarytool
/Applications/Xcode.app/Contents/Developer/usr/bin/notarytool
```

**结论**：签名/公证**工具**齐备（codesign / spctl / notarytool 均存在），缺的是**分发用凭据**（见第 5 节）。

---

## 2. 宿主软件版本

### 2.1 WPS Office（macOS）

```console
$ /usr/libexec/PlistBuddy -c "Print :CFBundleShortVersionString" /Applications/wpsoffice.app/Contents/Info.plist
12.1.28496

$ /usr/libexec/PlistBuddy -c "Print :CFBundleVersion" /Applications/wpsoffice.app/Contents/Info.plist
28496

$ /usr/libexec/PlistBuddy -c "Print :CFBundleIdentifier" /Applications/wpsoffice.app/Contents/Info.plist
com.kingsoft.wpsoffice.mac

$ file "/Applications/wpsoffice.app/Contents/MacOS/wpsoffice"
Mach-O 64-bit executable arm64
```

WPS 加载项部署现状（文字/演示/表格三者均已部署 2.1.0）：

```console
$ ls -la ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons
Office Agent Bridge (文字)_            Office Agent Bridge (文字)_2.1.0
Office Agent Bridge (演示)_            Office Agent Bridge (演示)_2.1.0
Office Agent Bridge (表格)_            Office Agent Bridge (表格)_2.1.0
authaddin.json  jsaddinblockhost.ini  jsplugins.xml  jsaddinblockhost.ini.backup-*（多份）
```

### 2.2 Microsoft Office（macOS）

```console
$ /usr/libexec/PlistBuddy -c "Print :CFBundleShortVersionString" "/Applications/Microsoft Excel.app/Contents/Info.plist"
16.113

$ /usr/libexec/PlistBuddy -c "Print :CFBundleVersion" "/Applications/Microsoft Excel.app/Contents/Info.plist"
16.113.26091433

$ file "/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel"
Mach-O universal binary with 2 architectures: [x86_64] [arm64]
```

Office 加载项清单部署位置现状（已存在本项目清单）：

```console
$ ls -la ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
total 16
drwxr-xr-x@ 3 jolin  staff    96 Sep 21 14:50 .
drwx------@ 4 jolin  staff   128 Sep 20 15:53 ..
-rw-r--r--@ 1 jolin  staff  4492 Sep 21 14:55 wps-bridge-manifest.xml
```

Word / PowerPoint 实测：

```console
$ /usr/libexec/PlistBuddy ... "/Applications/Microsoft Word.app"        → NOT FOUND
$ /usr/libexec/PlistBuddy ... "/Applications/Microsoft PowerPoint.app"  → NOT FOUND
$ ls /Applications | grep -iE "word|powerpoint|microsoft"
Microsoft Edge.app
Microsoft Excel.app
$ mdfind "kMDItemCFBundleIdentifier == 'com.microsoft.Word'"        → 无结果
$ mdfind "kMDItemCFBundleIdentifier == 'com.microsoft.Powerpoint'"  → 无结果
```

补充排查：`/Applications/Microsoft Office` 不存在；`~/Applications` 为空；`/usr/local/bin` 无 Office 相关；LibreOffice / OnlyOffice / OpenOffice 均未安装。

### 2.3 宿主版本汇总

| 宿主 | 是否安装 | 版本 | 架构 | 依据 |
|---|---|---|---|---|
| macOS WPS Office（表格/文字/演示） | 已安装 | **12.1.28496**（CFBundleVersion 28496） | arm64 | `/Applications/wpsoffice.app` |
| macOS Microsoft Excel | 已安装 | **16.113**（build 16.113.26091433） | universal（x86_64 + arm64） | `/Applications/Microsoft Excel.app` |
| macOS Microsoft Word | **未安装** | — | — | 应用不存在，mdfind 无结果 |
| macOS Microsoft PowerPoint | **未安装** | — | — | 应用不存在，mdfind 无结果 |
| macOS Microsoft Outlook / OneNote | **未安装** | — | — | 应用不存在 |
| Windows WPS / Microsoft Office | **未安装（本机无 Windows）** | — | — | 本机无 Windows 环境，见 1.6 |
| LibreOffice / OnlyOffice / OpenOffice | **未安装** | — | — | `/Applications` 无对应 .app |

**注意（残留痕迹，不等于已安装）**：`~/Library/Containers/` 下存在 `com.microsoft.Powerpoint`、`com.microsoft.Powerpoint.widgetextension`、`com.microsoft.Word.widgetextension` 三个容器目录，但 `Microsoft Word.app` / `Microsoft PowerPoint.app` 本体不存在。这属于历史安装残留，**不能**作为 Word/PPT 可用性依据。因此：

- Microsoft Word / PowerPoint 的 macOS 路线（JXA 原生脚本、PPT 预览）在本机**无法验收**。
- 改造计划第 5 节里"Microsoft Word/PPT"整行在本机的可用性为**不可用**。

---

## 3. 测试资产清单

`npm test` 的实际范围是 `tests/*.test.ts`（`package.json`：`node --import tsx --test --test-concurrency=1 tests/*.test.ts`），即下列 6 个 `.ts` 文件；2 个 `.ps1` **不**被 `npm test` 收集，仅在 CI 的 `windows-latest` 上单独执行。

### 3.1 Node 测试（6 个）

| 文件 | 用途（一句话） | 真实宿主需求 | 本机可跑性 |
|---|---|---|---|
| [tests/addon.test.ts](../../../tests/addon.test.ts) | 用 `vm` + 手写假 `wps.EtApplication` 对象加载 `wps-addon/addon-core.js`，验证加载项使用已安装凭据、"零值/公式读回不丢"、写入前拒绝尺寸不一致、拒绝部分文档名。 | **模拟宿主可跑**（假 WPS 对象，不是真 WPS） | 可跑（需 `wps-addon/addon-core.js` 存在） |
| [tests/office.test.ts](../../../tests/office.test.ts) | 纯函数单测：`normalizeOfficeRequest` / `normalizeOfficeResponse` 的扁平参数→Office.js 结构与返回值归一化。 | **模拟宿主可跑**（无宿主交互） | 可跑 |
| [tests/platform.test.ts](../../../tests/platform.test.ts) | 临时目录下验证 WPS 插件索引 XML 合并与损坏拒绝、MCP 配置逐字节保留、状态检查不建目录、工具名唯一且强制 `host`、目标锁按会话/宿主隔离、原生运行器不做字符串插值。 | **模拟宿主可跑**（假 COM 集合对象、临时目录） | 可跑 |
| [tests/service.test.ts](../../../tests/service.test.ts) | 起真实 HTTP/MCP 服务，验证 loopback health 免鉴权、文档操作需鉴权、跨源拒绝、schema 前置拒绝、回滚拒绝覆盖后续修改、HTTP MCP 握手与真实 RPC 路由、双 stdio 客户端复用同一后台。 | **模拟宿主可跑**（WebSocket 模拟 WPS RPC，不连真 WPS） | 可跑，但需本机允许监听 `127.0.0.1` 随机端口（沙箱可能 EPERM，见 `tests/AGENTS.md` 避坑） |
| [tests/lifecycle.test.ts](../../../tests/lifecycle.test.ts) | 冷启动 stdio 拉起后台、stdio 客户端退出后后台仍存活；需从 `dist/bridge/cli.cjs` 启动。 | **模拟宿主可跑** | 可跑，但**依赖已构建的 `dist/bridge/cli.cjs`**（本机 `dist/bridge/cli.cjs` 存在） |
| [tests/ppt-layout.test.ts](../../../tests/ppt-layout.test.ts) | 从 `wps-addon/addon-core.js` 源码文本截取 `pptPageSize` / `fitGeneratedPptShapes` 段落后在 `vm` 中执行，覆盖 720×405、960×540、1440×810、720×540 四种页面几何、非法页面尺寸、长文本溢出告警、原生脚本工具精确目标名。 | **模拟宿主可跑**（纯几何/逻辑，无宿主） | 可跑；注意该文件仍按字符串截取源码（计划 P3.6 要求改为直接导入布局模块） |

### 3.2 PowerShell 测试（2 个）

| 文件 | 用途（一句话） | 真实宿主需求 | 本机可跑性 |
|---|---|---|---|
| [tests/native-contracts.ps1](../../../tests/native-contracts.ps1) | 加载 `resources/office/common.ps1` + `excel.ps1`，用 `FakeCollection` 假 COM 对象验证 `Get-Target` 的单工作簿/全路径解析、拒绝部分名与重名、`Convert` 对 text/0/null/0.72 的处理等原生脚本契约。 | **模拟宿主可跑，但需 Windows + PowerShell**（假 COM 对象，不需要真 Excel） | **不可跑**：本机无 PowerShell（1.4） |
| [tests/windows-office.ps1](../../../tests/windows-office.ps1) | 默认模式只解析 `resources/office/*.ps1` 语法；加 `-RunOffice` 才创建自己的 Excel 实例与中文名临时工作簿，覆盖值/公式、样式、条件格式、图表、透视表、排序、验证、截图与直接回滚。 | **需要真机**：`-RunOffice` 要求 Windows 桌面 + 已安装 Excel 的真实 COM 宿主（默认模式仅需 Windows + PowerShell） | **不可跑**：本机无 Windows、无 PowerShell、无 Excel COM |

### 3.3 真实宿主验证脚本（`scripts/`，按任务输入要求一并清点）

| 文件 | 用途（一句话） | 真实宿主需求 | 本机可跑性 |
|---|---|---|---|
| [scripts/verify-all-mcp-tools.ts](../../../scripts/verify-all-mcp-tools.ts) | 读 `~/.wps-bridge/token`，对 `127.0.0.1:19890` 以 `host: 'microsoft'` 逐个调用工具。 | **需要真实 Microsoft 宿主**（默认写死 `host: 'microsoft'`） | 本机 Microsoft 侧仅有 Excel 16.113，Word/PPT 未安装；且脚本会执行写操作，**不宜**作为常规验证直接跑 |
| [scripts/verify_all_tiers.ts](../../../scripts/verify_all_tiers.ts) | 对运行中的 Bridge 做四梯队端到端集成调用（HTTP `19890`，脚本未带 token 头）。 | **需要真实运行中的 Bridge + 真实宿主** | 同上，写操作脚本，需明确测试文档后再用 |
| [scripts/build_beautified_sheet.ts](../../../scripts/build_beautified_sheet.ts) | 对运行中的 Bridge 写入并排版真实表格。 | **需要真实宿主 + 真实表格** | 会修改真实文件，**不能**当只读检查跑 |
| [scripts/inspect.ts](../../../scripts/inspect.ts) | 启动本地 Bridge 服务并探测宿主连通性。 | 需端口 19890 可用；探测目标为真实宿主 | 会启动服务，属有副作用脚本 |
| [scripts/snapshot-tools.ts](../../../scripts/snapshot-tools.ts) | 只读契约快照导出（工具名、schema、能力、逐工具 SHA-256）。 | **不需要宿主** | 可跑（其产物 `docs/acceptance/2.1.0-p0p1/tools-snapshot.baseline.json` 已在 P0.2 生成） |
| [scripts/check-agents.mjs](../../../scripts/check-agents.mjs) | 只读检查 AGENTS.md 命名与路径引用。 | **不需要宿主** | 可跑（`npm run check:agents`） |
| [scripts/dev.ts](../../../scripts/dev.ts) | 启动 Vite + Electron 开发环境。 | 需桌面会话 | 非验证脚本 |
| [scripts/create-icons.mjs](../../../scripts/create-icons.mjs) | 生成图标素材。 | **不需要宿主** | 非验证脚本 |

### 3.4 CI 实际执行内容（`.github/workflows/verify.yml` 全文）

```yaml
name: verify
on: [push, pull_request]
jobs:
  build-and-contracts:
    strategy:
      matrix:
        os: [macos-latest, windows-latest]
    runs-on: ${{ matrix.os }}
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 22
          cache: npm
      - run: npm ci
      - run: npm run typecheck
      - run: npm run build
      - run: npm test
      - if: runner.os == 'Windows'
        shell: powershell
        run: ./tests/windows-office.ps1
      - if: runner.os == 'Windows'
        shell: powershell
        run: ./tests/native-contracts.ps1
```

**结论**：

- CI 覆盖：typecheck → build → `tests/*.test.ts`（两系统）；Windows 额外跑 `windows-office.ps1`（**不带** `-RunOffice`，即只做语法解析）和 `native-contracts.ps1`。
- CI **没有**做：`npm run dist` 打包、`npm run check:agents`、`npm run build:website`、`build:renderer` 之外的加载项构建、任何真实宿主 `-RunOffice` 验收。改造计划 P1.6 要求"CI 接入导航检查、官网构建及对应依赖安装；保留现有双系统检查"——即这三项**尚未**接入（本盘点只陈述现状，不代计划勾选）。
- CI 的 `npm test` 在 `macos-latest`（arm64 或 x64 由 GitHub 决定）与 `windows-latest`（x64）上运行，但均为**模拟宿主**，不构成真实办公宿主验收。

### 3.5 测试资产可用性小结

- **模拟宿主测试资产齐备**：6 个 `.ts` 覆盖加载项 RPC、Office.js 归一化、安装/配置合并、服务与鉴权、生命周期、PPT 布局。可在本机运行（`lifecycle.test.ts` 依赖已构建的 `dist/bridge/cli.cjs`）。
- **真实宿主测试资产不齐备**：Windows 侧只有 2 个 `.ps1`，且本机无法执行；macOS 侧**没有**任何自动化真实宿主验收脚本（`docs/validation.md` 的 macOS WPS 步骤 1–9 全部是人工步骤）。
- **桌面环境资产**：本机具备 macOS 桌面会话（可驱动 WPS 12.1.28496 与 Excel 16.113 做人工验收）；Windows 桌面**不存在**。

---

## 4. 发布架构矩阵（P0.6）

### 4.1 `release/` 现状（原始输出，未做任何增删移动）

```console
$ ls release
2.0.0  2.1.0  _legacy-build  README.md

$ ls -la release/2.1.0/mac
Office Agent Bridge-2.1.0-arm64-mac.zip            110347331
Office Agent Bridge-2.1.0-arm64-mac.zip.blockmap      115557

$ ls -la release/2.1.0/win
Office Agent Bridge-2.1.0-arm64-win.zip            124407082
Office Agent Bridge-2.1.0-win.zip                  124767057

$ ls -la release/2.0.0/mac
Office Agent Bridge-2.0.0-arm64-mac.zip            105793634
Office Agent Bridge-2.0.0-arm64-mac.zip.blockmap      111187
Office Agent Bridge-2.0.0-arm64.dmg                109897538
Office Agent Bridge-2.0.0-arm64.dmg.blockmap          118120

$ ls -la release/2.0.0/win
Office Agent Bridge 2.0.0.exe                       80785363
Office Agent Bridge-2.0.0-win.zip                  122018324
```

`release/_legacy-build/`（历史归档，按 `release/README.md` 不作为最新发布入口）：`builder-debug.yml`、`builder-effective-config.yaml`、`relocation-manifest.json`、`mac-arm64/`（仅 `.DS_Store`）、`win-unpacked/`、`win-arm64-unpacked/`、`.icon-icns/`、`.icon-ico/`。

### 4.2 产物真实架构判定（读 PE / Mach-O 头，而非只看文件名）

```console
$ file "release/_legacy-build/win-unpacked/Office Agent Bridge.exe"
PE32+ executable (GUI) x86-64, for MS Windows

$ file "release/_legacy-build/win-arm64-unpacked/Office Agent Bridge.exe"
PE32+ executable (GUI) Aarch64, for MS Windows

$ unzip -p "release/2.1.0/win/Office Agent Bridge-2.1.0-win.zip" "Office Agent Bridge.exe" | file -
/dev/stdin: PE32+ executable (GUI) x86-64, for MS Windows

$ unzip -p "release/2.1.0/win/Office Agent Bridge-2.1.0-arm64-win.zip" "Office Agent Bridge.exe" | file -
/dev/stdin: PE32+ executable (GUI) Aarch64, for MS Windows

$ unzip -p "release/2.0.0/win/Office Agent Bridge-2.0.0-win.zip" "Office Agent Bridge.exe" | file -
/dev/stdin: PE32+ executable (GUI) x86-64, for MS Windows

$ file "release/2.0.0/win/Office Agent Bridge 2.0.0.exe"
PE32 executable (GUI) Intel 80386, for MS Windows, Nullsoft Installer self-extracting archive

$ unzip -p "release/2.1.0/mac/Office Agent Bridge-2.1.0-arm64-mac.zip" \
    "Office Agent Bridge.app/Contents/MacOS/Office Agent Bridge" | file -
/dev/stdin: Mach-O 64-bit executable arm64

$ unzip -p "release/2.0.0/mac/Office Agent Bridge-2.0.0-arm64-mac.zip" \
    "Office Agent Bridge.app/Contents/MacOS/Office Agent Bridge" | file -
/dev/stdin: Mach-O 64-bit executable arm64
```

**判定结果**：

- `2.1.0/win/...-win.zip`（**不带** arch 字样）内部主程序为 **x86-64**，即 **Windows x64**。
- `2.1.0/win/...-arm64-win.zip` 内部主程序为 **Aarch64**，即 **Windows arm64**。
- `2.0.0/win/...-win.zip` 为 x86-64；`2.0.0/win/Office Agent Bridge 2.0.0.exe` 是 NSIS 安装器（自解压外壳为 80386 属 NSIS 正常现象）。
- 两个 mac 包的主程序均为 **Mach-O arm64**。
- 全仓库 `find . -iname "*x64*"`（排除 `node_modules`）**无任何结果**；除 `2.0.0/mac/*.dmg` 外无其他 mac 产物。

### 4.3 `package.json` build 配置现状

```jsonc
"build": {
  "appId": "com.antigravity.officeagentbridge",
  "productName": "Office Agent Bridge",
  "directories": { "output": "release/${version}/${os}" },
  "mac": { "target": ["zip", "dir"], "category": "public.app-category.productivity", "extendInfo": { /* NSDocuments/NSDownloads 用途说明 */ } },
  "win": { "target": ["zip", "dir"] },
  "artifactName": "Office-Agent-Bridge-${version}-${os}-${arch}.${ext}"
}
```

要点：

- `mac.target` / `win.target` 均为 **`zip` + `dir`**，**没有** dmg / nsis 目标，也**没有任何 arch 限定**——即目标架构由命令行参数或宿主架构决定（macOS arm64 主机上不带参数执行 `npm run dist` 只产出 mac arm64）。计划第 6 节给出的调用方式（macOS 主机 `npm run dist -- --mac --arm64`、Windows 主机 `npm run dist -- --win --x64`）与之一致。
- `artifactName` 模板符合 README 的命名约定（`Office-Agent-Bridge-<版本>-<平台>-<架构>.<扩展名>`，README 明确"即使是 x64 也不省略架构"）。
- **命名不一致（需登记）**：磁盘上现有产物是 `Office Agent Bridge-2.1.0-win.zip` 这类"产品名带空格 + 省略架构"的名字，与当前 `artifactName` 模板不符。`release/_legacy-build/builder-effective-config.yaml` 中**没有** `artifactName` 字段（且 `directories.output: release`），说明现有产物由**更早版本**的构建配置生成。后果：下一次 `npm run dist -- --win --x64` 会产出 `Office-Agent-Bridge-2.1.0-win-x64.zip`，与旧文件**并存**而非覆盖，`release/2.1.0/win/` 下将同时出现两套命名——P5.2 的"检查命名、架构、资源、入口及校验值"必须显式处理这个歧义，避免验收时混用旧包。

### 4.4 发布架构矩阵（P0.6 结论表）

| 目标平台/架构 | 依据来源（release 已有产物 / package.json / 历史版本） | 本次是否在矩阵内 | 本机能否构建 | 本机能否实机运行验收 | 备注 |
|---|---|---|---|---|---|
| **macOS arm64** | `release/2.1.0/mac/Office Agent Bridge-2.1.0-arm64-mac.zip`（Mach-O arm64，2.1.0 唯一 mac 产物）；`release/2.0.0/mac/` zip + **dmg** 均为 arm64；`package.json` mac target = zip+dir，无 arch 限制 → 宿主 arm64 默认产出 | **在** | **能**：`npm run dist`（默认即宿主 arm64）或 `npm run dist -- --mac --arm64` | **能**：本机就是 macOS 27.0 arm64 桌面，宿主 WPS 12.1.28496 已部署 2.1.0 加载项 | 2.1.0 相对 2.0.0 **少了 dmg**（当前 `mac.target` 无 dmg，属配置有意收窄，非丢失目标）；本架构是本机唯一能做完整"构建+运行"闭环的目标 |
| **Windows x64** | `release/2.1.0/win/Office Agent Bridge-2.1.0-win.zip` 内 exe = PE32+ x86-64；`release/2.0.0/win/Office Agent Bridge-2.0.0-win.zip` 内 exe = x86-64；`release/2.0.0/win/Office Agent Bridge 2.0.0.exe`（NSIS）；`release/_legacy-build/win-unpacked/…exe` = x86-64；计划第 6 节命令示例 `npm run dist -- --win --x64` | **在**（有实际产物，且是最早出现的目标） | **能**（交叉打包）：需显式 `--win --x64`；缓存已有 `electron-v35.7.5-win32-x64.zip`、`nsis-3.0.4.1`、`winCodeSign`（含 `rcedit-x64.exe`、`darwin/`）、electron-builder 自带 `wine-4.0.1-mac` | **不能**：本机无 Windows 桌面、无虚拟机、无 Docker（1.6），无法运行 PE 产物做 Office/WPS 实机验收 | 计划 P0.6 明确要求"不能默默丢弃已有目标"，故必须保留 |
| **Windows arm64** | `release/2.1.0/win/Office Agent Bridge-2.1.0-arm64-win.zip` 内 exe = PE32+ Aarch64；`release/_legacy-build/win-arm64-unpacked/…exe` = Aarch64；缓存 `electron-v35.7.5-win32-arm64.zip` | **在** | **能**（交叉打包）：`npm run dist -- --win --arm64` | **不能**：同上，且更难——需要 Windows on ARM 设备 | 该目标自 **2.1.0 起才出现**（2.0.0/win 只有 x64 与 NSIS 安装器）；属"已出现"目标，P0.6 要求保留 |
| **macOS x64** | **未找到依据**：`release/` 无 x64 产物（`find -iname "*x64*"` 空结果、无 dmg 之外的 mac 包）；`package.json` 无 arch 限定（即未声明支持）；`~/Library/Caches/electron/` 无 `darwin-x64` 运行时（只有 `darwin-arm64`、`win32-x64`、`win32-arm64`）；README / 官网 / CHANGELOG / git log（仅 4 个提交、无 tag）均无 macOS x64 或 Intel 声明 | **待决策**：默认**不纳入**，直到用户明确 macOS x64 属实际支持范围 | **能**：`npm run dist -- --mac --x64`（工具链具备），但需现下载 `electron-v35.7.5-darwin-x64.zip`（缓存中没有） | **不能算实机**：本机为 arm64 硬件；Rosetta 2 已装且可用（1.7），只能提供"转译运行"证据，**不等于** Intel Mac 验收 | 计划 P0.6 措辞是"如属于实际支持范围则补入"；依据缺失，故按"未找到依据，需决策"上报，**不擅自补入也不擅自丢弃** |

### 4.5 必须明确回答的三个问题

**问题 1：Windows x64 / Windows arm64 是否在矩阵内？依据是什么？**

**都在矩阵内。** 依据是 `release/` 中的**实际产物文件名 + 二进制架构判定**，不是推测：

- Windows x64：`release/2.1.0/win/Office Agent Bridge-2.1.0-win.zip`（内部 `Office Agent Bridge.exe` = `PE32+ executable (GUI) x86-64`）；`release/2.0.0/win/Office Agent Bridge-2.0.0-win.zip`（同为 x86-64）；`release/2.0.0/win/Office Agent Bridge 2.0.0.exe`（NSIS）。
- Windows arm64：`release/2.1.0/win/Office Agent Bridge-2.1.0-arm64-win.zip`（内部 exe = `PE32+ executable (GUI) Aarch64`）。
- 交叉印证：`release/_legacy-build/win-unpacked/`（x86-64）与 `release/_legacy-build/win-arm64-unpacked/`（Aarch64）两个解包目录同时存在；Electron 缓存同时存在 `win32-x64` 与 `win32-arm64`。

注意文件名陷阱：**不带 arch 字样的 `...-win.zip` 就是 x64**（架构被省略），带 `-arm64-` 的才是 arm64。仅凭文件名容易误判，本盘点已用二进制头确认。

**问题 2：macOS x64 的依据是什么？**

**未找到依据，需决策。** 已排查位置与结果：

- `release/` 全部产物：仅 `arm64`（mac 侧 zip ×2、dmg ×1，主程序均为 Mach-O arm64）。
- `package.json`：mac target 为 `["zip","dir"]`，无 `arch` 数组、无 x64 声明。
- `~/Library/Caches/electron/`：无 `darwin-x64` 运行时包。
- 仓库文档与历史：README「当前支持范围」表只列 macOS/Windows 宿主，不涉及 macOS CPU 架构；官网源码（`website/src`）无架构或下载链接文案；CHANGELOG 无相关条目；`git log --all -S'"--x64"' -- package.json` 无结果；仓库**无 tag**（`git tag -l` 空）。
- 唯一出现 "macOS x64" 字样的地方是计划本身（[refactoring-plan.md:86](../../refactoring-plan.md)）。

因此按计划 P0.6 的条件句处理：**不纳入矩阵，作为待用户决策项**（决策选项与影响见 6.3）。

**问题 3：macOS 本机能否产出 Windows 包？是否等于可在 Windows 上运行验收？**

- **能产出**。证据链完整：(a) `release/` 中已存在 mac 主机产出的 Windows x64 与 arm64 包及对应解包目录；(b) `~/Library/Caches/electron/` 同时存在 `electron-v35.7.5-win32-x64.zip` 与 `win32-arm64.zip`；(c) `~/Library/Caches/electron-builder/` 存在 `nsis-3.0.4.1`、`nsis-resources-3.4.1`、`winCodeSign-2.6.0`（含 `darwin/` 与 `rcedit-x64.exe`）与 electron-builder 自带 `wine-4.0.1-mac`。系统级 `wine`/`mono` 缺失不构成阻碍，因为 zip/dir 目标由 electron-builder 自有工具链完成。
- **不等于可在 Windows 上运行验收**。本机无任何 Windows 桌面、无虚拟机、无容器运行时（1.6），无法启动 PE 产物，更无法在其中安装 Office/WPS 并验证 COM/加载项。计划 P5 退出条件写明"缺少真实 Windows 或某架构设备时，该项待验收，**不能用交叉打包代替运行通过**"。因此 Windows x64 / Windows arm64 的"构建"可在本机完成，但"启动、服务、业务矩阵、故障矩阵"必须留待真实 Windows 环境，本机只能产出"待验收"状态。

---

## 5. 签名与分发要求

### 5.1 仓库内签名/公证配置现状（grep 原始结果）

在本仓库下列范围 grep：`src` `tests` `scripts` `docs` `office-addon` `wps-addon` `resources` `.github` `package.json` `tsup.config.ts` `vite.config.ts`（排除 `node_modules` / `dist` / `release`）：

```console
$ grep -rniE "codesign"           → 无匹配
$ grep -rniE "notariz"            → 无匹配
$ grep -rniE "CSC_"               → 无匹配
$ grep -rniE "signing"            → 无匹配
$ grep -rniE "signAndEdit"        → 无匹配
$ grep -rniE "hardenedRuntime"    → 无匹配
$ grep -rniE "entitlements"       → 无匹配
$ grep -rniE "provisioningProfile"→ 无匹配
$ grep -rniE "APPLE_ID"           → 无匹配
$ grep -rniE "APPLE_TEAM"         → 无匹配
$ grep -rniE "certificate"        → 仅 resources/certs/localhost.crt（本地 HTTPS 开发证书，非代码签名）
$ grep -rniE "identity|gatekeeper"→ 无匹配
```

补充：

- `package.json` 的 `build.mac` / `build.win` **没有** `identity`、`hardenedRuntime`、`entitlements`、`entitlementsInherit`、`gatekeeperAssess`、`notarize`、`signAndEditExecutable` 等任何字段（4.3 已列全文）。
- `.github/workflows/verify.yml` **没有**任何签名相关步骤或 secret 引用；且 CI 根本不执行 `npm run dist`（3.4）。
- `release/_legacy-build/builder-effective-config.yaml` 是**旧构建**的有效配置快照，其中同样没有签名字段——即已产出的包也是**未签名**构建的产物。
- README 明确写明："安装包签名与公证需要发布者凭据，源码不包含签名密钥。"计划第 6 节要求："签名、公证、企业安全策略和首次运行拦截属于发布条件，P0 明确目标；未提供所需凭据时单独标记，不通过关闭系统保护伪造验收。"

**结论：仓库内不存在任何代码签名或公证配置，也没有对应的 CI secret 注入点。当前发布包全部是未签名、未公证产物。**

### 5.2 本机凭据现状（只查本机钥匙串，未导出任何值）

```console
$ security find-identity -v -p codesigning
  1) <REDACTED-SERIAL> "Apple Development: <REDACTED> (<REDACTED-TEAM>)"
     1 valid identities found
```

**结论**：本机钥匙串中有 **1 张 Apple Development（开发）证书**，**没有** `Developer ID Application` 证书，也没有 `Developer ID Installer` / `Apple Distribution` 证书，也没有 `Mac Developer` 之外的任何分发身份。

由此：

- 可以做**本机开发签名**（自签/开发证书），但**无法**做面向分发的 Developer ID 签名，**因而无法通过公证（notarization）**——notarytool 要求 Developer ID Application 签名。
- 工具链齐备（`codesign` / `spctl` / `notarytool` 均存在，1.8），缺的是**证书与账号凭据**。
- Windows 侧：`winCodeSign` 缓存里只有 electron-builder 的 rcedit 工具与通用资源，**没有**任何 Windows 代码签名证书；仓库也无 `CSC_LINK` / `WIN_CSC_LINK` 配置。Windows 包当前为**完全未签名**。

### 5.3 是否阻塞

| 事项 | 现状 | 是否阻塞 P5 | 说明 |
|---|---|---|---|
| macOS 分发签名（Developer ID Application） | 无证书 | **阻塞"可分发"结论**，不阻塞"构建+本机实机运行验收" | 本机 `codesign --sign -`（ad-hoc）或开发证书可让应用在本机运行；但对外分发会被 Gatekeeper 拦截 |
| macOS 公证（notarytool） | 无凭据（缺 Apple ID / App Store Connect API Key / Team ID） | **阻塞** | 未公证的包在他人机器上首次运行会被拦截；计划明确"不能用关闭系统保护伪造验收" |
| macOS hardened runtime / entitlements | 仓库无配置 | **阻塞**（公证的前置条件） | 需先补 `hardenedRuntime: true` 与最小 entitlements，才能提交公证 |
| Windows 代码签名（Authenticode / EV） | 无证书、无配置 | **阻塞"无 SmartScreen 警告的分发"** | 未签名的 exe/zip 会触发 SmartScreen；不影响本机交叉打包 |
| 企业安全策略 / 首次运行拦截 | 未验证 | 待验收 | 属目标用户环境，需真实终端确认 |
| 本机开发自签（ad-hoc） | 可行 | 不阻塞 | 可用于本机验收运行，**不能**替代分发签名 |

### 5.4 需要用户提供的凭据（清单，本盘点不索取具体值）

1. **macOS 分发证书**：`Developer ID Application` 证书（.p12 或已导入钥匙串），用于 `codesign`。
2. **macOS 公证凭据**（三选一）：
   - App Store Connect API Key（`.p8` + Key ID + Issuer ID，推荐），或
   - Apple ID + App 专用密码 + Team ID，或
   - 已登录的 `notarytool` keychain profile 名称。
3. **macOS 签名配置决策**：是否启用 `hardenedRuntime`、需要哪些 entitlements（本产品涉及文件系统访问、加载项部署，可能需要相应豁免条目）。
4. **Windows 代码签名证书**：`.pfx`/`.p12` + 密码（标准或 EV），用于 electron-builder 的 `WIN_CSC_LINK` / `CSC_KEY_PASSWORD` 等价配置。
5. **凭据注入方式决策**：CI secret（`.github/workflows/verify.yml` 需新增签名步骤与 `npm run dist`）还是本机签名；两者都不做时，需明确接受"仅内部分发未签名包"并在发布说明中写明。
6. **（若决定纳入 macOS x64）** Intel 目标的分发策略：是否需要单独签名/公证该包（同一证书可签多架构）。

**注意**：以上凭据**均未**在本仓库或本机配置中存在（除 5.2 的开发证书外）。在用户提供前，任何"已签名/已公证"的说法都不成立。

---

## 6. 可用性与阻塞登记

本节严格区分"本次已实测可用""已确认不可用""需用户提供"。**不把未来承诺写成已完成**。

### 6.1 可用

| 项 | 状态 | 依据 |
|---|---|---|
| macOS arm64 真机环境（构建 + 运行） | **可用** | `sw_vers` 27.0 / `uname -m` arm64；桌面会话存在 |
| macOS WPS 12.1.28496 真机宿主 | **可用** | `/Applications/wpsoffice.app`；加载项已部署（表格/文字/演示 2.1.0） |
| macOS Microsoft Excel 16.113 真机宿主 | **可用** | `/Applications/Microsoft Excel.app`；wef 清单已部署 `wps-bridge-manifest.xml` |
| 模拟宿主测试资产（6 个 `.ts`） | **可用（可运行）** | 3.1；`service.test.ts` 需允许 loopback 监听，`lifecycle.test.ts` 需已构建 `dist/bridge/cli.cjs`（存在） |
| 只读验证脚本 `snapshot-tools.ts` / `check-agents.mjs` | **可用** | 无宿主依赖 |
| 签名/公证工具链二进制 | **可用** | `codesign` / `spctl` / `notarytool` 均存在 |
| 交叉打包 Windows x64 / arm64 的工具链 | **可用（历史已证）** | electron-builder + 缓存 wine/nsis/winCodeSign + 两个 win32 electron 运行时 + `release/` 中已有产物 |
| Rosetta 2 | **可用** | `arch -x86_64 /usr/bin/true` → exit 0 |
| 桌面会话（可做 macOS 人工验收） | **可用** | 本机为真实 macOS 桌面，非 headless |

### 6.2 已确认不可用

| 项 | 状态 | 依据 | 影响 |
|---|---|---|---|
| **Windows 桌面环境（真机或虚拟机）** | **不可用** | 无 docker / utm / vmware / prlctl / qemu / VBoxManage / lima / colima / podman / multipass / vagrant（1.6） | 见下方 6.4 专项说明 |
| Windows PowerShell | **不可用** | `which pwsh` / `which powershell` 均未找到 | 两个 `.ps1` 测试在本机**完全无法执行**，连语法解析模式也不行 |
| Windows Excel / WPS 宿主 | **不可用** | 无 Windows 环境 | 计划第 5 节"WPS Excel｜macOS/Windows""Microsoft Excel｜Windows 原生分支"的 Windows 侧全部待验收 |
| macOS Microsoft Word / PowerPoint | **不可用** | 应用不存在（2.2） | 计划第 5 节"Microsoft Word/PPT"整行在本机不可验收；容器残留不是安装证据 |
| 真实 Office 加载项在 macOS 上的结构化 Excel 适配 | **不可用（按 README 声明）** | README「当前支持范围」："macOS Microsoft Office ｜ JXA 原生脚本与状态通道 ｜ 结构化 Excel 适配暂不支持" | mac Excel 只能走原生脚本通道，不能作为 Office.js 结构化工具的真机验收依据 |
| 分发代码签名（macOS Developer ID / Windows Authenticode） | **不可用** | 仓库无配置（5.1）；钥匙串仅有 Apple Development 证书（5.2） | 打包可做，"可分发/已公证"结论不可做 |
| 真实宿主自动化验收脚本（macOS 侧） | **不可用（不存在）** | `tests/` 中无 macOS 真实宿主脚本；`docs/validation.md` 的 macOS WPS 步骤全为人工 | macOS 验收需人工执行，无自动化回归 |

### 6.3 需用户决策 / 需用户提供

| 项 | 类型 | 需要的动作 |
|---|---|---|
| **macOS x64 是否纳入发布矩阵** | **需用户决策** | 依据缺失（4.5 问题 2）。三选一：(a) 明确不纳入，在 README/计划中写明仅支持 Apple Silicon，今后也不再声称；(b) 明确纳入，则需指定 `--mac --x64` 为发布命令之一、指定 Intel 验收机器或明确接受 Rosetta 2 转译证据；(c) 暂缓决定，P5 只验收 mac arm64 并登记缺口。**不建议**在不决策的情况下让该目标从矩阵中静默消失。 |
| macOS 分发证书 + 公证凭据 | **需用户提供** | 见 5.4 第 1–3 项 |
| Windows 代码签名证书 | **需用户提供** | 见 5.4 第 4 项（若暂不提供，必须在发布说明中写明"未签名，会触发 SmartScreen"） |
| 凭据注入方式（CI secret vs 本机签名） | **需用户决策** | 见 5.4 第 5 项。当前 CI 连 `npm run dist` 都没有，接入签名前需先决定是否让 CI 打包 |
| Windows 验收机器 / 虚拟机 | **需用户提供** | 本机无法自建（6.2）。需一台 Windows x64 桌面；Windows arm64 另需 ARM 设备。或授权在本机安装虚拟化软件（本盘点**未安装**） |
| Node 版本策略 | **需用户决策** | 本机 v25.9.0 vs CI Node 22（1.2）。是否补 `.nvmrc` 固定 22，或把 CI 提到 25，或明确接受两者并行 |

### 6.4 专项：本机没有 Windows 桌面环境，如何影响 P5

计划 P5 的 7 个子任务中，**4 个在本机无法完成**：

| P5 子任务 | 本机可行性 | 说明 |
|---|---|---|
| P5.1 在 macOS/Windows 环境分别执行完整检查与打包 | **部分** | macOS 侧可做；Windows 侧只能做**交叉打包**，不是"在 Windows 环境执行" |
| P5.2 按矩阵生成 ZIP 与解包目录，检查命名/架构/资源/入口/校验值 | **部分** | 生成与静态检查（含 `file` 判架构、ZIP 内容、SHA-256）可在本机做；"在目标系统上解压完整"做不到 |
| P5.3 从实际发布包运行，验证桌面/CLI/stdio MCP/HTTP MCP | **仅 macOS** | Windows 包无法在本机启动 |
| P5.4 完成业务矩阵与故障矩阵 | **仅 macOS** | 表内所有标注"Windows"的行（WPS Excel、Microsoft Excel 原生分支、WPS Word、WPS PPT）在 Windows 侧全部待验收 |
| P5.5 加载项安装/升级、保留其他配置、后台重启、双客户端共存 | **仅 macOS** | Windows 分支（`resources/office/*.ps1`、`AddonInstaller` 的 Windows 路径）在本机无法执行 |
| P5.6 保存后关闭测试文件并重新打开，核验实际落盘内容 | **仅 macOS** | 本机有 WPS 与 Excel，可做；Windows 不可 |
| P5.7 汇总所有目标平台结果 | **受阻** | 缺 Windows 结果即无法满足"所有必验平台/架构通过" |

**直接影响**：

- 按计划 P5 退出条件——"缺少真实 Windows 或某架构设备时，该项待验收，不能用交叉打包代替运行通过"——**P5 在当前条件下不能被标记为完成**。可以完成的是 macOS arm64 的构建与实机验收，以及 Windows x64 / arm64 的**构建 + 静态核验**。
- 计划允许的推进方式（第 8 节末）："环境不足时可以推进不依赖该环境的阶段，但不能提前宣告整体验收通过。"因此 P0–P4 中不依赖 Windows 实机的部分可继续（P0 其余项、P1、P2、P3 的模块化与加载项构建、P4 导航），P5 的 Windows 部分保持未勾选。
- 另外，CI 有 `windows-latest` runner，可运行 `npm run typecheck/build/test` 与两个 `.ps1`，但 CI **不**执行 `npm run dist`，且 `windows-office.ps1` 在 CI 中**不带** `-RunOffice`（只做语法解析）。因此 **CI 不能替代 Windows 实机验收**：它既跑不了发布包，也跑不了真实 Excel COM。这一条容易被误读为"Windows 已在 CI 覆盖"，需在 P5 报告中显式澄清。

### 6.5 与改造计划的对照（只陈述现状，不代计划勾选）

- P0.5 要求"准备独立测试文件、模拟宿主与对应系统的桌面环境；记录 Office/WPS 版本"：
  - 独立测试文件——**已有**（6 个 `.ts` + 2 个 `.ps1`）。
  - 模拟宿主——**已有且可用**（`addon.test.ts` 的假 `wps.EtApplication`、`native-contracts.ps1` 的 `FakeCollection`、`service.test.ts` 的模拟 WPS RPC）。
  - 对应系统的桌面环境——**macOS 有、Windows 无**。
  - Office/WPS 版本——**已记录**（第 2 节）。
  - 结论：P0.5 **部分完成**；Windows 桌面环境缺失部分按计划"缺失环境登记为阻塞，不把未来承诺标成完成"处理。
- P0.6 要求"确认发布架构矩阵和签名要求"：
  - 架构矩阵——**已确认**：macOS arm64、Windows x64、Windows arm64 在矩阵内；macOS x64 未找到依据，待决策（第 4 节）。
  - 签名要求——**已确认现状**：无配置、无分发凭据，属发布条件阻塞（第 5 节）。
  - 结论：P0.6 的**确认**工作已完成（即本节及第 4、5 节），但签名凭据的**提供**属于用户侧未完成项。

---

## 7. 未确认项

以下项目本次**未**取得可判定证据，明确保持"未确认"，不做推测：

1. **现有发布包的实际签名状态**：`release/` 中的 mac zip / dmg 与 win zip 内部二进制是否带 ad-hoc 或开发签名，本次**未验证**——验证需要把包解出到工作区之外，超出本次"只读且只写一个文件"的授权范围。仓库层面已确认**无签名配置**（5.1），但"产物实际签名位"未取得证据。
2. **macOS x64 的产品支持意图**：仓库内找不到任何依据（4.5 问题 2）。是"从未计划"还是"计划过但未产出"，无法从现有材料区分，需用户确认。
3. **`release/2.0.0` 与 `release/2.1.0` 产物是否由本机、由哪个源码版本构建**：`release/README.md` 明确指出"目录版本号不代表已经包含最新源码，需以实际构建为准"。本盘点确认了架构与时间戳，但**未**核对包内 `app.asar` 与当前源码的对应关系。P0.1/P5.2 需补做。
4. **`release/2.1.0/win/…-win.zip`（x64）与 `release/_legacy-build/win-unpacked/` 是否内容等价**：两者的 `Office Agent Bridge.exe` 大小一致（201270272），但未逐文件比对。
5. **Windows arm64 包在真实 Windows on ARM 上的可运行性**：无设备，无法验证。
6. **`2.0.0/mac/*.dmg` 是否仍为当前支持的发布形式**：当前 `package.json` mac target 为 `["zip","dir"]`，不含 dmg；dmg 属历史产物。是否恢复 dmg 目标**未见决策记录**。
7. **Windows 交叉打包的完整性**：本机已有 Windows 产物作为"曾成功"的证据，但本次**未重新运行** `npm run dist`（明确禁止），因此"当前配置 + 当前源码仍能一次成功产出 Windows 包"这一点的**当前有效性未确认**。旧产物可能由早于当前 `package.json` 的配置生成（4.3 已观察到 `artifactName` 与 `directory.output` 的差异），不能直接外推。
8. **CI 是否真正在远端跑过**：`README.md` 明确"CI 配置已提供，未将其写入等同于已在远端运行"。本次未查询 GitHub Actions 运行记录，CI 绿否**未确认**。
9. **`~/Library/Containers/com.microsoft.Powerpoint` 等残留容器的成因**：只观察到目录存在，未能确认是曾装 PowerPoint 后卸载，还是其他 Office 组件共用的容器。
10. **CrossOver 26.0 能否运行本产品的 Windows 包**：未测试（且按 6.2 不视为验收通道）。若用户希望以 CrossOver 作为降级验收手段，需单独立项评估——它不能替代真实 Windows + Office 宿主。
11. **`office-addon/excel/manifest.xml` 与已部署的 `wef/wps-bridge-manifest.xml` 是否一致**：两者大小均为 4492 字节，但未逐字节比对。
12. **WPS 12.1.28496 与 Microsoft Excel 16.113 之外的历史版本行为**：计划第 5 节要求覆盖"Office 2019/2021/365 版本差异"，本机仅有单一版本，**未确认**。
