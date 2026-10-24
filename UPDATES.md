# 客户端在线更新与发布

维护者本地操作入口为 `docs/update-operations.local.md`（若存在）。它记录实际服务器入口、当前发布与续签时间，受 Git 忽略和归档排除保护，不是公开仓库的构建前提。公开文件只保留通用流程；不能把实际服务器目录、账号、会话或个人机器路径写进来。

## 产品约定

首次安装与在线更新分开。Windows 品牌 EXE 用于首次安装，或没有更新功能的旧客户端；在线更新只接收完整程序 ZIP，Windows 为 `windows-zip`，macOS 为 `mac-zip`。已发布且没有更新器的旧客户端需要手动安装一次支持更新的新客户端。

客户端启动后自动检查，运行中每六小时检查一次，并保留手动检查。发现新版本时弹窗提醒，同时主界面持续显示“有新版本”。同一版本只自动弹一次；关闭提醒不清除版本提示。“查看新版”只打开客户端并定位版本更新区域。

在客户端点击“下载更新”后后台下载，显示真实进度，可取消或重试，不占用其他功能的忙碌状态。完成签名、大小和哈希校验后显示“重启并更新”。点击该按钮才关闭 GUI 及它拥有的 Bridge，独立助手完成替换并重新打开。正常更新没有首次安装向导、选目录或手工解压。

关闭窗口只隐藏。托盘或菜单栏退出会暂停自动拉起，等待自有 Bridge 及原生辅助作业退出；独立启动的服务不接管、不按名称批杀。停止失败保留明确错误；新的桌面启动或显式 CLI `--start` 可以恢复后台启动。

## 后续 AI 的代码入口

| 职责 | 入口 |
|---|---|
| 签名清单、版本、来源白名单 | `src/main/update-protocol.ts`、`update-trust.ts` |
| 检查、后台下载、取消、校验、安装结果 | `src/main/update-service.ts` |
| 弹窗、自动检查、客户端定位、退出 | `src/main/index.ts`、`desktop-service.ts` |
| 一次性授权、撤销与交接 | `src/main/update-install.ts`、`update-handoff.ts` |
| 独立助手、Mac 替换和回退 | `src/main/update-helper.ts`、`update-mac.ts` |
| Windows ZIP、应用验证及版本目录事务 | `windows-bootstrap/update-zip.mjs`、`update.mjs`、`core.mjs`、`shortcut.mjs` |
| Windows 既有 MCP 路径迁移 | `src/main/update-config.ts`，由新版本 GUI 启动时调用 |
| IPC 与客户端更新区域 | `src/preload/index.ts`、`src/renderer/App.tsx`、`components/SettingsDrawer.tsx` |
| AI 配置及稳定 Skill 链接 | `src/main/client-registry.ts`、`mcp-adapters.ts`、`dsh-mcp.ts`、`skill-links.ts`；两端矩阵见 [CLIENTS.md](CLIENTS.md) |

Windows helper 是另一个只运行本地已安装助手的 Electron 进程，不进入 GUI 或安装向导，也不运行下载的可执行文件。它复用首次安装引擎的安全解压、独立版本目录和快捷方式备份。重用已有目标版本目录时仍重新验证程序并记录实际旧目录，随后迁移启动入口。旧目录和用户配置保留。

## 每次修改与发布的顺序

1. 先读根目录和目标目录的 `AGENTS.md`，检查工作区改动。保留现有工作；源码快照仅保存固定基线、必要差异及新增源码，排除密钥、配置、依赖和构建物。
2. 完成客户端代码、相关安全测试、代码审查和真实界面验证；更新 `CHANGELOG.md` 的第一段 `Unreleased`。实现完成不代表实机验收通过。
   变更 AI 接入时同步核对两端路径和官方配置格式；技能源在用户数据下保持稳定，不能把客户端链接指向某个版本的安装目录。新增依赖需同步锁文件并重新核对隔离依赖目录。
3. 确认正式版本后运行 `npm run version:set -- <新版本>`。它同步锁文件、Bridge、加载项源码/清单、界面默认版本和变更日志，不修改生成物、不提交 Git。随后检查 diff 并运行 `npm run check:version`。
4. 核对依赖。`npm run check:dependencies` 必须无差异。需要安装时在独立构建副本运行 `npm ci`，不得把“锁文件存在”当作已安装依赖一致。可用 `node scripts/check-dependencies.mjs <独立node_modules目录>` 检查已有依赖目录。
5. 在已审查的源码副本执行下面的构建命令。输出必须是新目录；脚本执行类型/导航检查、构建、自动化测试、发布保护和两平台串行打包，最后核验 ZIP 并扫描签名私钥泄漏。失败即停止，不沿用上一轮产物或哈希。

```sh
npm run build:update-packages -- <全新本地输出目录> all
# 单平台也可使用 mac 或 win；Mac 必须在 macOS 构建
```

输出 `mac/`、`win/` 和 `packages.json`。两端在线更新都上传这里的程序 ZIP，不能把 Windows 品牌 EXE 写入在线清单。该脚本不运行真实应用更新、不签清单、不上传。

构建门禁会用实际更新器在临时目录解压每个 ZIP，并验证应用版本/身份/架构。可独立重跑：

```sh
npm run check:update-package -- darwin <版本> <Mac ZIP>
npm run check:update-package -- win32 <版本> <Windows ZIP>
```

6. 分别做 Mac 和 Windows 实机升级验收，再使用已有离线 Ed25519 私钥签名。**不要重新初始化、轮换或复制现有生产私钥**。私钥仅供本地签名工具和隔离扫描器读取；不得进入 Git、源码快照、应用包、日志或服务器。公钥已经固定在客户端中。
7. 先上传新程序包并核对远端大小、SHA-256、HTTPS 下载，再上传临时清单并原子替换目标清单。没有服务器写入授权时停在本地产物与具体上传清单，不能把构建授权当成部署授权。

## 离线签名命令

维护者在 Git 忽略的本地目录准备 `spec.json`，填写此次真实版本、递增 `releaseId`、客户端已信任的 `keyId`、最长 31 天有效期和刚构建且已验收的 ZIP 路径。`packageBases` 是包含 `packages/` 的更新根 URL，必须使用客户端允许的 HTTPS origin。

```json
{
  "version": "2.3.0",
  "releaseId": 1,
  "keyId": "<客户端内已固定的keyId>",
  "expiresAt": "<发布后31天以内的ISO时间>",
  "notes": "本次更新说明",
  "packageBases": ["https://updates.example.com/updates/"],
  "assets": [
    {"platform":"darwin","arch":"arm64","kind":"mac-zip","file":"<本次Mac程序ZIP>"},
    {"platform":"win32","arch":"x64","kind":"windows-zip","file":"<本次Windows程序ZIP>"}
  ]
}
```

```sh
npm run prepare:update -- <spec.json> <已有私钥文件路径> <全新签名输出目录>
node scripts/check-update-secrets.mjs <全新签名输出目录>
```

工具复制包、复算 SHA-256，并生成 `stable.json` 和仅供维护者核对的公钥文件。不会生成密钥、联网或覆盖已有输出。输出的 `packages/<版本>/<含完整哈希文件名>.zip` 是不可变上传对象；`stable.json` 是最后提交的签名清单。只上传这两类文件，不上传 spec、源码、报告、私钥或用户数据。

签名工具可以在依赖已锁定的构建副本中运行，但 spec、程序包和私钥参数要指向明确的现有文件；私钥使用原工作区的绝对路径，不能为方便运行而复制到构建副本。

服务端只需 HTTPS 静态文件服务，不需要数据库或额外后台。清单不缓存，带哈希的程序包可长期缓存。清单到期前需以新的递增 `releaseId` 续签；同编号内容变化、回退、过期、平台缺失或 404 都会明确失败，不能显示“已是最新”。

## 单平台先发与后续补齐

先确认本次已验收的平台，以及正式更新通道中已经投入使用的平台。其他平台尚未启用此通道时，可以只列入本次已验收的平台资产；不能把未验收的 Windows 包随 Mac 一起发布。

清单的 `version` 对其中所有平台共同生效。不能在新版本清单里挂旧版本的另一个平台 ZIP，应用版本校验会拒绝它。已投入使用的各平台需要同一目标版本；若以后要让各平台独立演进，应先明确独立清单方案，不能直接遗漏既有平台造成更新错误。

同版本后续加入另一个平台时，保留已发布平台原来的 ZIP、大小、哈希和 URL，只添加新验收的资产并递增 `releaseId` 重新签名。不要重新打包或覆盖已经发布的同版本 Mac 文件。

## 清单续签（不需要重新打包）

1. 读取当前线上清单，核对签名、版本、`releaseId`、资产与到期时间。
2. 复用已发布的原始 ZIP，保持版本、平台、哈希和 URL；设置更大的 `releaseId` 和新的到期时间（最多 31 天），由工具填写新的发布时间。
3. 用已有私钥运行 `prepare:update`，写到全新本地目录；用客户端内置公钥再验证。不要手改签名清单、重新生成密钥或降低过期校验。
4. 确认远端不可变包仍匹配；包未变化时无需重复上传，只上传新的临时清单并原子替换正式清单。
5. 从正式 HTTPS 地址读回验签，再在对应正式客户端检查；同版本应显示“已是最新版本”。

构建与签名脚本不会自动定时续签，也不会自动上传。定期维护需要单独安排，不能把写了流程当作已经配置了无人值守任务。

## 发布完成的判定与故障处理

分别记录“本地包生成”“用户机器换装”“更新源上线”“Git 提交/推送/Release”，不能用其中一项代替其他项。上传现有已核验包不隐含 GitHub 发布授权。

更新源上线至少需要：远端 ZIP 完整下载后的大小/哈希一致、正式清单 HTTPS 200 且不缓存、客户端内置公钥验签通过、实际客户端检查结果正确。保留发布前清单、发布后清单、包哈希和运行证据，私密记录只写本地。

正式包检查更新遇到 404，先核对正式与验收源是否混用，以及正式清单是否实际上传；不要改成“已是最新”，也不要把验收清单复制到正式源。新版本或签名清单有问题时停止继续发布并保留证据，按更大的 `releaseId` 和正确资产修正；已经发布的程序有缺陷应发更高版本，不覆盖旧包或强制降级。服务器撤下清单不等于已经下载的有效签名包被撤销，不能把它当作客户端交接授权撤销。

## A/B 验收

使用本地计划和 `scripts/prepare-upgrade-acceptance.mjs <plan.json> <已核对依赖目录>` 创建全新 A/B 源码副本。计划需有 `id`、`baseCommit`、`versions.A/B`、独立 `appId`、`runtimeDirectoryName`、`windowsRootName`、`windowsShortcut`、两个非生产 `bridgePort/httpsPort`、独立 HTTPS `feed` 与 `assetOrigin`。计划及部署细节保存在本地 `docs/`，不要提交公开仓库。

A/B 的顺序是：完成源码测试与审查 → 构建并预检新的 A/B ZIP → 使用已有密钥生成 B 的验收清单 → 获准后仅上传隔离验收源 → 两端实际升级验收。正式发布再使用正式身份、正式版本及正式清单，不能把验收清单直接复制到正式地址。

脚本只在副本替换版本、身份、数据目录、端口和验收源，禁止办公加载项、客户端真实配置和开机启动写入，记录逐文件覆盖指纹，不复制私钥或部署。先在未覆盖验收参数的源码副本通过测试，再在 A/B 副本构建发布保护和 ZIP；验收覆盖会改变固定版本及运行目录，不能用它运行期望生产路径的全量单测。

实机必须记录：A 的真实进程/版本、提醒与持续提示、后台下载时其他功能可用、取消重试、下载校验完成但未点击时不退出、点击后旧 GUI 和自有 Bridge 退出、新应用和 Bridge 重启、B 的实际版本、设置保持。还要测关闭窗口隐藏、托盘退出不复活、独立服务不误杀及更新失败的可恢复状态。

局部安全回归：

```sh
node --import tsx --test tests/update.test.ts tests/update-install.test.ts tests/desktop-lifecycle.test.ts tests/windows-bootstrap.test.ts tests/update-key-safety.test.ts
```

生命周期测试依赖本次构建的 CLI。`tests/desktop-update-ui.cjs` 验证真实 renderer/preload 和隔离 IPC，不连接办公宿主或执行替换。模拟 Windows ZIP、Mac 测试、构建成功都不能代替 Windows 实机；没有该环境时明确写“未验收”。

在发布前把真实 ZIP 送入更新器预检：Windows 使用 `inspectUpdateZip` 并验证 PE/ASAR；Mac 使用 `stageMacUpdate` 在隔离安装副本暂存并验证。库对 STORE 和 DEFLATE 的流处理不同，小样本全是 DEFLATE 时不能证明真实包可读。Windows 物理 ASAR 解压还需用完整 Electron 运行 `tests/windows-bootstrap-electron.cjs <ZIP> <由预检生成的文件清单> <新的报告>`，它仍不是 Windows 系统验收。

## 失败恢复与不可删除的保护

- 清单签名、来源、大小、哈希、平台、架构、应用身份及版本都要验证。签名失败不执行包；ZIP 必须限制展开量并拒绝路径穿越、重复路径、特殊文件和越界链接。
- helper 就绪后 GUI 才授予一次性授权；GUI 退出后 helper 消费授权并在提交点复查。超时或失败永久撤销，迟到 helper 或后来的普通退出不能复活旧授权。重试使用新交接，禁止删除撤销记录来复用。
- 退出前提交 Bridge 启动暂停代次，延迟启动仍须复查，不能在退出后复活。启动协调锁异常时保守报错，不按 PID 或进程名猜测并清除。
- 下载/交接结果在运行目录 `desktop-state/updates/`。helper 失败尝试重新打开保留的应用；无法重开时给出系统错误提示。客户端显示最近安装结果和重试入口。启动命令成功只代表请求已发出，业务健康还须实测。
- Mac 在应用同盘的私有事务目录保留旧 `.app`；替换失败自动回退。磁盘映像、App Translocation、链接目录和权限不足明确失败，不提权、不关闭 Gatekeeper、不移除隔离属性。
- Windows 只更新当前用户且明确归属本引导的版本目录，不覆盖旧程序；快捷方式先备份再原子切换。配置迁移只修改已明确指向旧 EXE 的当前品牌 MCP 路径，保留自定义参数、禁用状态及其他服务；解析失败不覆盖。

Mac 断电发生在两个 rename 之间、原路径缺失时，可在确认相关字浮实例关闭后，用保留的运行时执行离线恢复：

```sh
ELECTRON_RUN_AS_NODE=1 "<保留应用>/Contents/MacOS/字浮 CharFloat" "<交接目录>/helper.cjs" --recover "<事务目录>/transaction.json"
```

恢复只把经验证的旧包移回缺失路径，不覆盖任何现存应用、不联网或启动 Bridge。Windows 可从保留旧版本启动，再重新检查、下载和确认更新。旧程序、配置备份和失败证据不能在自动化脚本里一并清理。
