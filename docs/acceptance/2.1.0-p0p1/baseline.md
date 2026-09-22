# P0 基线与验收环境快照

候选标识：`2.1.0-p0p1`
记录时间：2026-09-22
对应阶段：改造计划 P0（P0.1、P0.2、P0.3、P0.4）

> 本文件只记录**测量结果**。P0 不修改业务行为；本目录内一切结论都不代表实机验收通过。

---

## 1. P0.1 源码基线与工作区指纹

### 1.1 版本基线

| 项 | 值 | 采集命令 |
|---|---|---|
| 分支 | `main` | `git branch --show-current` |
| HEAD（完整） | `cac6d38e87791e53e61f3ecfc7286e0f39606f68` | `git rev-parse HEAD` |
| HEAD（短） | `cac6d38` | `git rev-parse --short HEAD` |
| 工作区索引对象 | `b6837a21e0c4d12cba5c5271117c4ea190cfe5de` | `git stash create`（只读创建对象，未修改工作区，未建立 ref） |
| 未提交修改规模 | tracked 修改 11 个 + untracked 16 个 | `git diff --name-only`、`git ls-files --others --exclude-standard` |

`git stash create` 产生的是游离对象，可能被 `git gc` 回收，且**不包含未跟踪文件**；**权威可复现标识请以文件 SHA-256 表为准**，索引对象仅作辅助。

> **可恢复性**：§1.2 的指纹只是核验，不能恢复。可恢复的基线内容（22 个文件）已另存为 [baseline/baseline-prior.patch](baseline/baseline-prior.patch) + [baseline/MANIFEST.sha256](baseline/MANIFEST.sha256)，**锚点提交固定在 [baseline/BASELINE.json](baseline/BASELINE.json)**，并可用 [baseline/verify-baseline.mjs](baseline/verify-baseline.mjs) 实测恢复。**有 5 个文件不可恢复**，原因与名单见 [baseline/README.md](baseline/README.md) §无法恢复的部分。

### 1.2 前序未提交修改（本次会话之前已存在，必须保留）

**tracked 修改 11 个**

| 文件 | SHA-256 |
|---|---|
| `README.md` | `bc7f76e08632413cd080227a7ab288bbdceaa39499c40800d05ca2dfc729d18d` |
| `package.json` | `4032ed0b429c9213cf44151ca6b04bd02a651b8f7c8a0a83d5a30b5635197c84` |
| `skills/office-agent-bridge-chart-style/SKILL.md` | `abae4d9d45c7c5b566c7e40530dfc5583d8dff11053ca1006d6f5330e8db3317` |
| `skills/office-agent-bridge-ppt-design/SKILL.md` | `32ae2b1e3a0eced253a984bf9d68eb1f0a32fcf7f42300f18f3d441422fd8d82` |
| `skills/office-agent-bridge-word-batch-edit/SKILL.md` | `a25ca16b500629d43768296546ee838830af3c36faf64c0bf45183e03813baab` |
| `skills/office-agent-bridge/SKILL.md` | `46f7a55d7eda1458a211c1ac1d420c4f028b67d2d52155adae7a4cc4d2a2286d` |
| `skills/office-agent-bridge/references/capabilities.md` | `b68738e195b27f1a578112c11389c1bf0c03f77490725d6c29ad9e601f50dbde` |
| `skills/office-agent-bridge/references/operations.md` | `344dc46521c3ad66655e4a567bd336fcb3a2b9224c726450f5524a1172812a5c` |
| `src/bridge/gateway.ts` | `d801ebe08e3a0374af30b9c7ef64cb6c3afddfec95143d9c546caf642bf2def2` |
| `src/bridge/mcp-server.ts` | `1ede09d75127017ad021bc1ff49d0420a2d42809f2c4c2e8adadbbfc1c89dc31` |
| `wps-addon/addon-core.js` | `6450c171cdbbb6dc121b8bd7e371ec5efc39aee932d4e1158783562064735906` |

**untracked 16 个**

| 文件 | SHA-256 |
|---|---|
| `AGENTS.md` | `f25c93d6d8c8b43073f5bc0877f17c0623dd3e8f337e6dcdc3972c8e6a4405b2` |
| `docs/refactoring-plan.md` | `2d0d5087b06d0fdaf55f569ecca73281c2b1d886324abc3b957a35d0d8ae57d0` |
| `office-addon/AGENTS.md` | `336361af80a79ce068fec43aee9bd769373fe24848b44a1d50078a39aa8df503` |
| `resources/office/AGENTS.md` | `f68db7cc63b26cdda7d60facd968bf39692ea759a19fa5b0912cf099ae2bd94c` |
| `scripts/AGENTS.md` | `0f8915cda7a49371b7fce38797c0500d36f31888a902572320ef2feb15eb267c` |
| `scripts/check-agents.mjs` | `aeab1289ec12e9d3b2f5a2ee2964e19c1313e401025d77636a9f9a69482931fd` |
| `skills/AGENTS.md` | `fd10e309c62201adda35cc810cec16710136256725b21a04d286528f30541459` |
| `skills/office-agent-bridge/references/native-scripting.md` | `da9abcfeb7a4d894ff359bd5654eb2f51572921ba791658f8396c77b504af1f6` |
| `src/bridge/AGENTS.md` | `96e38b98161b764028b2df9bce9c1c089ddfc6eedec68b0f2b11936e7a4dcf40` |
| `src/main/AGENTS.md` | `f0841ff3908a0106eeeb3e43764f32e2a119daa656f096d52df7b54283c02850` |
| `src/preload/AGENTS.md` | `a2403752090a3f2582aee3eaeeb450c5f21ab9f72c690346b56581e043c413e0` |
| `src/renderer/AGENTS.md` | `bdbe078501153593d7c9b4ce1279b6c920bc0a91d823fb750b74a7292fbe4f8d` |
| `tests/AGENTS.md` | `8031f4def86ee1b2790ac8f054e0a83716b05800551ff2a85585bd490688cf07` |
| `tests/ppt-layout.test.ts` | `4b658591b8b393e4a0bbff761f3b7799843d20d44c9310f3c2485f33c9fa3f01` |
| `website/AGENTS.md` | `b1e2660c003d0d42d939e9b02b861002fa890cdf0c855de210e5d4dc658ab6ae` |
| `wps-addon/AGENTS.md` | `7f522f2d3c6dc0689b6240a3e3adc1be5d34b136b6c91e314f82c98dca9bd70a` |

**关于 `package.json` 指纹的诚实说明**：本次会话在采集指纹之前，已为 P0.2 的导出脚本新增了 `snapshot:tools` 一行。因此上表给出的 `4032ed0b…` 是**还原后的真实基线值**（由当前文件内容移除该行后重算 SHA-256 得到），而不是指纹命令直接打印的值（直接打印得到的是 `bc7f76e0…`，已包含本次新增行）。其余 26 个文件的指纹均为直接测量值，未受本次修改影响。

### 1.3 本次修改（P0/P1，与前序修改分离）

**新增文件**

| 文件 | SHA-256 | 归属 |
|---|---|---|
| `src/bridge/errors.ts` | `d5e4a66f973c16fd4144606ba0fc839619b600368ded1d06cbf1f7ca151b121c` | P1.2 / P1.3 |
| `tests/failure-routing.test.ts` | `c9f84d9006138e2f1931ab85d13f284a461deb1c8da4cc28d2acbef1b2c34900` | P1.4 |
| `scripts/snapshot-tools.ts` | `b1531a03e1527ca91f6150a7fc361884ceb7a7decadfe5b02fb4b09388f7cb67` | P0.2 |
| `docs/acceptance/2.1.0-p0p1/`（本目录全部文件） | 见各文件 | P0 / P1 证据 |

**本次修改的既有文件**

| 文件 | 改动性质 |
|---|---|
| `package.json` | 仅新增 `snapshot:tools` 脚本一行（P0.2） |
| `.github/workflows/verify.yml` | 新增导航检查与官网构建作业（P1.6） |
| `src/bridge/ws-server.ts` | 补 `appendServiceLog` 导入（P1.1）；通道边界抛出已分类错误（P1.2） |
| `src/bridge/office/adapter.ts` | 拆分失败阶段；跨通道回退改为策略判定（P1.3） |
| `src/bridge/catalog.ts` | 错误分类接入；能力声明如实化（P1.2 / P1.5） |
| `docs/refactoring-plan.md` | 勾选与台账更新 |

`src/bridge/catalog.ts`、`src/bridge/ws-server.ts`、`src/bridge/office/adapter.ts`、`.github/workflows/verify.yml` 在会话开始时**无未提交修改**（与 HEAD 一致），本次改动不会覆盖任何前序工作。

---

## 2. P0.2 兼容快照

### 2.1 导出方式

```bash
npm run snapshot:tools -- docs/acceptance/2.1.0-p0p1/tools-snapshot.baseline.json
```

脚本 `scripts/snapshot-tools.ts` 为**只读**：直接 import 真实实现（`catalog.getTools()`、`catalog.capabilities()`），不监听端口、不连接宿主、不读写用户文档，并剥离 `lastUpdated` / `pid` / `port` 等易变字段以保证重复导出结果一致。工具名重复时脚本直接抛错，避免产出不可信快照。

### 2.2 快照文件

| 文件 | 说明 | SHA-256 |
|---|---|---|
| `tools-snapshot.baseline.json` | P0 基线快照（P1 改动前） | `a016c01061813db775841f0b1637c1c73077cb8a58374b2e1f38ac400963b9da` |
| `tools-snapshot.p1.json` | P1 完成后快照，用于差异比对 | 见同目录 |
| `business-samples.json` | 模拟宿主实际执行的原始 RPC 请求/响应样例 | 见同目录 |

### 2.3 基线计数

| 项 | 值 |
|---|---|
| 对外工具总数 | **91** |
| `readOnlyHint = true` | 17 |
| `readOnlyHint ≠ true` | 74 |
| `EXCEL_METHODS` 宿主方法数 | 30 |
| 实际生成 `excel_*` 工具数 | 27 |

### 2.4 业务样例口径

`business-samples.json` 由 `wps-addon/addon-core.js` 在 `vm` 模拟宿主中**实际执行**得到（与 `tests/addon.test.ts` 相同的夹具），记录了三条真实往返：成功读取（含零值与公式）、写入维度不一致被拒（写入前校验）、目标文档名不完整被拒（不误命中）。这是**模拟宿主**证据，不是真实 WPS 进程验证。

---

## 3. P0.4 基线检查结果

采集环境：macOS 27.0（26A428）arm64，Node v25.9.0，npm 11.12.1。

| 检查 | 命令 | 基线结果 | 退出码 | 与 P1 后对比 |
|---|---|---|---|---|
| 类型检查 | `npm run typecheck` | **失败**：`src/bridge/ws-server.ts` 4 处 `TS2304: Cannot find name 'appendServiceLog'`（基线行号 256/269/293/300） | **2**（实测） | P1.1 修复后通过（退出码 0） |
| 单元/集成测试 | `npm test` | 通过：25 项，fail 0，skipped 0，todo 0 | 0 | P1 后 **35** 项全通过（基线 25 + `failure-routing` 10，未跳过任何原有用例；见本文件"测试构成"） |
| 主构建 | `npm run build` | 通过（renderer + tsup 三个入口） | 0 | 未变化 |
| 官网构建 | `npm run build:website` | 通过：1901 modules，`dist/` 产出 | 0 | 未变化 |
| 导航检查 | `npm run check:agents` | 通过：12 个 AGENTS.md | 0 | P1 后 12 个（新增文件未新增 AGENTS.md） |

**类型检查退出码的复现方式**（因为测量时该错误已被 P1.1 修复）：临时从 `src/bridge/ws-server.ts` 的既有 `./runtime.js` 导入行移除 `appendServiceLog` 名字，执行 `npx tsc --noEmit`，测得退出码 **2** 与 4 条 `TS2304`（行号 257/270/294/301，因 P1 增加了一行 import 而整体 +1），随后按 SHA-256 校验恢复原文（`RESTORE_OK`）。

### 3.1 关键基线事实

1. **类型检查失败而构建通过**：`tsup` 不做类型检查，因此 `npm run build` 退出码 0 **不能**证明类型正确。修复前若只看构建会漏掉 P1.1 的真实缺陷。这条也解释了为什么 CI 必须同时保留 `typecheck` 与 `build`。
2. `appendServiceLog` 在 `src/bridge/runtime.ts:103` 导出，已被 `src/main/index.ts`、`src/main/addon-installer.ts` 正常引用；只有 `ws-server.ts` 漏了导入，属**漏导入**，不是函数缺失。修复为最小改动（在既有 `./runtime.js` 导入行补一个名字），未新增实现。
3. 历史记录的“25 项测试通过”确认为真（本次基线重跑亦为 25 项通过），但按计划要求，它**不代表本候选版本的发布验收**。

### 3.2 环境限制与不一致

| 项 | 现状 | 影响 |
|---|---|---|
| Node 版本 | 本机 v25.9.0，CI 固定 Node 22，仓库无 `.nvmrc`/`.node-version` | 本地通过 ≠ CI 通过；本文件的检查结果只对本机版本成立 |
| Windows 桌面环境 | 本机无任何 Windows 环境（无 docker/UTM/VMware/Parallels/QEMU/VirtualBox/lima/podman/multipass/vagrant，无 `pwsh`/`powershell`） | P5 的 Windows 侧实机项全部待验收 |
| 签名凭据 | 仓库无 `codesign`/`notarize`/`CSC_`/`hardenedRuntime`/`entitlements` 配置；钥匙串无 Developer ID Application | 分发签名与公证待用户提供凭据 |
| 宿主软件 | WPS 12.1.28496、Microsoft Excel 16.113 已安装；Microsoft Word / PowerPoint 未安装 | Microsoft Word/PPT 实机项无法在本机验收 |

详细环境清单与发布矩阵见 [environment.md](environment.md)。

---

## 4. P0.3 工具与宿主映射清点

结论与逐工具表格见 [tool-inventory.md](tool-inventory.md)（91 行主表 + 9 条差异 D1–D9 + 4 条同源发现 A1–A4 + 未确认项 9 条）。

> **修订适用性说明**：该清单在本次 P1 修改过程中产出，正文引用的行号对应 **P0 基线源码**（HEAD `cac6d38` + 前序未提交修改）。其中的 D1、D2、D3、D6、D7 已在 P1.5 修正，见 [compatibility-diff.md](compatibility-diff.md) §3；引用行号时请以基线为口径。

---

## 5. P1 补正后最终复验（冻结态）

P1 行完评审补正 1–5 后重跑，命令与退出码均为实测：

| 检查 | 命令 | 退出码 | 结果 |
|---|---|---|---|
| 导航检查 | `npm run check:agents` | **0** | 12 个 AGENTS.md；路径、根导航、维护入口、npm 命令有效 |
| 类型检查 | `npm run typecheck` | **0** | 无输出（基线 4 处 TS2304 已消除） |
| 测试 | `npm test` | **0** | tests 35 / pass 35 / fail 0 / **skipped 0** / todo 0 |
| 主构建 | `npm run build` | **0** | renderer + main/preload/cli 三个入口 |
| 官网构建 | `npm run build:website` | **0** | 1901 modules（依赖由 `website/package-lock.json` 锁定） |
| 官网干净安装 | `npm ci --prefix website` | **0** | 193 包；随后官网构建产物资源名与安装前一致 |
| 快照可复现 | `npm run snapshot:tools -- /tmp/rc.json` 后 `cmp` | — | 与 `tools-snapshot.p1.json` **逐字节一致** |
| 基线可恢复 | `node docs/acceptance/2.1.0-p0p1/baseline/verify-baseline.mjs` | **0** | 按固定锚点 `cac6d38` + 补丁恢复后 **22/22** 文件 SHA-256 一致 |
| **HEAD 前进后仍可恢复** | `bash docs/acceptance/2.1.0-p0p1/baseline/verify-after-head-advance.sh` | **0** | 隔离副本仓库 HEAD 前进为新提交（`91e40ad…`），仍 **22/22** 通过 |
| 失败复现（不覆盖工作区） | `bash docs/acceptance/2.1.0-p0p1/evidence/reproduce-prefix-failure.sh` | **1**（预期） | 副本内 `tests 11 / pass 9 / fail 2`；工作区 `adapter.ts` 全程未被触碰 |

测试构成：基线 25 项 + `tests/failure-routing.test.ts` 10 项 = 35 项。**没有删除任何测试、没有跳过任何用例、没有放宽任何既有断言。**

复核本次修改的文件指纹：

| 文件 | SHA-256 |
|---|---|
| `src/bridge/errors.ts` | `2d1d53dcd2860db43f019695b5c6f665a2daffe4a8a604ea40ac034b7e94fa56` |
| `src/bridge/catalog.ts` | `faf5c811682ac5a8709aef7aaeaed8d688686e984fdd883df3561133e9f85895` |
| `src/bridge/ws-server.ts` | `43ce2c3fdeccddc38f1ea04a0c8c2048a2d7ee2a9b892bf9032468411b53e506` |
| `src/bridge/office/adapter.ts` | `30e7988fb185ffac63513d29e773600b7cf64c73515966bc47ead6f13623c817` |
| `tests/failure-routing.test.ts` | `6cbebfe5a786d57ee6e3068b408bcc06433a68115ad24171264fe61b1616f864` |
| `scripts/snapshot-tools.ts` | `b1531a03e1527ca91f6150a7fc361884ceb7a7decadfe5b02fb4b09388f7cb67` |
| `.github/workflows/verify.yml` | `3467767b5282bae9a621107df388d56b57c2981f50c06a1c3933fe8186c153a1` |
| `package.json` | `99ec42a8dee33c7ca591824da4a9992535131381b7382fa9eca9089a7b455fe9` |
| `website/package-lock.json` | `7b2f17fa16ba086a79526a732e21ff6db9d2878c418e5ddbdb04f22cff35c6d1` |
| `tools-snapshot.p1.json` | `2b1abe2df7822231f08950d7a2d7dbf49b57d7a9e3c26cd66a467a4d7371d5eb` |
| `baseline/baseline-prior.patch` | `3c27956f641f7316c4d050ebe0ce2e516fe5d41f478772863201e706ba963f42` |
| `baseline/BASELINE.json` | `4d077dfb5e6708181bcd5ea27cdd90879d878e8c4eed2756904ca1d0e8ca486e` |
| `baseline/MANIFEST.sha256` | `ec60b2fcc4f225c185963aa1caa4831f90c7d9fc2b32f7455b69ee82d339763c` |
| `baseline/verify-baseline.mjs` | `4e5f7681e5eb7fd1028060601ae04166ce4704548ec59c5ac7e8d65401e36a0e` |
| `baseline/verify-after-head-advance.sh` | `eb6befda551a1dd41410a3cbabf3b566148de91ce3804607e6979bcc98bce64c` |
| `evidence/reproduce-prefix-failure.sh` | `e46bfc3e6c3a28f9149f079ca335d6027b0112c6b75e58b08cb1ae2a07cf4c02` |
| `evidence/baseline-head-anchor-prefix-failure.txt` | `7ddf28ed2d140237f8fcba3c8c15d42c82d73488f084fd3bd164e597d464ec75` |
| `evidence/p1.4-prefix-failure.txt` | `d3dae8b3f9a5308836102bc93123e840c0836757a319c4f1df3de22c2c68e4a8` |

> `src/bridge/catalog.ts` 的指纹为 P1.5 最终态；`tool-inventory.md` 中针对该文件的行号引用对应**基线**，见该文件顶部的修订适用性说明。

### 5.1 评审补正内容（1–5）

| # | 补正 | 结果 |
|---|---|---|
| 1 | `inspect_api` 移出只读集合与跨通道重放集合 | `wps_inspect_api` 的 `readOnlyHint` 保持 **false**（与基线一致）；不在 `REPLAY_SAFE_METHODS`。只读计数由初版的 22 回正为 **21**（实际更正 4 个工具）。新增副作用表达式回归测试，断言"结果未知时不得重放、确认未执行时仍可回退" |
| 2 | `routeOfficeFailure` 在替代通道不支持时保留原始错误与 `executed` | 不再硬编码 `kind='rejected'` / `executed='no'`，只追加"未回退：不在 Windows 原生通道支持列表内"。新增三项断言覆盖 `unavailable`→`no`、`failed`→`unknown`、`unknown`→`unknown` |
| 3 | 补齐可恢复基线 | 新增 `baseline/`：`baseline-prior.patch`（22 文件）+ `MANIFEST.sha256` + `verify-baseline.mjs`。实测从 HEAD 恢复后 22/22 哈希一致。**5 个未跟踪且已被后续修改的文件如实登记为不可恢复，未重造内容** |
| 4 | 修复前日志纳入可跟踪证据 | 由 `.log` 改为 `.txt`（`.gitignore` 忽略 `*.log`）；`git check-ignore` 确认不再被忽略。重跑复现输出 `tests 11 / pass 9 / fail 2`（与证据文件"结果签名"及复现脚本 `EXPECTED_*` 一致）。产生方式与断言含义见 `evidence/README.md` |
| 5 | 统一阶段状态并更新台账 | 见 [refactoring-plan.md](../../refactoring-plan.md) 顶部状态、阶段台账与台账记录；同时记录 DP1–DP6 决议 |
| 6 | **恢复材料与 HEAD 解耦**（评审二轮） | 新增 `BASELINE.json` 固定锚点提交；恢复与失败复现均改读锚点而非 `HEAD`；README 恢复命令改用仓库根目录解析的绝对路径（无个人路径）；失败复现在临时副本执行，不覆盖工作区；新增隔离验证证明"HEAD 前进后仍 22/22" | `baseline/BASELINE.json`、`baseline/verify-baseline.mjs`、`baseline/verify-after-head-advance.sh`、`evidence/reproduce-prefix-failure.sh`、`evidence/baseline-head-anchor-prefix-failure.txt` |

### 5.2 本次未做（不得视为已通过）

- **没有任何实机验收**：没有启动真实 WPS / Microsoft Excel 执行任何工具；所有结论来自静态检查、构建与模拟宿主。
- **Windows 侧零验证**：`adapter.ts` 的新回退策略只在桩替换下被验证，未在 Windows + 真实 Office.js + 真实 COM 上跑过。
- **CI 未远端实跑**：`verify.yml` 只做了本机 YAML 解析校验与命令本地实测。
- **发布包未重新构建**：`release/` 下产物与当前源码的对应关系未核；`npm run dist` 未执行。
- `insert_dimension` 无分支、`clear_range` 有分支无 schema、`wps_eval_code` 等 6 个死分支均**未改动**，只登记未处理。
- 基线中 5 个未跟踪文件**无法恢复**（详见 `baseline/README.md`）。

---

## 6. 未完成项

- P0.5 判定为**部分完成**：独立测试文件、模拟宿主、Office/WPS 版本齐备；**Windows 桌面环境缺失**，属登记阻塞。
- P0.6 的“确认”工作已完成：矩阵为 macOS arm64 + Windows x64 + Windows arm64（**DP1 已决议**：本轮沿用这三个目标，Intel Mac 暂不扩大，Rosetta 转译不能代替 Intel 原生验收）。原始依据见 [environment.md](environment.md) 发布架构矩阵一节。
- **DP2 已决议**：内部运行验收与正式分发要求**分开记录**；需要签名/公证凭据时再处理，当前不作为验收门槛。
- 本机可交叉产出 Windows 包，但**不等于**可在 Windows 上运行验收。
- 现有发布包与当前源码的对应关系未核；Windows 交叉打包在本次未复跑；CI 是否在远端实际跑过未查。
- **基线有 5 个文件不可恢复**：`AGENTS.md`、`docs/refactoring-plan.md`、`scripts/AGENTS.md`、`src/bridge/AGENTS.md`、`tests/AGENTS.md` 在基线时为未跟踪文件且已被后续修改，无 git 历史可回退。清单与基线指纹见 [baseline/README.md](baseline/README.md)，**未做任何重建**。
- 仓库整体仍是"零提交基线"（HEAD 之外的前序工作全部未提交），其他未跟踪文件处于同样风险中；是否先提交一次属流程决策，本阶段未执行。
