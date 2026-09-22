# 14 · 验收证据与文档缺口处理计划（ISS-06 / ISS-10 / ISS-16 / ISS-65）

开工时间：2026-09-22 · 执行者：验收证据与文档缺口处理员（子代理）
写区：仅 `docs/**`。不碰源码、`skills/**`、`tests/**`、`scripts/**`；不跑 `npm test` / `npm run build` / `npm run dist`；不做 git 操作。
ISS-06 比对脚本放 `.scratch/evidence/`，不放仓库根目录。

## 任务清单

| # | 任务 | 状态 | 产出 |
|---|---|---|---|
| 0 | 写本计划文件（防中断丢成果） | ✅ 完成 | 本文件 |
| 1 | 核对 `issues.md` 原始证据：ISS-06 / ISS-10 / ISS-16 / ISS-65 | ✅ 完成 | 本文件 §核对 |
| 2 | ISS-06-A：acorn 顶层语句多重集比对（目标 120/121） | ✅ 完成，**实测 120/121** | `evidence/iss-06-acorn-multiset.txt` + `.mjs` |
| 3 | ISS-06-B：diff（目标 0 行删除/修改、仅新增 62 行） | ✅ 完成，**实测 0/0/+62（60 注释+2 空行）** | `evidence/iss-06-diff-stat.txt` |
| 4 | ISS-06-C：桩测试 20 条 RPC 序列快照比对 | ✅ 完成，**实测 20/20 逐行一致**（另加 77/77 扩展） | `evidence/iss-06-rpc-sequence.txt` + `.mjs` |
| 5 | ISS-06-D：`issues.md` ISS-06 条目引用改为指向证据文件 | ✅ 完成 | `issues.md` |
| 6 | ISS-10：9 条未确认项逐条定状态，产出 `p5/open-questions.md` | ✅ 完成（6 条已澄清 / 3 条仍未确认） | `p5/open-questions.md` |
| 7 | ISS-16：补"实测场景 + 规避建议"到 `issues.md`（证据取自 `mcp-sweep/`） | ✅ 完成 | `issues.md` §ISS-16 |
| 8 | ISS-65：补台账记录（相对偏移语义 + 复现 + 规避） | ✅ 完成 | `issues.md` §ISS-65（新建小节） |
| 9 | 自检：所有结论有可复核文件；链接可达；无越界改动 | ✅ 完成 | 本文件 §自检 |

## 核对（任务 1）

`docs/acceptance/2.1.0-p0p1/issues.md`：

- **ISS-06**（第 36 行摘要 / 第 247–255 行正文）：声称 P3.2/P3.3 有 "acorn 顶层语句多重集 120/121 逐字节相同"、"diff 证明 0 行删除/修改、仅新增 62 行"、"桩测试 20 条 RPC 序列快照逐行一致"，但 `docs/acceptance/2.1.0-p0p1/evidence/` 下**无对应文件**（实测该目录只有 6 个 p1/p2 期前缀还原文件 + README）。结论：问题成立。
- **ISS-10**（第 40 行摘要 / 第 284–291 行正文）：指向 `tool-inventory.md` 第 180 行起的未确认项清单。结论：问题成立，需逐条定状态。
- **ISS-16**（第 46 行摘要 / 第 377–391 行正文）：正文已有现象与建议修法，**缺"实测场景 + 规避建议"**。结论：需补。
- **ISS-65**（第 95 行摘要）：**只有摘要表一行，无正文小节**。结论：需新建小节。

## 执行记录（时间 / 任务 / 结论 / 证据文件）

| 时间 | 任务 | 结论 | 证据文件 |
|---|---|---|---|
| — | 任务 1 核对 | 4 条问题全部成立（ISS-65 连正文小节都没有） | 本文件 §核对 |
| — | 还原"改造前"文件 | **关键发现**：`release/_legacy-build/` 与 `release/2.1.0/` 里各存着改造前/改造后的一份入口，且**指纹与台账记录对得上**，因此三项比对**完全不需要 git 操作、不需要构建** | 见下 |
| — | ISS-06-A | 改造前 5170 行（sha256 `6450c171…`，**与 `baseline/MANIFEST.sha256` 逐字符相同**，证明就是 P3.2 改造前的字节）+ `baseline-prior.patch` 该文件补丁段 → 用系统 `patch` 还原。**实测 120/121**，改造前独有恰好 1 条 `fitGeneratedPptShapes`，改造后独有 5 条（该函数 + 4 个新纯函数） | `evidence/iss-06-acorn-multiset.txt`、`.mjs` |
| — | ISS-06-B | **实测 0 行删除 / 0 行修改 / +62 行新增 = 注释 60 + 空行 2**，非注释非空的代码行新增 0 条。附带声明同步复核：77 case / 58 函数 / 61 sync / 40 load / 46 run，前后完全一致 | `evidence/iss-06-diff-stat.txt` |
| — | ISS-06-C | 自建 vm + 记录型桩（Office/Excel/DOM/WebSocket），**20/20 条序列逐行一致**、响应签名逐条一致；另 `--all` 跑全部 **77/77 一致** | `evidence/iss-06-rpc-sequence.txt`、`.mjs` |
| — | ISS-06 台账更新 | 新增「✅ ISS-06 修复记录」小节：三项对照表 + 输入指纹表 + 重跑命令 + 两点如实登记的口径；摘要行改 `已修已验证` | `issues.md` §ISS-06 |
| — | ISS-10 | 9 条逐条定状态 → **6 条已澄清 / 3 条仍未确认**（#4 超时阈值、#6 MS 审计快照、#9 审计 500 上限）；每条写清证据指向，未确认的写清可执行下一步 | `p5/open-questions.md`；`issues.md` §ISS-10 |
| — | ISS-16 | 补两个实测场景（① 4 个并行子代理，`shapes` 保存把兄弟任务第 9、10 张表落盘；② `e2e` 只建 1 张表却把 17 张共用的表一并落盘，且它**只能声明"未改动"，无法不写**）+ 源码根因（`wps-addon/src/excel.js:989-998` `wb.Save()`）+ **6 条规避建议** | `issues.md` §ISS-16 |
| — | ISS-65 | 新建小节：三例实测复现表（`G4:G15`→`row:1,col:1`；`I4:I6`→`row:2,col:1`；`B4:B15`→`row:5,col:1`）+ 源码根因（`wps-addon/src/excel.js:911,928-929`，`r/c` 是区域矩阵下标）+ 3 条调用方规避 + 建议修法 | `issues.md` §ISS-65 |
| — | 任务 9 自检 | 见下 §自检 | — |

### 关键发现：改造前/后文件全在磁盘上，无需 git

三项比对的输入**都是仓库内既有文件**，来源与指纹如下（这本是本次能"重跑"而不是"重写"的前提）：

| 角色 | 文件 | 行数 | sha256 |
|---|---|---|---|
| 改造前 addon-core.js（还原后） | `.scratch/evidence/prior/wps-addon/addon-core.js` | 5170 | `6450c171…` ★ 等于 `baseline/MANIFEST.sha256` 条目 |
| 改造前 addon-core.js（还原前） | `release/_legacy-build/win-unpacked/resources/app.asar.unpacked/wps-addon/addon-core.js` | 5109 | `a5ca8e9f…` |
| 改造后 addon-core.js | `release/2.1.0/win/win-unpacked/resources/app.asar.unpacked/wps-addon/addon-core.js` | 5247 | `c30e548a…` ★ 等于 `p5.1-candidate-and-prebuilt-state.md` 的"构建前指纹" |
| 改造前 taskpane.js | `release/_legacy-build/…/office-addon/public/taskpane.js` | 1809 | `ca86ef85…` |
| 改造后 taskpane.js | `release/2.1.0/…/office-addon/public/taskpane.js` | 1871 | `5ef08e87…` ★ 同上"构建前指纹" |

两个"改造前"文件的关系：`_legacy-build` 里的是 **P0 锚点提交 `cac6d38` 时的版本**；addon-core.js 在锚点→基线之间被 `baseline-prior.patch` 改过（PPT 布局修复），因此需应用补丁段才等于台账说的"5170 行"；taskpane.js 不在补丁的 22 个文件里，说明该区间未被改动，`_legacy-build` 那份即改造前内容。
**这个判断不是靠推理，而是被 diff 结果反向验证的**：取错文件绝不可能得到"0 行删除/修改、62 行新增且全为注释或空行"。

## 自检（任务 9）

| 检查项 | 结果 |
|---|---|
| 三项数字与台账声称对得上 | ✅ 120/121、0 删 0 改 +62、20/20 逐行一致 |
| 每个结论都有可复核文件 | ✅ 三份 `.txt` 全部含**原始 stdout/diff 逐字留档**；两个 `.mjs` 随证据留档可原样复跑 |
| "改造前"文件身份可证 | ✅ addon-core.js 还原后 sha256 **等于 `MANIFEST.sha256` 条目**；两个"改造后"sha256 **等于 `p5.1` 记录的构建前指纹** |
| 相对链接可达 | ✅ 6 个改动/新建文件，68 条相对链接，**0 条不可达** |
| 未跑 `npm test` / `build` / `dist` | ✅ 未跑（只用 `node` 直接跑自写脚本 + 系统 `diff` / `patch`） |
| 未做 git 操作 | ✅ 未做（补丁用系统 `patch` 应用；改造前文件取自 `release/` 既有副本） |
| 未碰源码 / `skills/**` / `tests/**` / `scripts/**` | ✅ 只读；改动全部落在 `docs/**` 与 `.scratch/evidence/**` |
| 脚本未放仓库根目录 | ✅ 均在 `docs/acceptance/2.1.0-p0p1/evidence/` 与 `.scratch/evidence/` |
| ISS-10 未确认项未被"写成已完成" | ✅ 3 条明确标 ❌ 仍未确认，各给可执行下一步；并写明"已澄清 ≠ 已解决" |
| ISS-16 / ISS-65 未越界改说明文案 | ✅ 只补台账；两处均显式注明"说明文案由工具说明批次负责"并指向 `09-tool-desc-fixes.md` |
| 未与并行任务冲突 | ⚠️ 编辑期间 `issues.md` 曾被**其他子代理并发修改**（编辑报 "file changed since it was read"），已改为「重新读取 → 单点小编辑 → 逐条验证」，全部编辑均已落盘成功 |

## 遗留与交接

1. **ISS-16 / ISS-65 的说明文案仍未改**：属工具说明批次（`09-tool-desc-fixes.md`）。本批次只把**可写进说明的事实**（实测场景、根因源码位置、规避建议、换算示例）备齐，写说明者可直接取用。
2. **ISS-10 的 3 条未确认项**依赖外部前置条件：#4 需宿主会话（部分可静态先做）、#6 需 **Microsoft Excel 通道连上**（唯一阻塞点）、#9 可在隔离脚本里构造 500+ 条记录验证。三者的可执行步骤已写在 `p5/open-questions.md`。
3. **P3.2/P3.3 的真实宿主等价性仍未验证**：本次给的是**文本层（diff）+ 行为层（桩测试）**的等价性；两个入口都没有在改造后做 WPS/Excel 实机加载对比（P3.8 亦自述"未做真实宿主回归"）。这一点已在 `issues.md` §ISS-06 的"仍未覆盖"里如实保留。
4. **`.scratch/evidence/` 是工作副本**，可能被后续清理；**留档的正式位置是 `docs/acceptance/2.1.0-p0p1/evidence/`**（含两个 `.mjs`）。若 `.scratch/evidence/prior/**` 被清理，按本文"关键发现"表里的"还原前"文件 + `baseline-prior.patch` 可原样重建。
