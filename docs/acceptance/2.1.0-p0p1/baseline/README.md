# P0 基线可恢复快照

本目录解决一个具体问题：**SHA-256 只能核验，不能恢复**。原 [baseline.md](../baseline.md) §1 只有指纹表，一旦工作区被覆盖就没有任何东西能还原基线；`git stash create` 产生的游离对象也既不可靠引用、也不包含未跟踪文件。因此这里持久保存**可获得的基线内容**，并提供可重复执行的恢复验证。

## 文件

| 文件 | 内容 |
|---|---|
| `BASELINE.json` | **固定锚点**：基线提交 SHA（`cac6d38…`）、补丁与清单文件名、条目数、不可恢复清单。恢复与失败复现都读取这里，**不使用会前进的 `HEAD`**。 |
| `baseline-prior.patch` | 从锚点提交到"前序未提交基线"的完整补丁，22 个文件（11 个 tracked 修改 + 11 个 untracked 新建）。可直接 `git apply`。 |
| `MANIFEST.sha256` | 22 个可恢复文件的标准 sha256sum 清单，可用 `shasum -a 256 -c` 校验。 |
| `verify-baseline.mjs` | 恢复验证器：从**锚点提交**导出干净树 → 应用补丁 → 逐文件校验哈希。只写系统临时目录。支持 `--repo <路径>`，可对副本仓库验证。 |
| `verify-after-head-advance.sh` | 隔离验证：在临时副本仓库里模拟"基线内容已提交、HEAD 已前进"，证明恢复能力不随提交前进而失效。 |

## 恢复方法

全部路径都从仓库根目录解析，**不写死个人目录**。

### 方式一：验证器（推荐，含自动校验）

```bash
root=$(git rev-parse --show-toplevel)
node "$root/docs/acceptance/2.1.0-p0p1/baseline/verify-baseline.mjs"
```

### 方式二：手工恢复

```bash
root=$(git rev-parse --show-toplevel)
base="$root/docs/acceptance/2.1.0-p0p1/baseline"
anchor=$(node -e "process.stdout.write(require('$base/BASELINE.json').baselineCommit)")

tmp=$(mktemp -d)
git -C "$root" archive "$anchor" | tar -x -C "$tmp"   # 用锚点提交，不是 HEAD
git -C "$tmp" init -q
git -C "$tmp" apply --whitespace=nowarn "$base/baseline-prior.patch"
(cd "$tmp" && shasum -a 256 -c "$base/MANIFEST.sha256")
```

> `git -C <目录> apply` 会把工作目录切走，因此补丁与清单都用**绝对路径**，否则会报找不到文件。

### 验证结果（实测）

| 场景 | 命令 | 结果 |
|---|---|---|
| 当前仓库 | `node verify-baseline.mjs` | 锚点 `cac6d38`，MANIFEST 22 条，**22/22 通过**，退出码 0 |
| **HEAD 已前进**（隔离副本） | `bash verify-after-head-advance.sh` | 副本 HEAD 前进为新提交（如 `91e40ad…`），**仍 22/22 通过**，退出码 0 |

### 为什么锚点是固定提交而不是 HEAD

补丁内容对应的是提交 `cac6d38` 的工作树。若用 `git archive HEAD` 导出：一旦把当前基线内容提交，导出的就是**已修改后**的树，补丁会大面积失败。修复前行为已复现并留证：[evidence/baseline-head-anchor-prefix-failure.txt](../evidence/baseline-head-anchor-prefix-failure.txt)（`patch does not apply` / `already exists in working directory`，退出码 1）。改成读取 `BASELINE.json` 的 `baselineCommit` 后，恢复与 HEAD 解耦。

失败复现同理：原先用 `git show HEAD:src/bridge/office/adapter.ts`，提交修复后取到的会是**修复后**的实现，复现失效。现改为读取同一锚点，并在临时副本中执行（[evidence/reproduce-prefix-failure.sh](../evidence/reproduce-prefix-failure.sh)），不覆盖当前工作区。

## 可恢复范围（22 个）

- **tracked 修改 11 个**：`README.md`、`package.json`、7 个 skill 文件、`src/bridge/gateway.ts`、`src/bridge/mcp-server.ts`、`wps-addon/addon-core.js`。
  其中 10 个在本轮未被再次修改，可直接取用；`package.json` 因 P0.2 新增了 `snapshot:tools` 一行已被改动，基线内容由**还原该行后重算 SHA-256 校验**得到（`4032ed0b…`，校验通过）。
- **untracked 新建 11 个**：`docs/refactoring-plan.md` 以外的 10 个模块 `AGENTS.md`/`scripts/check-agents.mjs`/`tests/ppt-layout.test.ts`/`skills/.../native-scripting.md` 等。
  完整名单以 `MANIFEST.sha256` 为准。

## 无法恢复的部分（如实登记，未重造）

以下 5 个文件在基线时是**未跟踪文件**，随后在 P0/P1 补正过程中被修改，因此：

- 不在任何 git 提交里（无历史可回退）；
- 当前工作区内容已不是基线内容；
- 游离 stash 对象不包含未跟踪文件，不能作为恢复来源。

**因此它们不在 `MANIFEST.sha256` 与补丁内，本目录不提供其基线内容，也不做"看起来像"的重建。** 同一份清单也结构化记录在 `BASELINE.json` 的 `unrecoverable` 字段中。

| 文件 | 基线 SHA-256 | 后续被谁改动 |
|---|---|---|
| `AGENTS.md` | `f25c93d6d8c8b43073f5bc0877f17c0623dd3e8f337e6dcdc3972c8e6a4405b2` | P0 补记 `docs/acceptance/` 证据目录约定 |
| `docs/refactoring-plan.md` | `2d0d5087b06d0fdaf55f569ecca73281c2b1d886324abc3b957a35d0d8ae57d0` | P0/P1 勾选、台账与决策点更新 |
| `scripts/AGENTS.md` | `0f8915cda7a49371b7fce38797c0500d36f31888a902572320ef2feb15eb267c` | P0.2 新增 `snapshot-tools.ts` 条目 |
| `src/bridge/AGENTS.md` | `96e38b98161b764028b2df9bce9c1c089ddfc6eedec68b0f2b11936e7a4dcf40` | P1 新增 `errors.ts` 条目与避坑条目 |
| `tests/AGENTS.md` | `8031f4def86ee1b2790ac8f054e0a83716b05800551ff2a85585bd490688cf07` | P1.4 新增 `failure-routing.test.ts` 条目 |

保留基线指纹仍有用：`git status` 会显示这 5 个文件为 untracked，若日后需要确认"某次改动是否动了它们"，可比对指纹判断当前内容是否仍等于基线。

**补救建议**（需用户决定，本阶段未执行）：把这 5 个文件先提交一次，之后所有改动即可由 git 历史恢复。当前仓库整体仍是"零提交基线"，HEAD 之外的前序工作全部未提交，其他未跟踪文件处于同样风险中——这属于流程决策。

## 边界说明

- 补丁只覆盖**基线时点**的前序修改，不包含 P0/P1 本轮新增的 `src/bridge/errors.ts`、`tests/failure-routing.test.ts`、`scripts/snapshot-tools.ts`、`.github/workflows/verify.yml` 改动、`src/bridge/{catalog,ws-server}.ts`、`src/bridge/office/adapter.ts`、`website/package-lock.json` 等。
- 恢复验证在系统临时目录进行，**不会**改动当前工作区；不要把它当作"回退已完成改动"的手段。
- 锚点是固定提交，因此**后续产生新提交不影响恢复**；前提是该提交仍在仓库历史中可达（未被 `gc` 清理、未做历史重写）。验证器会先检查锚点是否存在，缺失时明确报错而不是静默失败。
- `verify-after-head-advance.sh` 会在临时副本里执行 `git commit` 以制造"HEAD 已前进"的局面；**所有提交都发生在系统临时目录的副本中**，不会对当前仓库 commit / push / 改 ref。
