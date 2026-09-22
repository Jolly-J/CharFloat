# P1 证据文件说明

本目录保存**修复前**的失败证据。修复后测试会通过，因此这些输出是唯一能说明缺陷曾经存在、且确实被本次修复消除的依据。

所有复现都读取 [baseline/BASELINE.json](../baseline/BASELINE.json) 里的**固定锚点提交**（`cac6d38…`），不使用会随提交前进的 `HEAD`。

---

## `p1.4-prefix-failure.txt`

**用途**：证明 P1.3 修复的"跨通道重复写入"缺陷在修复前真实存在。

**产生方式**（在**临时副本**中执行，不覆盖当前工作区）：

```bash
bash docs/acceptance/2.1.0-p0p1/evidence/reproduce-prefix-failure.sh
```

脚本做三件事：

1. 把当前工作区复制到系统临时目录（排除 `node_modules`/`dist`/`release`/`.git`，软链当前 `node_modules`）；
2. 从锚点提交取 `src/bridge/office/adapter.ts` 覆盖**副本**中的同名文件——该文件在本轮之前与提交一致，因此锚点版本就是修复前实现；
3. 在副本内运行 `tests/failure-routing.test.ts`。

**为什么不能在原地做**：早先的复现方式直接覆盖工作区的 `adapter.ts` 再按 SHA-256 还原。那依赖"还原步骤不出错"，一旦中断就会把修复后的实现留在被破坏状态。现在改为副本执行后，工作区的 `src/bridge/office/adapter.ts` **全程未被触碰**（复现后实测指纹仍为 `30e7988f…`）。

**结果**：`tests 11 / pass 9 / fail 2`（与 `p1.4-prefix-failure.txt` 末尾的"结果签名"一致）。

**判定方式（不看退出码本身）**：脚本会解析实际结果签名并与 `EXPECTED_TESTS/PASS/FAIL` 及两个预期失败的用例名逐一比对，只有完全一致才判定复现成功：

| 脚本退出码 | 含义 |
|---|---|
| **0** | 复现成功：结果签名与预期一致，缺陷确实存在 |
| 1 | 未复现：签名不符（缺陷已消失，或测试集已变化需同步 `EXPECTED_*`） |
| 2 | 环境/锚点问题，无法判定 |

**不能只凭"退出码为 1"判断复现成功**：测试进程本身的退出码为 1 只说明"有测试失败"，也可能是环境问题或测试集漂移。必须以结果签名与失败用例名为准。

**失败断言的含义**：不是"报错文本不符"，而是 `AssertionError: 必须抛出错误而不是静默成功`（`actual: null`）。也就是说修复前 `callOffice` **成功返回**了——因为原生 COM 通道把同一次写入又执行了一遍，第二次的成功掩盖了第一次的失败。这正是"重复写入"缺陷的可观察后果。

**注意**：文件是原始输出（含脚本自身的锚点回显），未做编辑。扩展名用 `.txt` 而非 `.log`，因为 `.gitignore` 忽略 `*.log`，用 `.log` 会导致证据无法纳入版本跟踪（`git check-ignore` 已确认 `.txt` 不被忽略）。

---

## `baseline-head-anchor-prefix-failure.txt`

**用途**：证明"恢复逻辑依赖 HEAD"这个缺陷真实存在，且修复必要。

**背景**：评审发现 `verify-baseline.mjs` 原先用 `git archive HEAD`，而补丁对应固定提交 `cac6d38`。一旦把当前基线内容提交，导出树就变了，补丁无法应用。

**产生方式**（在临时副本仓库中制造"HEAD 已前进"，再按旧逻辑操作）：

```bash
root=$(git rev-parse --show-toplevel)
base="$root/docs/acceptance/2.1.0-p0p1/baseline"
tmp=$(mktemp -d)
cp -R "$root/.git" "$tmp/.git"
rsync -a --exclude node_modules --exclude dist --exclude release --exclude website/dist --exclude .git "$root/" "$tmp/"
git -C "$tmp" add -A
git -C "$tmp" -c user.email=e@l -c user.name=e commit -qm "simulate: baseline content committed"
mkdir -p "$tmp/tree"
git -C "$tmp" archive HEAD | tar -x -C "$tmp/tree"     # 旧实现用的是 HEAD
git -C "$tmp/tree" init -q
git -C "$tmp/tree" apply --whitespace=nowarn "$base/baseline-prior.patch"
```

**结果**：退出码 1，`patch does not apply` 与 `already exists in working directory` 交替出现（`README.md`、`package.json`、`src/bridge/gateway.ts`、`wps-addon/addon-core.js` 等大面积冲突）。

**修复后**：恢复逻辑改读 `BASELINE.json` 的 `baselineCommit`，并用隔离脚本验证：

```bash
bash docs/acceptance/2.1.0-p0p1/baseline/verify-after-head-advance.sh
# 副本 HEAD 前进为新提交 → 仍 22/22 通过，退出码 0
```

---

## 已知限制

- 两份证据都只覆盖各自的单一缺陷，不能推断其他风险已消除。
- 复现依赖 `rsync`、`tar`、`node`、`git`；Windows 原生命令行需改用等价命令。
- 锚点提交必须在仓库历史中可达。若将来做历史重写或让该提交被 `gc` 清理，两者都会失效；`verify-baseline.mjs` 会检测并明确报错。
