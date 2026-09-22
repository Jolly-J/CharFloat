#!/usr/bin/env bash
#
# 在**临时副本**里复现 P1.3 修复前的失败，不覆盖当前工作区。
#
# 背景：`tests/failure-routing.test.ts` 的故障注入测试在修复后会通过，
# 所以必须用"修复前的 adapter + 当前测试文件"跑一遍，才能证明缺陷曾经存在。
#
# 做法：
#   1. 把当前工作区复制到系统临时目录（排除 node_modules/dist/release/.git）；
#   2. 从 BASELINE.json 的**固定提交**取出 `src/bridge/office/adapter.ts` 覆盖副本里的同名文件；
#      —— 该文件在本轮之前与提交一致，因此固定提交里的版本就是修复前的实现；
#   3. 软链当前 node_modules，在副本内运行测试。
#
# 判定：**不只看退出码**。脚本核对实际结果签名（用例数、通过数、失败数、失败用例名），
#      只有签名与预期一致才判定复现成功。
#
# 全程不改动工作区。锚点提交见 BASELINE.json，不使用会前进的 HEAD。
#
# 用法：bash docs/acceptance/2.1.0-p0p1/evidence/reproduce-prefix-failure.sh
# 退出码：0 = 复现成功（缺陷确实存在，签名一致）
#         1 = 未复现（签名不符：缺陷已消失，或测试集已变化）
#         2 = 环境/锚点问题，无法判定

set -uo pipefail

EXPECTED_TESTS=11
EXPECTED_PASS=9
EXPECTED_FAIL=2
EXPECTED_FAILING_NAMES=(
  '故障注入：写入已发生但响应超时，不得在原生通道重复执行'
  '故障注入：写入已完成但响应无法解析，同样不得回退重放'
)

here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
baseline_dir="$(cd "$here/../baseline" && pwd)"
repo_root="$(git -C "$here" rev-parse --show-toplevel)"

pinned="$(node -e "process.stdout.write(require('$baseline_dir/BASELINE.json').baselineCommit)")"
echo "锚点提交: $pinned"
echo "仓库根目录: $repo_root"

if ! git -C "$repo_root" cat-file -e "${pinned}^{commit}" 2>/dev/null; then
  echo "✖ 锚点提交 $pinned 不存在，无法复现" >&2
  exit 2
fi

tmp="$(mktemp -d)"
output="$(mktemp)"
cleanup() { rm -rf "$tmp" "$output"; }
trap cleanup EXIT

# 1. 复制工作区（不含依赖与产物）
rsync -a \
  --exclude node_modules --exclude dist --exclude release --exclude website/dist \
  --exclude .git "$repo_root/" "$tmp/"
ln -s "$repo_root/node_modules" "$tmp/node_modules"

# 2. 用固定提交里的实现覆盖副本中的 adapter（修复前版本）
git -C "$repo_root" show "${pinned}:src/bridge/office/adapter.ts" > "$tmp/src/bridge/office/adapter.ts"
echo "已把副本中的 adapter 替换为锚点提交版本（修复前实现）"

# 3. 在副本内运行，原始输出同时落盘
cd "$tmp"
node --import tsx --test tests/failure-routing.test.ts 2>&1 | tee "$output"
test_status="${PIPESTATUS[0]}"

# 4. 核对结果签名，而不是只看退出码
tests_count="$(sed -n 's/^ℹ tests \([0-9]*\)$/\1/p' "$output" | tail -1)"
pass_count="$(sed -n 's/^ℹ pass \([0-9]*\)$/\1/p' "$output" | tail -1)"
fail_count="$(sed -n 's/^ℹ fail \([0-9]*\)$/\1/p' "$output" | tail -1)"

echo
echo "---- 复现判定 ----"
echo "测试进程退出码: ${test_status}（修复前预期为 1）"
echo "结果签名: tests=${tests_count:-未取到} pass=${pass_count:-未取到} fail=${fail_count:-未取到}（预期 ${EXPECTED_TESTS}/${EXPECTED_PASS}/${EXPECTED_FAIL}）"

verdict=0
if [ -z "$tests_count" ] || [ -z "$pass_count" ] || [ -z "$fail_count" ]; then
  echo "✖ 未取到结果签名，无法判定（测试可能未能启动）" >&2
  verdict=2
elif [ "$tests_count" != "$EXPECTED_TESTS" ] || [ "$pass_count" != "$EXPECTED_PASS" ] || [ "$fail_count" != "$EXPECTED_FAIL" ]; then
  echo "✖ 结果签名不符：缺陷未复现，或测试集已变化（请先同步 EXPECTED_* 常量）" >&2
  verdict=1
else
  for name in "${EXPECTED_FAILING_NAMES[@]}"; do
    if ! grep -qF "✖ ${name}" "$output"; then
      echo "✖ 预期失败的用例未出现在失败列表: ${name}" >&2
      verdict=1
    fi
  done
  if [ "$verdict" -eq 0 ]; then
    echo "✔ REPRODUCED：修复前缺陷确实存在，${EXPECTED_FAIL} 项故障注入失败，失败用例名与预期一致"
    echo "   失败含义：callOffice 曾“成功返回”——原生通道把同一次写入又执行了一遍，第二次的成功掩盖了第一次的失败。"
  fi
fi

exit "$verdict"
