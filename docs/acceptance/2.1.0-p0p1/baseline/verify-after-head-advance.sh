#!/usr/bin/env bash
#
# 隔离验证：模拟"当前基线内容已被提交、HEAD 已经前进"之后，恢复是否仍然成立。
#
# 为什么需要：补丁对应的是固定提交 cac6d38。如果恢复逻辑用 HEAD 导出工作树，
# 一旦产生新提交，导出内容就变了，补丁会无法应用。本脚本在**临时副本**里制造
# "HEAD 已前进"的局面并重跑验证器，证明恢复能力不随提交前进而失效。
#
# 安全说明：所有提交动作都发生在系统临时目录里的副本仓库；
#          **不会**对当前仓库 commit / push / 改 ref，也不改动工作区。
#
# 用法：bash docs/acceptance/2.1.0-p0p1/baseline/verify-after-head-advance.sh
# 期望：退出码 0，且验证器输出 22/22 通过、并标明"已前进，恢复仍按锚点提交进行"。

set -uo pipefail

here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
repo_root="$(git -C "$here" rev-parse --show-toplevel)"
pinned="$(node -e "process.stdout.write(require('$here/BASELINE.json').baselineCommit)")"

tmp="$(mktemp -d)"
cleanup() { rm -rf "$tmp"; }
trap cleanup EXIT

echo "真实仓库: $repo_root"
echo "锚点提交: $pinned"

# 1. 带历史复制到临时目录，使固定锚点在副本里可达
cp -R "$repo_root/.git" "$tmp/.git"
rsync -a \
  --exclude node_modules --exclude dist --exclude release --exclude website/dist \
  --exclude .git "$repo_root/" "$tmp/"

# 2. 在副本里提交当前（已修改）内容 —— 相当于"基线内容落地成提交"
git -C "$tmp" add -A
git -C "$tmp" -c user.email=baseline-check@local -c user.name=baseline-check \
  commit -qm "simulate: baseline content committed"

new_head="$(git -C "$tmp" rev-parse HEAD)"
echo "副本新 HEAD: $new_head"
if [ "$new_head" = "$pinned" ]; then
  echo "✖ 副本 HEAD 未前进，隔离条件不成立" >&2
  exit 2
fi
echo "✔ 副本 HEAD 已前进（与锚点不同），开始恢复验证"

# 3. 对着"HEAD 已前进"的副本跑验证器
node "$here/verify-baseline.mjs" --repo "$tmp"
exit $?
