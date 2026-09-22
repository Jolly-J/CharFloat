#!/usr/bin/env node
/**
 * P0 基线恢复验证器（只读仓库）。
 *
 * 作用：证明基线**可以被真正恢复**，而不只是"有一个指纹可以核对"。
 *   1. 读取 `BASELINE.json` 里**固定的基线提交**（不使用会前进的 HEAD）；
 *   2. 从该提交导出一份干净树到系统临时目录；
 *   3. 应用 `baseline-prior.patch`（该提交 → 前序未提交基线）；
 *   4. 按 `MANIFEST.sha256` 逐文件校验 SHA-256。
 *
 * 因为锚点是固定提交，**后续产生新提交不影响恢复能力**（可用 `--repo` 在副本上验证这一点）。
 *
 * 副作用：只写系统临时目录；不改动工作区、不建 git ref、不碰 node_modules。
 *
 * 用法：
 *   node docs/acceptance/2.1.0-p0p1/baseline/verify-baseline.mjs
 *   node docs/acceptance/2.1.0-p0p1/baseline/verify-baseline.mjs --repo /path/to/other/repo
 * 退出码：0 = 全部清单条目校验通过；1 = 锚点缺失、补丁失败或有哈希不符。
 */
import { execFileSync } from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));

// 解析 --repo；默认取本脚本所在仓库的根目录。全部路径从仓库根目录解析，不写死个人目录。
const repoFlagIndex = process.argv.indexOf('--repo');
const repoRoot = repoFlagIndex !== -1
  ? path.resolve(process.argv[repoFlagIndex + 1] ?? '')
  : execFileSync('git', ['rev-parse', '--show-toplevel'], { cwd: here, encoding: 'utf8' }).trim();
if (!fs.existsSync(path.join(repoRoot, '.git'))) throw new Error(`--repo 指向的不是 git 仓库根目录: ${repoRoot}`);

const meta = JSON.parse(fs.readFileSync(path.join(here, 'BASELINE.json'), 'utf8'));
const patchPath = path.join(here, meta.patch);
const manifestPath = path.join(here, meta.manifest);
const pinned = meta.baselineCommit;

// 锚点必须真实存在于该仓库，否则补丁无从应用。
try {
  execFileSync('git', ['cat-file', '-e', `${pinned}^{commit}`], { cwd: repoRoot, stdio: 'ignore' });
} catch {
  throw new Error(`基线锚点提交 ${pinned} 在 ${repoRoot} 中不存在，无法恢复；请确认仓库历史完整`);
}

const manifest = fs.readFileSync(manifestPath, 'utf8').split('\n')
  .map(line => line.trim()).filter(Boolean)
  .map(line => {
    const match = line.match(/^([0-9a-f]{64})\s+(.+)$/);
    if (!match) throw new Error(`MANIFEST 格式无法解析: ${line}`);
    return { sha256: match[1], file: match[2] };
  });
if (manifest.length !== meta.entryCount) throw new Error(`MANIFEST 条目数 ${manifest.length} 与 BASELINE.json 的 ${meta.entryCount} 不一致`);

const work = fs.mkdtempSync(path.join(os.tmpdir(), 'bridge-baseline-'));
const tarPath = path.join(work, 'baseline.tar');
try {
  // 1. 从固定提交导出干净树（不是 HEAD）
  fs.writeFileSync(tarPath, execFileSync('git', ['archive', pinned], { cwd: repoRoot, maxBuffer: 1 << 28 }));
  const tree = path.join(work, 'tree');
  fs.mkdirSync(tree);
  execFileSync('tar', ['-xf', tarPath, '-C', tree]);
  execFileSync('git', ['init', '-q'], { cwd: tree });

  // 2. 应用基线补丁（用绝对路径，避免 cwd 变化导致找不到补丁）
  execFileSync('git', ['apply', '--whitespace=nowarn', patchPath], { cwd: tree });

  // 3. 逐文件校验
  const failed = [];
  for (const entry of manifest) {
    const full = path.join(tree, entry.file);
    if (!fs.existsSync(full)) { failed.push(`${entry.file}: 恢复后不存在`); continue; }
    const actual = crypto.createHash('sha256').update(fs.readFileSync(full)).digest('hex');
    if (actual !== entry.sha256) failed.push(`${entry.file}: 期望 ${entry.sha256} 实际 ${actual}`);
  }

  const currentHead = execFileSync('git', ['rev-parse', 'HEAD'], { cwd: repoRoot, encoding: 'utf8' }).trim();
  const headMoved = currentHead !== pinned;
  console.log(`基线恢复验证：锚点提交 ${pinned}（固定）`);
  console.log(`  当前 HEAD ${currentHead}${headMoved ? '（已前进，恢复仍按锚点提交进行）' : '（与锚点一致）'}`);
  console.log(`  MANIFEST 条目 ${manifest.length}，通过 ${manifest.length - failed.length}，失败 ${failed.length}`);
  console.log(`  登记为不可恢复的文件 ${meta.unrecoverable.length} 个（不在清单内，见 README.md）`);
  if (failed.length) { for (const line of failed) console.error(`  ✖ ${line}`); process.exitCode = 1; }
  else console.log(`  ✔ 全部条目从锚点提交 + ${meta.patch} 恢复后 SHA-256 一致`);
} finally {
  fs.rmSync(work, { recursive: true, force: true });
}
