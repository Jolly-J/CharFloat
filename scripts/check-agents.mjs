import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

// Read-only; exported for isolated fixture tests. Never traverse build/dependency trees.
export function checkAgents(root) {
  const errors = [], files = [];
  const skip = new Set(['node_modules', 'dist', 'release', 'out', '.git']);
  function walk(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      if (entry.isSymbolicLink() || skip.has(entry.name) || entry.name.startsWith('.')) continue;
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) walk(full);
      else if (/^agents\.md$/i.test(entry.name)) {
        if (entry.name !== 'AGENTS.md') errors.push(`${path.relative(root, full)}: 文件名应为 AGENTS.md`);
        files.push(full);
      }
    }
  }
  walk(root);
  const rootGuide = path.join(root, 'AGENTS.md');
  if (!files.includes(rootGuide)) errors.push('缺少根 AGENTS.md');
  const linked = new Set();
  for (const file of files) {
    const text = fs.readFileSync(file, 'utf8');
    const label = path.relative(root, file);
    for (const match of text.matchAll(/\[[^\]]*\]\(([^)]+)\)/g)) {
      let target = match[1].trim().replace(/^<|>$/g, '');
      if (/^[a-z]+:/i.test(target) || target.startsWith('#')) continue;
      target = decodeURIComponent(target.split('#')[0]);
      const full = path.resolve(path.dirname(file), target);
      if (!fs.existsSync(full)) errors.push(`${label}: 无效引用 ${target}`);
      if (file === rootGuide) linked.add(full);
    }
    if (file !== rootGuide && !text.includes('## 同步维护')) errors.push(`${label}: 缺少同步维护约定`);
    if (file !== rootGuide && !text.includes('## 避坑')) errors.push(`${label}: 缺少避坑入口`);
  }
  for (const file of files) {
    if (file !== rootGuide && !linked.has(file)) errors.push(`${path.relative(root, file)}: 未纳入根导航`);
  }
  const pkgPath = path.join(root, 'package.json');
  if (fs.existsSync(pkgPath)) {
    const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
    for (const file of files) for (const match of fs.readFileSync(file, 'utf8').matchAll(/npm run ([\w:-]+)/g)) {
      if (!pkg.scripts?.[match[1]]) errors.push(`${path.relative(root, file)}: 根 package.json 缺少命令 ${match[1]}`);
    }
    if (!pkg.build?.files?.includes('!**/AGENTS.md')) errors.push('打包配置缺少 !**/AGENTS.md 排除项');
  }
  return { errors, count: files.length };
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const result = checkAgents(root);
  if (result.errors.length) {
    console.error(result.errors.join('\n')); process.exitCode = 1;
  } else console.log(`协作导航检查通过：${result.count} 个 AGENTS.md；路径、根导航、维护入口和 npm 命令有效。职责准确性仍需人工核对。`);
}
