import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

// 只读品牌检查。历史验收与既有发布包保留原貌；不依赖 Git，也不扫描工作区父路径。
export function checkBrand(root, mode = 'source') {
  const errors = [];
  const pattern = /office(?:[\s_-]|%20|%2d|%5f)*agent(?:[\s_-]|%20|%2d|%5f)*bridge/i;
  const skipped = new Set(['.git', 'node_modules', 'dist', 'release', 'out', '.scratch', '.cache', '.venv', '__pycache__']);
  const generated = new Set(['wps-addon/addon-core.js', 'office-addon/public/taskpane.js']);
  let count = 0;
  function inspect(full) {
    const relative = path.relative(root, full).split(path.sep).join('/');
    if (pattern.test(relative)) errors.push(`${relative}: 路径包含旧品牌标识`);
    const data = fs.readFileSync(full);
    // UTF-16 清单/脚本也检查；其他二进制不作为文本解读。
    const text = data[0] === 0xff && data[1] === 0xfe ? data.subarray(2).toString('utf16le')
      : data.includes(0) ? null : data.toString('utf8');
    if (text === null) return;
    count++;
    const lines = text.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      if (pattern.test(lines[i])) errors.push(`${relative}:${i + 1}: 包含旧品牌标识`);
    }
  }
  function walk(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      if (entry.isSymbolicLink()) continue;
      const full = path.join(dir, entry.name);
      const relative = path.relative(root, full).split(path.sep).join('/');
      if (mode === 'source' && (skipped.has(entry.name) || relative === 'docs/acceptance' || generated.has(relative))) continue;
      if (entry.isDirectory()) {
        if (pattern.test(relative)) errors.push(`${relative}: 目录包含旧品牌标识`);
        walk(full);
      } else inspect(full);
    }
  }
  if (mode === 'source') walk(root);
  else {
    const targets = mode === 'generated'
      ? ['dist/main', 'dist/preload', 'dist/bridge', 'dist/renderer', ...generated]
      : mode === 'website-output' ? ['website/dist']
      : mode === 'stage' ? ['dist/release-app'] : null;
    if (!targets) throw new Error(`未知检查范围: ${mode}`);
    for (const target of targets) {
      const full = path.join(root, target);
      if (!fs.existsSync(full)) errors.push(`${target}: 缺少构建产物`);
      else if (fs.statSync(full).isDirectory()) walk(full);
      else inspect(full);
    }
  }
  return { errors, count };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const mode = process.argv[2]?.replace(/^--/, '') || 'source';
  const result = checkBrand(root, mode);
  if (result.errors.length) {
    console.error(result.errors.join('\n'));
    process.exitCode = 1;
  } else console.log(`字浮品牌检查通过：${mode}，${result.count} 个文本文件，旧品牌标识零命中。`);
}
