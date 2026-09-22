#!/usr/bin/env node
/**
 * 构建 Office.js 任务窗格部署入口（P3.3/P3.4）。
 *
 *   node scripts/build-office-addon.mjs            # 生成 office-addon/public/taskpane.js
 *   node scripts/build-office-addon.mjs --check    # 只校验生成物是否与源码一致，不写文件
 *
 * 输入：office-addon/src/**（仅此目录，按下方 MODULE_ORDER 固定顺序）
 * 输出：office-addon/public/taskpane.js（唯一部署入口；由 public/taskpane.html 加载、excel/manifest.xml 部署，路径不变）
 *
 * 约定与边界：
 * - src 下每个 .js 都是同一个 IIFE 内的“拼接片段”：不使用 import/export，不引入依赖，不引入打包器；
 *   片段之间依靠 IIFE 内的函数声明提升和共享作用域协作，因此拼接顺序只影响顶层语句的执行次序。
 * - 只做拼接、去重尾部空行和语法校验；不改写片段内容，不压缩、不转译、不加时间戳。
 * - 输出是 src 内容的纯函数，可重复执行且逐字节幂等。
 * - 禁止手改生成物：要改行为请改 office-addon/src/**，再重新执行本脚本。
 */
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const SCRIPT_PATH = fileURLToPath(import.meta.url);
const REPO_ROOT = path.resolve(path.dirname(SCRIPT_PATH), '..');
const SRC_DIR = path.join(REPO_ROOT, 'office-addon', 'src');
const OUT_FILE = path.join(REPO_ROOT, 'office-addon', 'public', 'taskpane.js');

const GENERATED_NOTICE =
  '// 本文件由 scripts/build-office-addon.mjs 生成，请勿手改；改动请改 office-addon/src/**';

// 迁移前 taskpane.js 的文件头注释，逐字保留。
const PRELUDE = [
  '/**',
  ' * WPS Bridge - Microsoft Office (Excel) 官方 Office.js 核心运行时',
  ' * 运行在 Microsoft Excel 任务窗格 WebView (WebKit / Edge WebView2)',
  ' */',
];

// 固定拼接顺序：与拆分前 taskpane.js 的语句顺序一致，不要随意调整。
const MODULE_ORDER = [
  'state.js',
  'ui.js',
  'lifecycle.js',
  'connection.js',
  'rpc.js',
  'excel/shared.js',
  'excel/workbook.js',
  'excel/sheets.js',
  'excel/range.js',
  'excel/formula.js',
  'excel/format.js',
  'excel/filter-sort.js',
  'excel/table.js',
  'excel/chart.js',
  'excel/pivot.js',
  'excel/shape.js',
  'excel/comment.js',
  'excel/script.js',
  'bootstrap.js',
];

function listModuleFiles(dir, prefix = '') {
  const found = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
    const rel = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isDirectory()) found.push(...listModuleFiles(path.join(dir, entry.name), rel));
    else if (entry.isFile() && entry.name.endsWith('.js')) found.push(rel);
  }
  return found;
}

/** 校验 src 目录与 MODULE_ORDER 严格一一对应，避免新增模块被静默漏拼。 */
function assertManifestMatchesSources() {
  if (!fs.existsSync(SRC_DIR)) throw new Error(`源码目录不存在: ${path.relative(REPO_ROOT, SRC_DIR)}`);
  const onDisk = listModuleFiles(SRC_DIR).sort();
  const ordered = [...MODULE_ORDER].sort();
  const missing = ordered.filter((rel) => !onDisk.includes(rel));
  const unlisted = onDisk.filter((rel) => !MODULE_ORDER.includes(rel));
  if (missing.length || unlisted.length) {
    const lines = ['src 目录与 MODULE_ORDER 不一致，构建已中止：'];
    if (missing.length) lines.push(`  MODULE_ORDER 列出但文件不存在: ${missing.join(', ')}`);
    if (unlisted.length) lines.push(`  src 下存在但未纳入拼接顺序: ${unlisted.join(', ')}`);
    lines.push('  请同步 scripts/build-office-addon.mjs 的 MODULE_ORDER。');
    throw new Error(lines.join('\n'));
  }
}

/** 读取片段：去掉尾部空行（由构建统一插入单行分隔），内容其余部分逐字保留。 */
function readModule(rel) {
  const full = path.join(SRC_DIR, rel);
  const source = fs.readFileSync(full, 'utf8').replace(/\r\n/g, '\n');
  if (!source.trim()) throw new Error(`模块内容为空: src/${rel}`);
  return source.replace(/\n+$/, '');
}

function buildSource() {
  assertManifestMatchesSources();

  const lines = [GENERATED_NOTICE];
  lines.push(`// 构建命令: node scripts/build-office-addon.mjs（按固定顺序拼接 src 下 ${MODULE_ORDER.length} 个片段，无打包器、无依赖）`);
  lines.push('// 拼接顺序:');
  MODULE_ORDER.forEach((rel, index) => {
    lines.push(`//   ${String(index + 1).padStart(2, '0')}. src/${rel}`);
  });
  lines.push('');
  lines.push(...PRELUDE);
  lines.push('');
  lines.push('(function () {');
  for (const rel of MODULE_ORDER) {
    lines.push('');
    lines.push(readModule(rel));
  }
  lines.push('');
  lines.push('})();');

  const output = lines.join('\n') + '\n';
  if (!output.startsWith(GENERATED_NOTICE + '\n')) throw new Error('生成物缺少顶部生成声明');
  new vm.Script(output, { filename: path.relative(REPO_ROOT, OUT_FILE) });
  return output;
}

const output = buildSource();
const previous = fs.existsSync(OUT_FILE) ? fs.readFileSync(OUT_FILE, 'utf8') : null;
const relOut = path.relative(REPO_ROOT, OUT_FILE);

if (process.argv.includes('--check')) {
  if (previous === output) {
    console.log(`生成物与源码一致: ${relOut}`);
  } else {
    console.error(`生成物与源码不一致，请执行 node scripts/build-office-addon.mjs: ${relOut}`);
    process.exitCode = 1;
  }
} else {
  fs.mkdirSync(path.dirname(OUT_FILE), { recursive: true });
  fs.writeFileSync(OUT_FILE, output, 'utf8');
  const action = previous === null ? '已生成' : previous === output ? '内容一致，已重写' : '已更新';
  console.log(`${action}: ${relOut}（${MODULE_ORDER.length} 个模块，${Buffer.byteLength(output, 'utf8')} 字节，语法校验通过）`);
}
