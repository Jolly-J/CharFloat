#!/usr/bin/env node
/**
 * scripts/build-wps-addon.mjs
 *
 * 把 wps-addon/src/** 的源码模块按固定顺序拼成一个 IIFE，写出 WPS 加载项部署入口
 * wps-addon/addon-core.js。
 *
 * 约定（与 wps-addon/src/*.js 的文件头说明一致）：
 *   1. 每个 src 模块都是"构建片段"，内部语句保留 2 空格基础缩进，文本按原样拼接，
 *      拼接结果就是 (function () { ... })() 的函数体，因此片段之间共享同一个闭包作用域。
 *   2. 不引入任何 npm 依赖、不使用打包器、不做语法转换、不做 tree-shaking、不做压缩。
 *   3. 唯一的文本处理是移除 @build-strip:start / @build-strip:end 之间的内容块，
 *      该块仅用于让 ppt-layout.js 能被 Node 单元测试直接 import（见该文件末尾）。
 *   4. 输出完全由输入决定：同样的 src 内容必然得到逐字节相同的 addon-core.js。
 *
 * 用法：
 *   node scripts/build-wps-addon.mjs           生成 wps-addon/addon-core.js
 *   node scripts/build-wps-addon.mjs --check   只校验生成物是否与源码一致，不写文件
 */

import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";

const CHECK_ONLY = process.argv.includes("--check");

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SRC_DIR = path.join(ROOT, "wps-addon", "src");
const OUT_FILE = path.join(ROOT, "wps-addon", "addon-core.js");

/** 生成文件顶部注释（要求逐字固定） */
const HEADER = "// 本文件由 scripts/build-wps-addon.mjs 生成，请勿手改；改动请改 wps-addon/src/**";

/**
 * 构建指纹变量名。加载项在注册报文里上报它的值，桥接再与磁盘/部署副本比对，
 * 从而回答"WPS 进程里加载的到底是哪一版"（ISS-59：部署了新构建但没重载时，
 * 磁盘是新的、进程里跑的是旧的，此前没有任何指纹能识别出来）。
 */
const FINGERPRINT_VAR = "ADDON_BUILD_FINGERPRINT";

/**
 * 固定拼接顺序 = IIFE 内的语句顺序。
 * 顺序约束（改动前请先确认）：
 *   - shared.js 必须最前：它初始化 const config / let ws / let isConnected 等运行态变量，
 *     后续模块的函数虽然会被提升，但这些 const/let 必须在 bootstrap 执行前完成初始化。
 *   - bootstrap.js 必须最后：它执行 initWebSocket()、注册心跳、挂载 DOM 监听，是启动副作用。
 *   - ppt-layout.js 必须紧跟 ppt.js 之前，保证 pptPageSize/fitGeneratedPptShapes 在产物里相邻
 *     （tests/ppt-layout.test.ts 按这两个函数名截取源码片段做纯几何验证）。
 *   - sheet-sort.js 必须紧跟 excel.js 之前：它是排序写完后"读回校验"用的纯函数，
 *     由 tests/sheet-sort.test.ts 直接 import 验证（尤其是不得漏比较最后一行）。
 */

import * as esbuild from "esbuild";

const MODULES = [
  "shared.js",
  "connection.js",
  "ribbon.js",
  "dispatch.js",
  "sheet-sort.js",
  "excel.js",
  "word.js",
  "ppt-layout.js",
  "ppt.js",
  "bootstrap.js",
];

const STRIP_START = "@build-strip:start";
const STRIP_END = "@build-strip:end";

/** 移除 @build-strip:start ... @build-strip:end 之间的整段内容（含标记行） */
function stripBuildOnlyBlocks(text, file) {
  const lines = text.split("\n");
  const kept = [];
  let inside = false;
  let stripped = 0;
  for (const line of lines) {
    if (!inside && line.includes(STRIP_START)) {
      inside = true;
      continue;
    }
    if (inside) {
      if (line.includes(STRIP_END)) inside = false;
      stripped++;
      continue;
    }
    kept.push(line);
  }
  if (inside) {
    throw new Error(`${file}: 发现未闭合的 ${STRIP_START}（缺少 ${STRIP_END}）`);
  }
  return { text: kept.join("\n"), stripped };
}

function fail(message) {
  console.error(`[build:wps-addon] 失败：${message}`);
  process.exit(1);
}

if (!fs.existsSync(SRC_DIR)) {
  fail(`源码目录不存在：${path.relative(ROOT, SRC_DIR)}`);
}

/**
 * 校验 src 目录与 MODULES 清单**双向**一致，避免新增模块被静默漏拼。
 *
 * 只做"清单里的文件是否存在"是单向检查：新加一个 src/*.js 而忘记登记，构建会照常成功，
 * 该模块的代码却不会进入部署入口——功能静默缺失，且测试与构建都不会报错。
 */
function verifyModuleManifest() {
  const onDisk = fs.readdirSync(SRC_DIR, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.endsWith(".js"))
    .map((entry) => entry.name)
    .sort();
  const ordered = [...MODULES].sort();
  const duplicates = MODULES.filter((name, index) => MODULES.indexOf(name) !== index);
  const missing = ordered.filter((rel) => !onDisk.includes(rel));
  const unlisted = onDisk.filter((rel) => !MODULES.includes(rel));
  if (!duplicates.length && !missing.length && !unlisted.length) return;
  const lines = ["src 目录与 MODULES 清单不一致，构建已中止："];
  if (duplicates.length) lines.push(`  MODULES 中存在重复登记: ${[...new Set(duplicates)].join(", ")}`);
  if (missing.length) lines.push(`  MODULES 列出但文件不存在: ${missing.join(", ")}`);
  if (unlisted.length) lines.push(`  src 中存在未登记模块（其代码不会进入部署入口）: ${unlisted.join(", ")}`);
  lines.push("  请同步 scripts/build-wps-addon.mjs 的 MODULES。");
  fail(lines.join("\n"));
}

verifyModuleManifest();

const bodies = [];
const report = [];

for (const name of MODULES) {
  const file = path.join(SRC_DIR, name);
  if (!fs.existsSync(file)) {
    fail(`缺少源码模块：${path.relative(ROOT, file)}`);
  }
  const raw = fs.readFileSync(file, "utf8").replace(/\r\n/g, "\n");
  const { text, stripped } = stripBuildOnlyBlocks(raw, name);
  const body = text.replace(/\s+$/, "");
  if (body.length === 0) {
    fail(`源码模块内容为空：${path.relative(ROOT, file)}`);
  }
  // 拼接产物是普通脚本（非 ESM），残留的 import/export 会让 WPS 直接加载失败。
  const illegal = body
    .split("\n")
    .map((line, i) => ({ line, no: i + 1 }))
    .filter(({ line }) => /^\s*(import|export)\s/.test(line));
  if (illegal.length > 0) {
    fail(
      `${name} 第 ${illegal.map((x) => x.no).join(", ")} 行仍有 import/export；` +
        `请改用 @build-strip 包裹或删除`
    );
  }
  bodies.push(body);
  report.push({ name, lines: body.split("\n").length, strippedLines: stripped });
}

// 构建指纹：对"源码拼接体"取 sha256。它**只依赖源码**、不依赖注入结果，
// 因此可以安全地写进产物再让加载项上报；桥接拿它与磁盘构建比对，
// 就能回答"WPS 进程里加载的是哪一版"（ISS-59）。
//
// 指纹同时以两种形式落到产物里，缺一不可：
//   1. 头部注释 `// ADDON_BUILD_FINGERPRINT: <sha>` —— 桥接读**磁盘/已部署副本**时解析它；
//   2. IIFE 内 `var ADDON_BUILD_FINGERPRINT = "<sha>"` —— 加载项**运行时报**的上报值，
//      代表"WPS 进程里真正加载的字节"。
// 两者不一致即说明部署了新构建但进程里还是旧代码。
const sourceFingerprint = createHash("sha256").update(bodies.join("\n\n"), "utf8").digest("hex");
const fingerprintComment = `// ADDON_BUILD_FINGERPRINT: ${sourceFingerprint}`;
let output = `${HEADER}\n${fingerprintComment}\n(function () {\n  var ${FINGERPRINT_VAR} = "${sourceFingerprint}";\n${bodies.join("\n\n")}\n})();\n`;

// 产物自检：必须保持外层 IIFE 结构（WPS 通过 <script> 直接加载该文件），
// 且构建指纹注释与变量都要存在——桥接靠它们判断"进程里跑的是哪一版"。
if (
  !output.startsWith(`${HEADER}\n${fingerprintComment}\n(function () {\n`) ||
  !output.endsWith("\n})();\n") ||
  !output.includes(`var ${FINGERPRINT_VAR} = "${sourceFingerprint}";`)
) {
  fail("产物未形成合法的外层 IIFE 或缺少构建指纹，已中止写出");
}

// 产物自检：语法必须可编译。拆分/搬迁最容易出的错就是边界少一个大括号，
// 这里在写出前就拦下，避免把语法错误的加载项部署到 WPS。
try {
  new vm.Script(output, { filename: path.relative(ROOT, OUT_FILE) });
} catch (err) {
  fail(`产物语法检查未通过：${err.message}`);
}

// ── 压缩：去注释与空白、缩短局部变量名。
// **头部注释在压缩后重新注入**（必须留在文件最前面：桥接只读前 4096 字节匹配指纹）。
// 注意：压缩会 mangle 变量名，所以下面的自检只校验**指纹串**而不是变量名。
let minified = output;
try {
  const res = esbuild.transformSync(output, {
    minify: true,
    target: "es2018",          // 保守目标：WPS 的 JSA 与 Office 的 webview 都比现代浏览器保守
    legalComments: "none",
  });
  if (res.code && res.code.trim().length > 0) minified = res.code;
} catch (err) {
  fail(`压缩失败，已中止写出：${err.message}`);
}
minified = `${HEADER}\n${fingerprintComment}\n` + minified;
if (!minified.includes(sourceFingerprint)) fail("压缩后构建指纹丢失，已中止写出");
if (!minified.trimStart().startsWith(HEADER)) fail("压缩后产物未以头部注释开头，桥接将读不到指纹，已中止写出");
output = minified;

const previous = fs.existsSync(OUT_FILE) ? fs.readFileSync(OUT_FILE, "utf8") : null;
const changed = previous !== output;

if (CHECK_ONLY) {
  console.log(`[build:wps-addon] --check 模式：只校验不写文件`);
  if (previous === null) fail(`生成物不存在：${path.relative(ROOT, OUT_FILE)}`);
  if (changed) fail(`生成物与 wps-addon/src/** 不一致，请执行 npm run build:wps-addon 重新生成`);
  console.log(`[build:wps-addon] 生成物与源码一致：${path.relative(ROOT, OUT_FILE)}`);
  process.exit(0);
}

if (changed) {
  fs.writeFileSync(OUT_FILE, output, "utf8");
}

const totalLines = output.split("\n").length - 1;
console.log(`[build:wps-addon] 拼接顺序: ${MODULES.join(" -> ")}`);
for (const item of report) {
  const note = item.strippedLines > 0 ? `（另移除 ${item.strippedLines} 行 @build-strip 测试导出）` : "";
  console.log(`[build:wps-addon]   ${item.name.padEnd(16)} ${String(item.lines).padStart(5)} 行${note}`);
}
console.log(
  `[build:wps-addon] ${changed ? "已写出" : "产物未变化（逐字节一致，跳过写入）"} ` +
    `${path.relative(ROOT, OUT_FILE)}，共 ${totalLines} 行`
);
