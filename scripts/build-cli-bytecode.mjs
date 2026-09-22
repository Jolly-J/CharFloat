// 把 dist/bridge/cli.cjs 编译成 V8 字节码 dist/bridge/cli.jsc，并生成极小的加载器。
//
// ## 为什么必须用 Electron 的 Node 编译
// V8 字节码与 V8 版本**锁死**：本机 node（V8 14.x）编译出来的，
// 在 Electron（V8 13.4）里**加载即失败**。必须用实际运行时的二进制来编译。
//
// ## 目录约定（开发可用 / 发布不泄源码）
//   dist/bridge/cli-full.cjs —— 真实实现（minify 后的明文 JS）。**不进包**。
//   dist/bridge/cli.jsc      —— 字节码（用 Electron 的 V8 编译）。进包。
//   dist/bridge/cli.cjs      —— 加载器：优先字节码；V8 不匹配时回退明文。进包。
//
// 回退存在的原因：**本地开发与测试用系统 node 启动**（V8 与 Electron 不同），
// 没有回退开发环境直接跑不起来。发布包里**没有** cli-full.cjs，
// 所以回退在客户机上不会生效、也不会因此泄漏源码。
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const root = process.cwd();
const bundle = path.join(root, 'dist/bridge/cli.cjs');
const full = path.join(root, 'dist/bridge/cli-full.cjs');
const jsc = path.join(root, 'dist/bridge/cli.jsc');

if (!fs.existsSync(bundle)) {
  console.error('✖ 找不到 dist/bridge/cli.cjs，请先跑 npm run build:main');
  process.exit(1);
}

// ① tsup 产出的明文包另存为 cli-full.cjs（开发/测试用；发布时排除）
fs.copyFileSync(bundle, full);

// ② 定位 Electron 二进制（打包后 cli 的实际运行环境）
const exe = process.platform === 'darwin'
  ? path.join(root, 'node_modules/electron/dist/Electron.app/Contents/MacOS/Electron')
  : process.platform === 'win32'
    ? path.join(root, 'node_modules/electron/dist/electron.exe')
    : path.join(root, 'node_modules/electron/dist/electron');
if (!fs.existsSync(exe)) { console.error(`✖ 找不到 Electron 二进制：${exe}`); process.exit(1); }
const bytenodeCli = path.join(root, 'node_modules/bytenode/lib/cli.js');
if (!fs.existsSync(bytenodeCli)) { console.error('✖ 缺少 bytenode'); process.exit(1); }

// 记下编译时的 V8 版本，写进加载器便于排查版本不符
const runtimeV8 = execFileSync(exe, ['-p', 'process.versions.v8'], {
  env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }, encoding: 'utf8',
}).trim();
const v8Tag = runtimeV8.split('.').slice(0, 2).join('.');

// ③ 用 Electron 的 Node 编译字节码
// ⚠️ bytenode 按**输入文件名**产出（cli-full.cjs → cli-full.jsc）。
// 必须先清掉旧字节码再编译，否则产物缺失时会被上一轮的陈旧文件蒙混过关。
const produced = path.join(root, 'dist/bridge/cli-full.jsc');
for (const f of [jsc, produced]) { try { fs.unlinkSync(f); } catch {} }
execFileSync(exe, [bytenodeCli, '-c', full], {
  env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }, stdio: 'inherit',
});
if (!fs.existsSync(produced)) { console.error('✖ 字节码未生成'); process.exit(1); }
fs.renameSync(produced, jsc);
// 编译完就没用了；留着只会让"哪个才是发布产物"变得含糊
try { fs.unlinkSync(full); } catch {}
fs.copyFileSync(bundle, full);  // 但 cli-full.cjs 要保留给开发/测试

// ④ 覆盖 tsup 的明文产出为加载器
fs.writeFileSync(bundle, `// 字节码加载器——无任何业务逻辑，可安全随包发布。
// 真实实现在 cli.jsc（编译时 V8 = ${runtimeV8}）。
// 发布包里**没有** cli-full.cjs，所以下面的回退只在本地开发/测试时生效。
const fs = require('fs'), path = require('path');
const jscFile = path.join(__dirname, 'cli.jsc');
const fullFile = path.join(__dirname, 'cli-full.cjs');
const need = '${v8Tag}';
const have = (process.versions.v8 || '').split('.').slice(0, 2).join('.');
if (have === need) {
  require('bytenode');
  require(jscFile);
} else if (fs.existsSync(fullFile)) {
  require(fullFile);
} else {
  throw new Error('字节码版本不匹配：cli.jsc 需要 V8 ' + need + '，当前为 ' + have +
    '。请用与目标 Electron 相同的版本执行 npm run build:bytecode。');
}
`);

const mb = (x) => (fs.statSync(x).size / 1048576).toFixed(2);
console.log(`[build:bytecode] cli.jsc ${mb(jsc)} MB（编译时 V8 = ${runtimeV8}）`);
console.log(`[build:bytecode] cli.cjs 加载器 ${fs.statSync(bundle).size} 字节`);
console.log(`[build:bytecode] cli-full.cjs ${mb(full)} MB —— 发布时必须排除`);
