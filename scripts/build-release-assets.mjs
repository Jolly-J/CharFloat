// 构建独立发布副本；不修改任何业务源码、开发加载项或技能原稿。
import fs from 'node:fs';
import path from 'node:path';
import { ROOT, STAGE, POLICY, PROFILE, MANIFEST, ADDONS, sha256, protectAddon, walk, verifyStage } from './lib/release-protection.mjs';

const inputs = [...POLICY.files, 'package.json', 'scripts/release-files.json',
  'scripts/build-release-assets.mjs', 'scripts/lib/release-protection.mjs', 'package-lock.json'];
const assets = walk(path.join(ROOT, 'dist/renderer/assets')).map(n => `dist/renderer/assets/${n}`);
for (const name of assets) {
  if (!new RegExp(POLICY.rendererAssetPattern).test(name)) throw new Error(`意外的界面构建资源: ${name}`);
}
// ⚠️ **先清空副本再重建**：`beforePack` 会按平台往副本里增删文件
// （Windows 加明文兜底 `cli-full.cjs`、macOS 移除它），
// 若不清理，上一轮 Windows 构建留下的兜底文件会让本次 `verifyStage()` 直接失败——
// 表现为"改了源码却构建报禁止发布"，很容易误以为是源码问题。
fs.rmSync(STAGE, { recursive: true, force: true });

const names = [...POLICY.files, ...assets, 'package.json'];
for (const name of names) if (!fs.existsSync(path.join(ROOT, name))) throw new Error(`缺少构建输入: ${name}；请先 npm run build`);
// 仅清理本脚本专用的可重建目录，不清理开发 dist/ 或已有 release/。
fs.rmSync(STAGE, { recursive: true, force: true });
const manifest = { version: JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'))).version,
  profile: PROFILE, inputs: {}, files: {} };
for (const name of names) {
  const original = fs.readFileSync(path.join(ROOT, name));
  const content = ADDONS.includes(name) ? Buffer.from(protectAddon(original.toString(), name)) : original;
  const dest = path.join(STAGE, name);
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.writeFileSync(dest, content);
  if (name !== 'package.json') manifest.files[name] = sha256(content);
}
for (const name of new Set([...inputs, ...assets])) manifest.inputs[name] = sha256(fs.readFileSync(path.join(ROOT, name)));
fs.writeFileSync(path.join(STAGE, MANIFEST), JSON.stringify(manifest, null, 2) + '\n');
verifyStage();
console.log(`发布副本已生成：${path.relative(ROOT, STAGE)}，${names.length} 个文件；技能和提示词逐字保留。`);
