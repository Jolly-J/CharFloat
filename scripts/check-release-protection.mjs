import fs from 'node:fs';
import path from 'node:path';
import { ROOT } from './lib/release-protection.mjs';
import { checkRelease } from './lib/check-release.mjs';

const args = process.argv.slice(2);
const version = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'))).version;
const targets = args.length ? args : ['mac', 'win'].map(os => path.join(ROOT, 'release', version, os))
  .filter(dir => fs.existsSync(dir)).flatMap(dir => fs.readdirSync(dir).filter(n => n.endsWith('.zip')).map(n => path.join(dir, n)));
if (!targets.length) throw new Error('未找到 ZIP 发布包；没有检查任何包，不能判定通过。请显式提供 ZIP/app.asar/.app 路径。');
for (const target of targets) console.log(JSON.stringify(await checkRelease(target), null, 2));
