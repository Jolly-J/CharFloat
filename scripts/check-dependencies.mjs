// Read-only: compare the actual dependency tree to package-lock.json. Never installs or changes it.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('../', import.meta.url));
export function checkDependencies(directory = path.join(root, 'node_modules'), project = root) {
  const lock = JSON.parse(fs.readFileSync(path.join(project, 'package-lock.json'), 'utf8'));
  const pkg = JSON.parse(fs.readFileSync(path.join(project, 'package.json'), 'utf8'));
  const errors = []; let checked = 0;
  for (const kind of ['dependencies', 'devDependencies']) {
    if (JSON.stringify(pkg[kind]) !== JSON.stringify(lock.packages[''][kind])) errors.push(`package.json 与锁文件的 ${kind} 不一致`);
  }
  for (const [name, item] of Object.entries(lock.packages)) {
    if (!name || item.link) continue;
    const file = path.join(directory, name.replace(/^node_modules\//, ''), 'package.json');
    if (!fs.existsSync(file)) { if (!item.optional) errors.push(`缺少 ${name}`); continue; }
    const actual = JSON.parse(fs.readFileSync(file, 'utf8')).version; checked++;
    if (actual !== item.version) errors.push(`${name}: 锁定 ${item.version}，实际 ${actual}`);
  }
  return { checked, errors };
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const result = checkDependencies(process.argv[2]);
  console.log(JSON.stringify(result, null, 2)); if (result.errors.length) process.exitCode = 1;
}
