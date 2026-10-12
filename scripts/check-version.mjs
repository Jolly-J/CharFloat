import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

// 只读门禁；package.json 是版本唯一基准，不检查历史证据或手改生成物。
export function checkVersion(root) {
  const errors = [];
  const read = name => fs.readFileSync(path.join(root, name), 'utf8');
  let version;
  try {
    version = JSON.parse(read('package.json')).version;
    if (typeof version !== 'string' || !/^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/.test(version)) {
      throw new Error('当前发布流程只接受无前导零的 MAJOR.MINOR.PATCH；预发布版本需先定义 Office 四段版本映射');
    }
  } catch (error) {
    return { version, errors: [`package.json: ${error.message}`] };
  }
  const compare = (name, actual, expected = version) => {
    if (actual !== expected) errors.push(`${name}: 期望 ${expected}，实际 ${String(actual)}`);
  };
  try {
    const lock = JSON.parse(read('package-lock.json'));
    compare('package-lock.json version', lock.version);
    compare('package-lock.json packages[""].version', lock.packages?.['']?.version);
  } catch (error) { errors.push(`package-lock.json: ${error.message}`); }

  const declarations = [
    ['src/bridge/runtime.ts', /export\s+const\s+VERSION\s*=\s*['"]([^'"]+)['"]/g],
    ['wps-addon/src/shared.js', /const\s+ADDON_VERSION\s*=\s*['"]([^'"]+)['"]/g],
    ['office-addon/src/state.js', /const\s+ADDON_VERSION\s*=\s*['"]([^'"]+)['"]/g],
    ['wps-addon/manifest.xml', /<version>\s*([^<]+?)\s*<\/version>/g],
    ['office-addon/excel/manifest.xml', /<Version>\s*([^<]+?)\s*<\/Version>/g, `${version}.0`],
    ['src/renderer/App.tsx', /useState[^\n]*\(\{\s*version:\s*['"]([^'"]+)['"]/g],
    ['src/renderer/components/LiveWorkspaceCard.tsx', /installedVersion\s*\|\|\s*['"]([^'"]+)['"]/g],
    ['src/renderer/components/SettingsDrawer.tsx', /appInfo\?\.version\s*\|\|\s*['"]([^'"]+)['"]/g],
  ];
  for (const [name, pattern, expected] of declarations) {
    try {
      const matches = [...read(name).matchAll(pattern)];
      if (matches.length !== 1) errors.push(`${name}: 应有且仅有一个版本声明，匹配到 ${matches.length} 个；声明结构变化时同步维护检查器`);
      else compare(name, matches[0][1], expected);
    } catch (error) { errors.push(`${name}: ${error.message}`); }
  }
  try {
    const headings = [...read('CHANGELOG.md').matchAll(/^## \[([^\]]+)\]/gm)];
    const count = headings.filter(match => match[1] === version).length;
    if (count !== 1) errors.push(`CHANGELOG.md: 当前版本 [${version}] 应有且仅有一个条目，实际 ${count} 个`);
  } catch (error) { errors.push(`CHANGELOG.md: ${error.message}`); }
  return { version, errors };
}

if (process.argv[1] && fs.realpathSync(path.resolve(process.argv[1])) === fs.realpathSync(fileURLToPath(import.meta.url))) {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const result = checkVersion(root);
  if (result.errors.length) {
    console.error(result.errors.join('\n'));
    process.exitCode = 1;
  } else console.log(`版本一致性检查通过：${result.version}（锁文件、Bridge、加载项、清单、界面及变更日志）。`);
}
