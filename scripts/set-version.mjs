// Update declared source versions only. No build, key initialization, Git write or publication.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { checkVersion } from './check-version.mjs';
export function setVersion(root, version, date = new Date().toISOString().slice(0, 10)) {
  const current = checkVersion(root);
  if (current.errors.length) throw new Error(current.errors.join('\n'));
  if (!/^(0|[1-9]\d{0,5})\.(0|[1-9]\d{0,5})\.(0|[1-9]\d{0,5})$/.test(version)) throw new Error('需要 MAJOR.MINOR.PATCH 版本号。');
  const a = version.split('.').map(Number), b = current.version.split('.').map(Number), index = a.findIndex((n, i) => n !== b[i]);
  if (index < 0 || a[index] <= b[index]) throw new Error('新版本必须高于当前版本；不得覆盖已发布版本。');
  const changes = new Map();
  const edit = (file, action) => { const old = fs.readFileSync(path.join(root, file), 'utf8'); changes.set(file, { old, next: action(old) }); };
  const replace = (text, pattern, value) => {
    const matches = [...text.matchAll(pattern)]; if (matches.length !== 1) throw new Error('版本声明匹配不唯一，未写入任何文件。');
    return text.replace(pattern, value);
  };
  edit('package.json', text => { const p = JSON.parse(text); p.version = version; return JSON.stringify(p, null, 2) + '\n'; });
  edit('package-lock.json', text => { const p = JSON.parse(text); p.version = p.packages[''].version = version; return JSON.stringify(p, null, 2) + '\n'; });
  for (const [file, pattern, replacement] of [
    ['src/bridge/runtime.ts', /export const VERSION = '[^']+'/g, `export const VERSION = '${version}'`],
    ['wps-addon/src/shared.js', /const ADDON_VERSION = ["'][^"']+["']/g, `const ADDON_VERSION = '${version}'`],
    ['office-addon/src/state.js', /const ADDON_VERSION = ["'][^"']+["']/g, `const ADDON_VERSION = '${version}'`],
    ['wps-addon/manifest.xml', /<version>[^<]+<\/version>/g, `<version>${version}</version>`],
    ['office-addon/excel/manifest.xml', /<Version>[^<]+<\/Version>/g, `<Version>${version}.0</Version>`],
    ['src/renderer/App.tsx', /version: '[^']+'/g, `version: '${version}'`],
    ['src/renderer/components/LiveWorkspaceCard.tsx', /installedVersion \|\| '[^']+'/g, `installedVersion || '${version}'`],
    ['src/renderer/components/SettingsDrawer.tsx', /appInfo\?\.version \|\| '[^']+'/g, `appInfo?.version || '${version}'`],
  ]) edit(file, text => replace(text, pattern, replacement));
  edit('CHANGELOG.md', text => {
    if (text.includes(`## [${version}]`)) throw new Error('变更日志已有此版本，未覆盖。');
    // Only the first active Unreleased section is released; historical sections remain intact.
    if (!text.includes('## [Unreleased]')) throw new Error('缺少 Unreleased 变更说明。');
    return text.replace('## [Unreleased]', `## [Unreleased]\n\n## [${version}] - ${date}`);
  });
  const written = [];
  try {
    for (const [file, { old, next }] of changes) {
      if (fs.readFileSync(path.join(root, file), 'utf8') !== old) throw new Error('写入前源码被其他操作修改，已停止。');
      fs.writeFileSync(path.join(root, file), next); written.push(file);
    }
    const result = checkVersion(root); if (result.errors.length) throw new Error(result.errors.join('\n'));
  } catch (error) {
    for (const file of written.reverse()) fs.writeFileSync(path.join(root, file), changes.get(file).old);
    throw error;
  }
  return { from: current.version, version, files: [...changes.keys()] };
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try { console.log(JSON.stringify(setVersion(fileURLToPath(new URL('../', import.meta.url)), process.argv[2]), null, 2)); }
  catch (error) { console.error(error.message); process.exitCode = 1; }
}
