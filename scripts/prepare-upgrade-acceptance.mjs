// Copies public build inputs and applies a recorded isolation overlay only to new A/B source copies.
// Does not build, sign, install, deploy, commit or read the production private key.
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { assertPublicPaths } from './lib/update-key-safety.mjs';
import { setVersion } from './set-version.mjs';
import { checkDependencies } from './check-dependencies.mjs';
const root = fileURLToPath(new URL('../', import.meta.url));
const [planArg, dependenciesArg] = process.argv.slice(2);
if (!planArg || !dependenciesArg) throw new Error('需要现有验收 plan.json 和锁文件一致的 node_modules 路径');
const planFile = path.resolve(planArg), output = path.dirname(planFile), plan = JSON.parse(fs.readFileSync(planFile));
const versionNumber = value => /^\d+\.\d+\.\d+$/.test(value) ? value.split('.').reduce((n, v) => n * 1000000 + Number(v), 0) : NaN;
const feed = new URL(plan.feed);
if (!/^[a-zA-Z0-9-]{1,80}$/.test(plan.id) || !(versionNumber(plan.versions.B) > versionNumber(plan.versions.A)) || feed.protocol !== 'https:' || feed.origin !== plan.assetOrigin || feed.username || feed.password || feed.hash || feed.pathname === '/updates/stable.json' || !/^com\.[a-z0-9.-]+$/.test(plan.appId) || plan.appId === 'com.antigravity.charfloat' || !/^[a-zA-Z0-9-]{1,80}$/.test(plan.runtimeDirectoryName) || !/^[a-zA-Z0-9-]{1,80}$/.test(plan.windowsRootName) || plan.windowsRootName === 'CharFloat' || !/^[^\\/:*?"<>|]{1,80}\.lnk$/.test(plan.windowsShortcut) || plan.windowsShortcut === '字浮 CharFloat.lnk' || ![plan.bridgePort, plan.httpsPort].every(n => Number.isInteger(n) && n > 1024 && n < 65536 && ![19890,19891].includes(n)) || plan.bridgePort === plan.httpsPort || !/^[a-f0-9]{40}$/.test(plan.baseCommit)) throw new Error('验收计划缺少安全的独立身份、版本、目录、端口或 HTTPS 源');
const dependencies = fs.realpathSync(dependenciesArg);
const dependencyCheck = checkDependencies(dependencies);
if (dependencyCheck.errors.length) throw new Error(dependencyCheck.errors.join('\n'));
const directories = ['src','scripts','tests','resources','skills','prompts','build','wps-addon','office-addon','windows-bootstrap'];
const files = ['package.json','package-lock.json','pnpm-lock.yaml','pnpm-workspace.yaml','AGENTS.md','CHANGELOG.md','README.md','UPDATES.md','CLIENTS.md','LICENSE','.gitignore','.gitattributes','index.html','tsconfig.json','tsup.config.ts','vite.config.ts','tailwind.config.js','postcss.config.js'];
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const original = {}, overlays = {};
for (const stage of ['A','B']) {
  const destination = path.join(output, stage);
  if (fs.existsSync(destination)) throw new Error('验收源码目录已存在，未覆盖');
}
for (const stage of ['A','B']) {
  const destination = path.join(output, stage), version = plan.versions[stage]; fs.mkdirSync(destination);
  for (const dir of directories) fs.cpSync(path.join(root, dir), path.join(destination, dir), { recursive: true, filter: source => { assertPublicPaths([path.relative(root, source)]); return true; } });
  for (const file of files) fs.copyFileSync(path.join(root, file), path.join(destination, file));
  fs.symlinkSync(dependencies, path.join(destination, 'node_modules'), 'dir');
  execFileSync('git', ['init', '-q', destination], { stdio: 'ignore' });
  const edits = [];
  function edit(file, action) {
    const full = path.join(destination, file), before = fs.readFileSync(full, 'utf8'), after = action(before);
    if (after === before) throw new Error(`验收转换未生效: ${file}`);
    original[file] ||= sha(Buffer.from(before));
    fs.writeFileSync(full, after); edits.push({ file, before: sha(Buffer.from(before)), after: sha(Buffer.from(after)) });
  }
  function replaceOnce(text, before, after) {
    if (text.split(before).length !== 2) throw new Error('验收转换锚点不唯一');
    return text.replace(before, after);
  }
  const versionFiles = ['package.json','package-lock.json','src/bridge/runtime.ts','wps-addon/src/shared.js','office-addon/src/state.js','wps-addon/manifest.xml','office-addon/excel/manifest.xml','src/renderer/App.tsx','src/renderer/components/LiveWorkspaceCard.tsx','src/renderer/components/SettingsDrawer.tsx','CHANGELOG.md'];
  const beforeVersions = Object.fromEntries(versionFiles.map(file => [file, sha(fs.readFileSync(path.join(destination, file)))]));
  setVersion(destination, version);
  for (const file of versionFiles) { original[file] ||= beforeVersions[file]; edits.push({ file, before: beforeVersions[file], after: sha(fs.readFileSync(path.join(destination, file))) }); }
  edit('package.json', text => {
    const pkg = JSON.parse(text); pkg.version = version; pkg.build.appId = plan.appId; pkg.build.electronVersion = JSON.parse(fs.readFileSync(path.join(dependencies, 'electron/package.json'))).version;
    pkg.charfloatAcceptance = { ...plan, stage, version, runtimeBase: 'system temporary directory; canonicalized', production: false };
    return JSON.stringify(pkg, null, 2) + '\n';
  });
  edit('src/main/update-trust.ts', text => replaceOnce(text, "feeds: ['https://charfloat.online/updates/stable.json'], assetOrigins: ['https://charfloat.online']", `feeds: [${JSON.stringify(plan.feed)}], assetOrigins: [${JSON.stringify(plan.assetOrigin)}]`));
  edit('src/bridge/runtime.ts', text => {
    text = replaceOnce(text, 'export function runtimeHome() {', `export function runtimeHome() {\n  return path.join(fs.realpathSync(os.tmpdir()), ${JSON.stringify(plan.runtimeDirectoryName)});\n}\nfunction unusedProductionRuntimeHome() {`);
    text = replaceOnce(text, 'process.env.WPS_BRIDGE_PORT || 19890', String(plan.bridgePort));
    return replaceOnce(text, 'process.env.WPS_BRIDGE_HTTPS_PORT || 19891', String(plan.httpsPort));
  });
  edit('src/main/update-mac.ts', text => replaceOnce(text, "const BUNDLE_ID = 'com.antigravity.charfloat';", `const BUNDLE_ID = ${JSON.stringify(plan.appId)};`));
  edit('src/main/index.ts', text => {
    text = replaceOnce(text, "process.env.WPS_BRIDGE_RESOURCES ||= app.getAppPath();", "delete process.env.VITE_DEV_SERVER_URL;\nprocess.env.WPS_BRIDGE_RESOURCES = app.getAppPath();\nconst acceptanceMode = true;\nfunction acceptanceReadOnlyHost() { throw new Error('升级验收包禁止写入办公加载项、客户端配置或开机启动。'); }");
    text = replaceOnce(text, "let notice = '';", `let notice = ${JSON.stringify(`升级验收 ${plan.id} · ${stage} ${version} · 独立用户数据和 Bridge，办公加载项部署已禁用`)};`);
    text = replaceOnce(text, "title: '字浮 CharFloat',", `title: ${JSON.stringify(`字浮 CharFloat · 升级验收 ${stage} ${version}`)},`);
    text = replaceOnce(text, 'desktopService.closing ? null : ensureAddonUpToDate()', 'desktopService.closing || acceptanceMode ? null : ensureAddonUpToDate()');
    text = replaceOnce(text, 'function migrateUpdatedWindowsPaths() {', 'function migrateUpdatedWindowsPaths() {\n  if (acceptanceMode) return;');
    for (const entry of ["ipcMain.handle('install-addon', async () => {", "ipcMain.handle('install-office-addon', async () => {", "ipcMain.handle('ensure-addon-uptodate', async (_e, opts?: { force?: boolean }) => {"]) text = replaceOnce(text, entry, entry + '\n    acceptanceReadOnlyHost();');
    text = replaceOnce(text, "ipcMain.handle('installer:execute', (_e, options) => InstallerEngine.executeInstall(options));", "ipcMain.handle('installer:execute', (_e, options) => { acceptanceReadOnlyHost(); return InstallerEngine.executeInstall(options); });");
    text = replaceOnce(text, "ipcMain.handle('set-login', (_e, enabled) => {", "ipcMain.handle('set-login', (_e, enabled) => { acceptanceReadOnlyHost();");
    text = replaceOnce(text, 'prefs.updateLogin = app.getLoginItemSettings().openAtLogin;', 'prefs.updateLogin = false;');
    return replaceOnce(text, "tray.setToolTip('字浮 CharFloat')", "tray.setToolTip('字浮 CharFloat · 升级验收')");
  });
  edit('windows-bootstrap/core.mjs', text => text.replaceAll("'CharFloat'", JSON.stringify(plan.windowsRootName)));
  edit('windows-bootstrap/shortcut.mjs', text => text.replaceAll("'字浮 CharFloat.lnk'", JSON.stringify(plan.windowsShortcut)).replaceAll("'com.antigravity.charfloat'", JSON.stringify(plan.appId)));
  edit('windows-bootstrap/main.mjs', text => {
    text = replaceOnce(text, "import fs from 'node:fs';", "import fs from 'node:fs';\nimport os from 'node:os';");
    return replaceOnce(text, "path.join(app.getPath('appData'), 'CharFloat Setup')", `path.join(fs.realpathSync(os.tmpdir()), ${JSON.stringify(plan.runtimeDirectoryName)}, 'setup-profile')`);
  });
  edit('scripts/build-windows-bootstrap.mjs', text => replaceOnce(text, "appId:'com.charfloat.user-setup'", `appId:${JSON.stringify(plan.appId + '.setup')}`));
  overlays[stage] = edits;
}
fs.writeFileSync(path.join(output, 'isolation-overlays.json'), JSON.stringify({ baseCommit: plan.baseCommit, originalFiles: original, overlays, updaterAlgorithmChanged: false, exceptions: ['test bundle identity check', 'private runtime/ports', 'host configuration writes disabled', 'test versions/feed'] }, null, 2) + '\n');
console.log(JSON.stringify({ output, versions: plan.versions, feed: plan.feed, privateKeyCopied: false, installStarted: false }));
