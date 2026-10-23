// 用完整 Electron 主进程验证物理 ASAR 解压；不能用 ELECTRON_RUN_AS_NODE 替代。
// 用法：electron tests/windows-bootstrap-electron.cjs <payload.zip> <manifest.json> <新的report.json>
// 仅写独立临时目录；快捷方式是无副作用适配器，不安装/启动真实应用。
const { app } = require('electron');
const fs = require('original-fs');
const path = require('node:path');
const os = require('node:os');
const assert = require('node:assert/strict');
const { pathToFileURL } = require('node:url');
const [payload, manifestFile, reportFile] = process.argv.slice(2);
if (!payload || !manifestFile || !reportFile || process.env.ELECTRON_RUN_AS_NODE) throw new Error('必须使用完整 Electron 并提供三个文件路径');
const root = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-electron-extract-'));
app.setPath('userData', path.join(root, 'profile'));
let state, calls = 0, settled = false;
const started = Date.now(), noAsar = process.noAsar;
const report = { runtime: { electron: process.versions.electron, node: process.versions.node }, events: [] };
function finish(error) {
  if (settled) return; settled = true; clearTimeout(deadline);
  report.elapsedMs = Date.now() - started; report.success = !error;
  if (error) report.failure = { message: String(error.message), phase: state?.phase, file: state?.file, bytes: state?.bytes };
  try { fs.writeFileSync(reportFile, JSON.stringify(report, null, 2), { flag: 'wx' }); }
  catch (writeError) { console.error('无法保存测试报告：', writeError.code); app.exit(1); return; }
  if (!error) fs.rmSync(root, { recursive: true, force: true });
  app.exit(error ? 1 : 0);
}
const deadline = setTimeout(() => finish(new Error('真实 Electron 解压超过 30 秒；保留本测试临时目录')), 30000);
app.whenReady().then(async () => {
  const { install, verifyInstallation } = await import(pathToFileURL(path.resolve(__dirname, '../windows-bootstrap/core.mjs')).href);
  const { createProgressSender } = await import(pathToFileURL(path.resolve(__dirname, '../windows-bootstrap/progress.mjs')).href);
  const m = JSON.parse(fs.readFileSync(manifestFile, 'utf8'));
  const progress = createProgressSender(value => report.events.push({ ...value, ms: Date.now() - started }));
  const options = { localAppData: root, payload, manifest: m, shortcut: async () => { calls++; }, onProgress: value => { state = value; progress.push(value); } };
  let result;
  try { result = await install(options); } finally { progress.dispose(); }
  assert.ok(fs.statSync(path.join(result.directory, 'resources/app.asar')).isFile());
  await verifyInstallation(result.directory, m);
  const extra = path.join(result.directory, 'local-user-note.txt'); fs.writeFileSync(extra, 'keep');
  await install({ ...options, onProgress: value => { state = value; } });
  assert.equal(fs.readFileSync(extra, 'utf8'), 'keep'); assert.equal(calls, 2);
  assert.equal(process.noAsar, noAsar, '不能全局禁用 ASAR');
  report.files = m.files.length; report.totalBytes = m.totalBytes; report.payloadSha256 = m.payloadSha256;
  report.shortcutAdapterCalls = calls; report.extraFilePreserved = true;
}).then(() => finish(), finish);
