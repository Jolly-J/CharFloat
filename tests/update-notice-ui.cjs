// Real branded reminder window + real preload/IPC, isolated profile. No feeds, downloads or installs.
const { app, BrowserWindow, nativeTheme } = require('electron');
const fs = require('original-fs'), path = require('node:path'), os = require('node:os'), assert = require('node:assert/strict');
const { pathToFileURL } = require('node:url');
const { build } = require('esbuild');
const output = path.resolve(process.argv[2]);
if (fs.existsSync(output)) throw new Error('Refusing to overwrite evidence');
fs.mkdirSync(output, { recursive: true });
const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-notice-ui-')); app.setPath('userData', profile);
const deadline = setTimeout(() => app.exit(1), 30000);
const settle = () => new Promise(r => setTimeout(r, 700));
app.on('window-all-closed', () => {});
app.whenReady().then(async () => {
  const moduleFile = path.resolve('dist/main/update-notice-ui-test.mjs');
  await build({ entryPoints: ['src/main/update-notice-window.ts'], bundle: true, platform: 'node', format: 'esm', external: ['electron'], outfile: moduleFile });
  const notice = await import(pathToFileURL(moduleFile).href);
  let opened = 0; notice.setupUpdateNoticeIpc(() => opened++, () => false);
  const screenshots = [];
  for (const theme of ['light', 'dark']) {
    nativeTheme.themeSource = theme;
    await notice.showUpdateNotice({ version: '2.2.903', currentVersion: '2.2.902', notes: '在线更新与彻底退出，使用起来更省心。' });
    const win = BrowserWindow.getAllWindows()[0]; await settle();
    const info = await win.webContents.executeJavaScript('window.api.updateNoticeInfo()'); assert.equal(info.version, '2.2.903'); assert.equal(win.isVisible(), true);
    const state = await win.webContents.executeJavaScript(`({ text:document.body.innerText, image:[...document.images].every(i=>i.complete&&i.naturalWidth>0), height:document.querySelector('.update-reminder').scrollHeight, viewport:window.innerHeight })`);
    assert.match(state.text, /2\.2\.903/); assert.match(state.text, /2\.2\.902/); assert.equal(state.image, true); assert.match(state.text, /字浮，又进步了一点/); assert.equal(state.height, state.viewport);
    assert.equal(opened, 0);
    const file = `notice-${theme}.png`; fs.writeFileSync(path.join(output, file), (await win.webContents.capturePage()).toPNG()); screenshots.push(file);
    await win.webContents.executeJavaScript(`[...document.querySelectorAll('button')].find(b=>b.textContent.trim()===${JSON.stringify(theme === 'light' ? '稍后' : '查看新版')}).click()`);
    await settle(); assert.equal(BrowserWindow.getAllWindows().length, 0);
    assert.equal(opened, theme === 'light' ? 0 : 1);
  }
  fs.writeFileSync(path.join(output, 'report.json'), JSON.stringify({ passed: true, actualWindowAndPreload: true, screenshots, openUpdateAreaCalls: opened, scope: 'Branded reminder dismissal/open navigation only; no download or installation' }, null, 2));
  fs.unlinkSync(moduleFile); clearTimeout(deadline); app.quit();
}).catch(error => { console.error(error); app.exit(1); });
app.on('will-quit', () => { try { fs.rmSync(profile, { recursive: true, force: true }); } catch {} });
