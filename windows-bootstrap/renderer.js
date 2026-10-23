/* 同一界面用于 Electron 和静态视觉验收；预览没有任何文件或系统操作。 */
const byId = id => document.getElementById(id);
let currentStage, diagnosticsStarted = false;
function enableDiagnostics() {
  if (diagnosticsStarted) return; diagnosticsStarted = true;
  window.setup.telemetry('ready', { stage: currentStage });
  setInterval(() => window.setup.telemetry('heartbeat', { stage: currentStage }), 5000);
  window.addEventListener('error', e => window.setup.telemetry('window-error', { message: e.message }));
  window.addEventListener('unhandledrejection', e => window.setup.telemetry('window-rejection', { message: e.reason?.message || String(e.reason) }));
}
function render(state) {
  const stage = state.stage;
  if (currentStage !== stage) {
    currentStage = stage;
    const content = {
      hello: ['hello', '字浮向你打招呼', '你好，很高兴见到你', '字浮，马上就绪。', '稍等片刻，我来为你准备好一切。', '01 / 打个招呼'],
      work: ['work', '字浮正在工作', '交给我，正在准备中', '把一切安排妥当。', '正在准备应用文件，并为你创建桌面快捷方式。', '02 / 配置中'],
      done: ['done', '字浮开心地庆祝', '准备完成，现在就出发', '一切就绪，开始吧。', '字浮已准备好，桌面快捷方式也已经就位。', '03 / 准备好了'],
      error: ['hello', '字浮等待重试', '还差一步，我们慢慢来', '配置暂未完成。', '已有应用和文件已保留，请关闭引导后重试。', '配置遇到问题']
    }[stage];
    if (!content) return;
    const [asset, alt, eyebrow, title, description, step] = content;
    byId('character').src = `assets/${asset}.png`; byId('character').alt = alt;
    byId('eyebrow').textContent = eyebrow; byId('title').textContent = title;
    byId('description').textContent = description; byId('step').textContent = step;
  }
  byId('close').disabled = stage === 'hello' || stage === 'work' || (stage === 'done' && !state.enabled);
  byId('progress-area').hidden = stage !== 'work'; byId('launch').hidden = stage !== 'done';
  byId('launch').disabled = !state.enabled; byId('error').hidden = stage !== 'error';
  if (stage === 'error') byId('error').textContent = state.message;
  if (stage === 'work') {
    byId('status').textContent = { verify: '正在核对安装包', extract: state.activity === 'check-file' ? '正在核对已写入文件' : '正在解压应用文件', 'verify-files': '正在复核应用文件', shortcut: '文件已核验，正在创建快捷方式' }[state.phase] || '正在完成配置';
    byId('progress').value = state.totalBytes ? state.bytes / state.totalBytes * 100 : 0;
    byId('amount').textContent = state.totalBytes && ['verify', 'extract', 'verify-files'].includes(state.phase) ? `${(state.bytes / 1048576).toFixed(1)} / ${(state.totalBytes / 1048576).toFixed(1)} MB` : '';
    byId('current-file').textContent = state.file || (state.phase === 'verify' ? '安装包 SHA-256 校验' : '');
    byId('current-file').title = state.file || '';
  }
  byId('note').textContent = stage === 'done' && !state.enabled ? '马上就能开始使用' : '仅为当前用户配置 · 无需管理员权限';
}
byId('close').onclick = () => window.setup?.close();
byId('launch').onclick = async () => {
  byId('launch').disabled = true; byId('launch').textContent = '正在启动…';
  const result = await window.setup.launch();
  if (!result.ok) { byId('launch').textContent = '重新启动'; byId('launch').disabled = false; byId('description').textContent = result.message || '启动未完成，请重试。'; }
};
if (window.setup) {
  window.setup.listen(state => {
    render(state);
    if (state.diagnosticsEnabled) enableDiagnostics();
    if (state.paintToken) requestAnimationFrame(() => requestAnimationFrame(() => window.setup.painted(state.paintToken)));
  });
  Promise.all(['hello','work','done'].map(name => { const img = new Image(); img.src = `assets/${name}.png`; return img.decode(); }))
    .then(() => requestAnimationFrame(() => requestAnimationFrame(() => window.setup.start())))
    .catch(() => { window.setup.telemetry('window-error', { message: '引导图片无法读取' }); render({ stage: 'error', message: '引导图片无法读取，请重新下载。' }); });
} else {
  // 仅在没有 preload 的普通浏览器中预览静态阶段，不伪造安装过程。
  render({ stage: new URLSearchParams(location.search).get('stage') || 'hello', enabled: true, phase: 'extract', bytes: 140 * 1048576, totalBytes: 300 * 1048576 });
}
