// electron tests/windows-bootstrap-ui.cjs <output-directory> [zoom] [ui-directory]
// 静态 UI 验收；无 preload，不启动安装或应用。放大窗口保持逻辑视口，模拟显示缩放。
const { app, BrowserWindow } = require('electron');
const fs = require('original-fs'), path = require('node:path'), os = require('node:os');
const zoom = Number(process.argv[3] || 1), output = path.resolve(process.argv[2] || fs.mkdtempSync(path.join(os.tmpdir(),'charfloat-ui-results-')));
const source = path.resolve(process.argv[4] || path.join(__dirname,'../windows-bootstrap'));
if (!Number.isFinite(zoom) || zoom < 0.5 || zoom > 2) throw new Error('Invalid zoom');
const profile = fs.mkdtempSync(path.join(os.tmpdir(),'charfloat-ui-profile-')); app.setPath('userData',profile); fs.mkdirSync(output,{recursive:true});
const deadline = setTimeout(() => app.exit(1),20000);
app.whenReady().then(async () => {
  const report = [], win = new BrowserWindow({ width:Math.round(720*zoom),height:Math.round(550*zoom),show:false,frame:false,transparent:true,webPreferences:{offscreen:true,backgroundThrottling:false} });
  for (const stage of ['hello','work','done','error']) {
    await win.loadFile(path.join(source,'index.html'),{query:{stage}}); win.webContents.setZoomFactor(zoom);
    await win.webContents.executeJavaScript('document.getElementById("character").decode()');
    await win.webContents.executeJavaScript(`render({stage:${JSON.stringify(stage)},enabled:true,phase:'extract',file:'resources/app.asar',bytes:102963619,totalBytes:330283707,message:'合成验收：文件未完成，已有内容已保留。'})`);
    await new Promise(resolve => setTimeout(resolve,300));
    fs.writeFileSync(path.join(output,`${stage}-${zoom}.png`),(await win.webContents.capturePage()).toPNG(),{flag:'wx'});
    report.push(await win.webContents.executeJavaScript(`(() => { const close=document.querySelector('#close').getBoundingClientRect(),svg=document.querySelector('#close svg').getBoundingClientRect(); return { stage:${JSON.stringify(stage)},devicePixelRatio,viewportWidth:innerWidth,viewportHeight:innerHeight,scrollHeight:document.documentElement.scrollHeight,clientHeight:document.documentElement.clientHeight,centerOffsetX:svg.x+svg.width/2-close.x-close.width/2,centerOffsetY:svg.y+svg.height/2-close.y-close.height/2,closeWidth:close.width,closeHeight:close.height,closeLabel:document.querySelector('#close').getAttribute('aria-label'),footerCount:document.querySelectorAll('footer,#step').length,launchText:document.querySelector('#launch').textContent,launchVisible:!document.querySelector('#launch').hidden,closeDisabled:document.querySelector('#close').disabled,errorVisible:!document.querySelector('#error').hidden}; })()`));
  }
  const passed = report.every(value => Math.abs(value.centerOffsetX)<0.1 && Math.abs(value.centerOffsetY)<0.1 && value.footerCount===0 && value.launchText==='开始使用' && value.closeLabel==='关闭' && value.closeWidth>=31.9 && value.closeHeight>=31.9 && value.scrollHeight<=value.clientHeight+1);
  fs.writeFileSync(path.join(output,`layout-${zoom}.json`),JSON.stringify({simulation:'content zoom with constant logical viewport; not Windows OS execution',zoom,actualZoom:win.webContents.getZoomFactor(),passed,states:report},null,2),{flag:'wx'});
  if (!passed) throw new Error('UI layout check failed'); clearTimeout(deadline); win.destroy(); app.quit();
}).catch(error => {console.error(error);app.exit(1);});
app.on('will-quit',() => {try {fs.rmSync(profile,{recursive:true,force:true});} catch {console.warn('Own UI test profile cleanup incomplete.');}});
