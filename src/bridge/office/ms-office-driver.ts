import fs from 'node:fs';
import { currentSession } from '../context.js';
import { previewPath, resourcePath } from '../runtime.js';
import { runProcess } from '../process-runner.js';
export type OfficeComponent = 'word' | 'excel' | 'ppt';
export class MsOfficeDriver {
  private static locks = new Map<string, Partial<Record<OfficeComponent, string>>>();
  static getLockedTargets() { return { ...this.locks.get(currentSession()) }; }
  static lockTarget(component: OfficeComponent, target: string) { this.locks.set(currentSession(), { ...this.getLockedTargets(), [component]: target }); }
  static unlockTarget(component?: OfficeComponent) { const locks = this.getLockedTargets(); if (component) { delete locks[component]; this.locks.set(currentSession(), locks); } else this.locks.delete(currentSession()); }
  static async windows(payload: any) {
    const output = await runProcess('powershell.exe', ['-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-File', resourcePath('resources/office/runner.ps1')], JSON.stringify(payload), 25000);
    let result: any; try { result = JSON.parse(output.replace(/^\uFEFF/, '')); } catch { throw new Error(`Office 返回无效数据：${output.slice(0, 200)}`); }
    if (!result.success) throw new Error(result.error || 'Office 调用失败');
    return result.data;
  }
  private static async jxa(code: string) {
    const file = previewPath('js');
    try { fs.writeFileSync(file, code, { mode: 0o600 }); return JSON.parse(await runProcess('/usr/bin/osascript', ['-l', 'JavaScript', file])); }
    finally { if (fs.existsSync(file)) fs.unlinkSync(file); }
  }
  static async getStatus() {
    if (process.platform === 'win32') return { ...await this.windows({ action: 'status' }), platform: process.platform, lockedTargets: this.getLockedTargets(), validation: 'Windows 实机待验收' };
    if (process.platform !== 'darwin') return { platform: process.platform, supported: false, runningComponents: { excel: false, word: false, ppt: false } };
    const data = await this.jxa(`(function(){var result={runningComponents:{},openDocuments:{},errors:{}}; var specs=[['excel','Microsoft Excel','workbooks'],['word','Microsoft Word','documents'],['ppt','Microsoft PowerPoint','presentations']]; specs.forEach(function(s){try {var app=Application(s[1]); result.runningComponents[s[0]]=app.running(); result.openDocuments[s[0]]=[]; if(app.running()){var docs=app[s[2]];for(var i=0;i<docs.length;i++)result.openDocuments[s[0]].push(docs[i].name());}}catch(e){result.errors[s[0]]=e.message;}});return JSON.stringify(result);})()`);
    return { ...data, platform: process.platform, lockedTargets: this.getLockedTargets() };
  }
  static async executeScript(component: OfficeComponent, script: string, targetName?: string, params: any = {}) {
    const started = Date.now(); const effective = targetName || this.getLockedTargets()[component];
    let result: any;
    if (process.platform === 'win32') result = await this.windows({ action: 'script', component, script, targetName: effective, params });
    else if (process.platform === 'darwin') {
      const appName = { word: 'Microsoft Word', excel: 'Microsoft Excel', ppt: 'Microsoft PowerPoint' }[component];
      const collection = { word: 'documents', excel: 'workbooks', ppt: 'presentations' }[component];
      const active = { word: 'activeDocument', excel: 'activeWorkbook', ppt: 'activePresentation' }[component];
      result = await this.jxa(`(function(){var app=Application(${JSON.stringify(appName)});if(!app.running())throw new Error('请先打开目标 Office 组件');var targetName=${JSON.stringify(effective || null)},params=${JSON.stringify(params)},target=null;var items=app.${collection};if(targetName){for(var i=0;i<items.length;i++){if(items[i].name()===targetName){if(target)throw new Error('目标名称不唯一');target=items[i];}}if(!target)throw new Error('找不到指定目标文档');}else{target=app.${active}();}var doc=target,wb=target,pres=target;var value=(function(){${script}\n})();return JSON.stringify(value===undefined?null:value);})()`);
    } else throw new Error('当前平台不支持 Microsoft Office 自动化');
    return { success: true, component, executionTimeMs: Date.now() - started, returnValue: result };
  }
  static async capturePptSlide(slideIndex = 1, presentationName?: string) {
    if (!Number.isInteger(slideIndex) || slideIndex < 1) throw new Error('slideIndex 必须为正整数');
    const out = previewPath('png');
    if (process.platform === 'win32') await this.windows({ action: 'export', component: 'ppt', targetName: presentationName || this.getLockedTargets().ppt, slideIndex, outputPath: out });
    else if (process.platform === 'darwin') await this.executeScript('ppt', `pres.slides[params.index-1].export({to:Path(params.output),as:'PNG'});return true;`, presentationName, { index: slideIndex, output: out });
    else throw new Error('当前平台不支持 PowerPoint');
    if (!fs.existsSync(out)) throw new Error('PowerPoint 未生成预览文件');
    return { success: true, imagePath: out, imageBase64: fs.readFileSync(out).toString('base64') };
  }
}
