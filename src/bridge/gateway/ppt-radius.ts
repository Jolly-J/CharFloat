import fs from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import type { Handler } from './types.js';
import { digest, readPptx, writePptx, inspectRadius, patchRadius, publicTargets, verifyRadius } from './pptx-radius.js';

// One tool transaction at a time, in addition to the per-call host queue. The host
// also checks the live identities immediately before import and before rollback.
let queue: Promise<unknown> = Promise.resolve();
export const setCornerRadius: Handler = ctx => {
  const job=queue.catch(()=>{}).then(()=>run(ctx));queue=job;return job;
};
async function run(ctx: Parameters<Handler>[0]) {
  const {args,callOffice}=ctx, action=args.action || 'plan';
  if(ctx.currentHost()!=='wps' || process.platform!=='darwin') throw Error('圆角工具首版仅支持 macOS WPS 演示；其他平台尚未验收');
  if(typeof args.presentationName!=='string'||!args.presentationName.trim()||!Number.isInteger(args.slideIndex)||args.slideIndex<1) throw Error('必须指定文稿名称与整数页码');
  const operationId=args.operationId || (action==='plan' ? randomUUID() : '');
  if(!/^[A-Za-z0-9_-]{1,40}$/.test(operationId)) throw Error('写入、查询和回退必须提供 operationId（1–40 位字母数字下划线或连字符）');
  const base=path.join(path.dirname(ctx.previewPath('json')),'corner-radius-operations');
  fs.mkdirSync(base,{recursive:true,mode:0o700});
  const dir=path.join(base,digest(args.presentationName+'|'+operationId)), statePath=path.join(dir,'operation.json');
  const save=(state: any)=>{ const temp=path.join(dir,randomUUID()+'.json');fs.writeFileSync(temp,JSON.stringify(state,null,2),{flag:'wx',mode:0o600});fs.renameSync(temp,statePath); };
  let state:any=fs.existsSync(statePath) ? JSON.parse(fs.readFileSync(statePath,'utf8')) : null;
  if(state && (state.presentationName!==args.presentationName||state.slideIndex!==args.slideIndex)) throw Error('operationId 已绑定其他页面');
  if(action==='status'||action==='rollback') {
    if(!state) throw Error('没有本次 operationId 的记录');
    const host:any=await callOffice('ppt_corner_radius',{action,operationId,targetPath:state.targetPath,presentationName:args.presentationName,slideIndex:args.slideIndex,slideId:state.slideId});
    // Bridge verification flags cannot be inferred from the host's mutation receipt.
    state={...state,...host,success:true,ooxmlVerified:false,previewVerified:false};
    if(action==='status'&&host.liveApplied&&host.saved&&!host.targetChanged) {
      try {
        const baseline=await readPptx(state.backupPath),p=inspectRadius(baseline,state.slideIndex,state.shapeNames,state.radiusPt);
        state.verifiedTargets=verifyRadius(baseline,await readPptx(state.targetPath),p,host.receipts);state.ooxmlVerified=true;
      } catch(e:any) {state.verificationError=String(e.message||e);}
    }
    save(state);return state;
  }
  if(!['plan','apply'].includes(action)) throw Error('未知圆角动作');
  if(!Array.isArray(args.shapeNames)||!args.shapeNames.length||args.shapeNames.length>100||args.shapeNames.some((n:any)=>typeof n!=='string'||!n.trim()||n.length>200)||new Set(args.shapeNames).size!==args.shapeNames.length||typeof args.radiusPt!=='number'||!Number.isFinite(args.radiusPt)||args.radiusPt<0) throw Error('目标名称必须唯一，半径必须为有限非负数（pt）');
  if(state) {
    if(state.radiusPt!==args.radiusPt||JSON.stringify(state.shapeNames)!==JSON.stringify(args.shapeNames)) throw Error('operationId 不能用于不同请求');
    if(action==='plan'||state.state!=='planned') return {...state,success:state.state==='planned'||state.state==='applied'&&state.ooxmlVerified===true,requiresReadbackBeforeRetry:state.state!=='applied'||state.ooxmlVerified!==true};
  }
  fs.mkdirSync(dir,{recursive:true,mode:0o700});
  const snapshotPath=path.join(dir,randomUUID()+'.pptx');
  const snap:any=await callOffice('ppt_corner_radius',{action:'snapshot',presentationName:args.presentationName,slideIndex:args.slideIndex,shapeNames:args.shapeNames,snapshotPath});
  if(state && state.targetPath!==snap.targetPath) throw Error('文稿路径与预检查不符，未执行');
  const before=await readPptx(snapshotPath), plan=inspectRadius(before,args.slideIndex,args.shapeNames,args.radiusPt);
  if(plan.slideId!==snap.slideId||plan.targets.some(t=>snap.targets.find((s:any)=>s.name===t.name)?.id!==t.shapeId)) throw Error('快照与实时目标身份不符');
  const planned={...snap,diagnosticType:'ppt_corner_radius',presentationName:args.presentationName,slideIndex:args.slideIndex,shapeNames:args.shapeNames,radiusPt:args.radiusPt,operationId,state:'planned',liveApplied:false,saved:false,ooxmlVerified:false,previewVerified:false,backupPath:snapshotPath,targets:publicTargets(plan)};
  if(action==='plan') {save(planned);return planned;}
  const sourcePath=path.join(dir,randomUUID()+'.pptx');writePptx(patchRadius(before,inspectRadius(before,args.slideIndex,args.shapeNames,args.radiusPt)),sourcePath);
  const guardPath=path.join(dir,randomUUID()+'.pptx');
  const fresh:any=await callOffice('ppt_corner_radius',{action:'snapshot',presentationName:args.presentationName,targetPath:snap.targetPath,slideIndex:args.slideIndex,slideId:snap.slideId,shapeNames:args.shapeNames,snapshotPath:guardPath});
  const latest=await readPptx(guardPath), guard=inspectRadius(latest,args.slideIndex,args.shapeNames,args.radiusPt);
  if(guard.slideId!==plan.slideId||plan.targets.some(t=>guard.targets.find(g=>g.name===t.name)?.fingerprint!==t.fingerprint)) throw Error('目标在生成副本期间改变，已停止；使用新的 operationId 重新预检查');
  // A newer backup includes unrelated edits made while preparing the source copy.
  state={...planned,backupPath:guardPath,state:'applying'};save(state);
  try {
    const applied:any=await callOffice('ppt_corner_radius',{action:'apply',presentationName:args.presentationName,operationId,targetPath:snap.targetPath,slideIndex:args.slideIndex,slideId:snap.slideId,shapeCount:fresh.shapeCount,targets:fresh.targets,sourcePath},120000);
    state={...state,...applied};
    if(!applied.liveApplied||!applied.saved) {state.success=false;state.error=applied.failure||'圆角导入未完成，请先查询操作状态';save(state);return state;}
    save(state);
    state.verifiedTargets=verifyRadius(latest,await readPptx(snap.targetPath),guard,applied.receipts);
    state.ooxmlVerified=true;save(state);
    const previewPath=path.join(dir,'preview.png');
    try {
      const preview:any=await callOffice('ppt_capture_slide_preview',{presentationName:args.presentationName,slideIndex:args.slideIndex,outputPath:previewPath});
      if(preview.success&&fs.existsSync(previewPath)) state.previewPath=previewPath;
    } catch {state.previewUnavailable=true;}
    // A generated PNG is evidence for the caller to inspect, not a visual verdict.
    state.previewVerified=false;save(state);return state;
  } catch(e:any) {
    state={...state,success:false,state:state.state==='applying'?'unknown':state.state,error:String(e.message||e),requiresReadbackBeforeRetry:true};save(state);return state;
  }
}
