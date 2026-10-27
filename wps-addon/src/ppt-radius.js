  // 圆角只通过原生 PPTX 参数副本导入；绝不把 Adjustments getter 当作验收。
  const pptRadiusOperations = {};
  function pptRadiusSignature(shape) {
    const line={};
    for(const key of ['Visible','Weight','Transparency','DashStyle','Style','BeginArrowheadStyle','EndArrowheadStyle']) {
      try {line[key]=shape.Line[key];}catch(e){line[key]='unavailable';}
    }
    return JSON.stringify({id:shape.Id,name:shape.Name,type:shape.Type,auto:shape.AutoShapeType,
      visible:shape.Visible,z:shape.ZOrderPosition,x:shape.Left,y:shape.Top,w:shape.Width,h:shape.Height,rotation:shape.Rotation,
      text:String(shape.TextFrame.TextRange.Text),line,style:pptReadShapeStyle(shape)});
  }
  function pptRadiusLocate(app, p) {
    const matches = [];
    for (let i=1;i<=app.Presentations.Count;i++) {
      const pres=app.Presentations.Item(i);
      if (p.targetPath ? pres.FullName===p.targetPath : pres.Name===p.presentationName) matches.push(pres);
    }
    if(matches.length!==1) throw Error('必须精确指定唯一已打开的 PPTX 文稿');
    const pres=matches[0];
    if(!pres.Path || !/\.pptx$/i.test(pres.FullName)) throw Error('圆角工具只支持已有路径的 PPTX，请先另存');
    const slide=pres.Slides.Item(pptRequireSlideIndex(pres,p.slideIndex,'设置圆角'));
    if(p.slideId!==undefined && slide.SlideID!==p.slideId) throw Error('幻灯片身份变化，请重新预检查');
    return {pres,slide};
  }
  function pptRadiusTargets(slide,names) {
    if(!Array.isArray(names)||!names.length||new Set(names).size!==names.length) throw Error('目标名称为空或重复');
    return names.map(name=>{
      const matches=[];
      for(let i=1;i<=slide.Shapes.Count;i++) if(slide.Shapes.Item(i).Name===name) matches.push(slide.Shapes.Item(i));
      if(matches.length!==1) throw Error('目标名称必须唯一：'+name);
      const s=matches[0];
      if(s.Type!==1||s.AutoShapeType!==5||s.Visible===0||String(s.TextFrame.TextRange.Text).trim()||s.Rotation!==0) throw Error('仅支持可见、无文字、无旋转的顶层 roundRect 背景');
      return {name,id:s.Id,x:s.Left,y:s.Top,w:s.Width,h:s.Height,signature:pptRadiusSignature(s)};
    });
  }
  function pptRadiusSnapshot(app,p) {
    const {pres,slide}=pptRadiusLocate(app,p),targets=pptRadiusTargets(slide,p.shapeNames);
    if(p.snapshotPath===pres.FullName) throw Error('快照不能覆盖正式文稿');
    const targetPath=pres.FullName;
    pres.SaveCopyAs(p.snapshotPath,24);
    if(pres.FullName!==targetPath) throw Error('宿主快照改变了文稿路径，已停止');
    return {targetPath:pres.FullName,slideId:slide.SlideID,shapeCount:slide.Shapes.Count,targets};
  }
  function pptRadiusById(slide,id) {
    for(let i=1;i<=slide.Shapes.Count;i++) {const s=slide.Shapes.Item(i);if(s.Id===id)return s;}
    return null;
  }
  function pptCornerRadius(app,p) {
    if(p.action==='snapshot') return pptRadiusSnapshot(app,p);
    const {pres,slide}=pptRadiusLocate(app,p);
    const key=pres.FullName+'|'+p.operationId;
    if(!/^[a-zA-Z0-9_-]{1,40}$/.test(p.operationId||'')) throw Error('operationId 无效');
    if(p.action==='status') {
      const op=pptRadiusOperations[key];
      if(!op) return {operationId:p.operationId,state:'unknown',liveApplied:false,saved:false,requiresReadbackBeforeRetry:true};
      const readback=op.receipts.map(r=>{
        const old=pptRadiusById(slide,r.oldId),fresh=pptRadiusById(slide,r.newId);
        return {oldId:r.oldId,newId:r.newId,unchanged:!!old&&!!fresh&&pptRadiusSignature(old)===r.oldSignature&&pptRadiusSignature(fresh)===r.newSignature};
      });
      return {...op,readback,targetChanged:readback.some(r=>!r.unchanged)};
    }
    if(p.action==='rollback') {
      const op=pptRadiusOperations[key];
      if(!op||!op.receipts.length||op.state==='rolledBack') throw Error('没有可恢复的本次操作记录');
      // Compare every affected identity before making the first rollback write.
      for(const r of op.receipts) for(const pair of [[r.oldId,r.oldSignature],[r.newId,r.newSignature]]) {
        const s=pptRadiusById(slide,pair[0]);
        if(!s||!pair[1]||pptRadiusSignature(s)!==pair[1]) throw Error('相关对象已有后续修改或读回未知，需要人工审阅，未回退');
      }
      op.state='rollingBack';op.saved=false;
      try {
        for(const r of [...op.receipts].reverse()) {
          pptRadiusById(slide,r.newId).Delete();
          const old=pptRadiusById(slide,r.oldId);old.Name=r.name;old.Visible=-1;
        }
        pres.Save();op.liveApplied=false;op.saved=!!pres.Saved;
        if(!op.saved) throw Error('宿主未确认回退保存');
        op.state='rolledBack';
      } catch(e) {op.state='rollbackFailed';op.failure=String(e.message||e);}
      return op;
    }
    if(p.action!=='apply') throw Error('未知圆角动作');
    if(pptRadiusOperations[key]) return pptRadiusOperations[key];
    if(slide.Shapes.Count!==p.shapeCount) throw Error('目标页对象数变化，请重新预检查');
    const current=pptRadiusTargets(slide,p.targets.map(t=>t.name));
    if(JSON.stringify(current)!==JSON.stringify(p.targets)) throw Error('目标在快照后被修改，未执行导入');
    const prefix='__cfRadius_'+p.operationId+'_';
    for(let i=1;i<=slide.Shapes.Count;i++) if(slide.Shapes.Item(i).Name.indexOf(prefix)===0) throw Error('发现同操作标识的对象，必须先读回，不能重放');
    let source=null;
    if(p.sourcePath===pres.FullName) throw Error('参数副本不能等于目标文稿');
    const originalWindow=app.ActiveWindow;
    const op={operationId:p.operationId,targetPath:pres.FullName,slideId:slide.SlideID,state:'applying',liveApplied:false,saved:false,ooxmlVerified:false,previewVerified:false,receipts:[]};
    pptRadiusOperations[key]=op;
    try {
      app.Presentations.Open(p.sourcePath,0,0,-1);
      for(let i=1;i<=app.Presentations.Count;i++) {const v=app.Presentations.Item(i);if(v.FullName===p.sourcePath)source=v;}
      if(!source) throw Error('参数副本没有实际打开');
      const sourcePage=source.Slides.Item(p.slideIndex);
      for(let i=0;i<p.targets.length;i++) {
        const d=p.targets[i],old=pptRadiusById(slide,d.id),src=sourcePage.Shapes.Item(d.name);
        if(!old) throw Error('目标在打开副本期间变化，停止');
        const expected=JSON.parse(d.signature),observed=JSON.parse(pptRadiusSignature(old));
        const insertedBefore=op.receipts.filter(r=>r.originalZ<=expected.z).length;
        if(observed.z!==expected.z+insertedBefore) throw Error('目标层级在打开副本期间变化，停止');
        observed.z=expected.z;
        if(JSON.stringify(observed)!==d.signature) throw Error('目标在打开副本期间变化，停止');
        if(src.Type!==1||src.AutoShapeType!==5||String(src.TextFrame.TextRange.Text).trim()) throw Error('参数副本对象类型不符');
        const z=old.ZOrderPosition;
        src.Copy();const fresh=slide.Shapes.Paste().Item(1);
        const r={name:d.name,oldId:old.Id,newId:fresh.Id,oldName:prefix+i,originalZ:expected.z,phase:'pasted',oldSignature:null,newSignature:null};
        op.receipts.push(r);
        fresh.Visible=0;
        fresh.Name=prefix+'pending_'+i;
        fresh.Left=d.x;fresh.Top=d.y;fresh.Width=d.w;fresh.Height=d.h;
        // Insert at the old position; preserve all other visible objects' relative order.
        fresh.ZOrder(1);for(let j=1;j<z;j++)fresh.ZOrder(2);
        r.phase='positioned';old.Name=r.oldName;old.Visible=0;r.phase='hidden';
        fresh.Name=d.name;fresh.Visible=-1;r.phase='replaced';
        if(fresh.Type!==1||fresh.AutoShapeType!==5||Math.abs(fresh.Width-d.w)>.01||Math.abs(fresh.Height-d.h)>.01||Math.abs(fresh.Left-d.x)>.01||Math.abs(fresh.Top-d.y)>.01) throw Error('导入对象读回不符');
      }
      op.liveApplied=true;pres.Save();op.saved=!!pres.Saved;
      if(!op.saved) throw Error('宿主未确认保存');
      op.state='applied';
    } catch(e) {
      op.state='partial';op.failure=String(e.message||e);op.requiresReadbackBeforeRetry=true;
      for(const r of op.receipts) if(r.phase!=='replaced') {
        try {pptRadiusById(slide,r.newId).Visible=0;}catch(hideError){op.pendingHideFailed=true;}
      }
    }
    finally {
      for(const r of op.receipts) {
        try {r.oldSignature=pptRadiusSignature(pptRadiusById(slide,r.oldId));r.newSignature=pptRadiusSignature(pptRadiusById(slide,r.newId));} catch(e) {op.requiresReadbackBeforeRetry=true;}
      }
      if(source)try {source.Close();}catch(e){op.sourceCloseFailed=true;}
      try {originalWindow.Activate();}catch(e){}
    }
    return op;
  }
