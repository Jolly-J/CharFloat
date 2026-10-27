import test from 'node:test';
import assert from 'node:assert/strict';
import { BridgeError, toolFailureResponse } from '../src/bridge/errors.js';
test('corner failure preserves retry identity and numeric receipts, never shape contents or signatures',()=>{
  const raw={success:false,error:'import failed',diagnosticType:'ppt_corner_radius',operationId:'op_8',state:'partial',liveApplied:false,saved:false,ooxmlVerified:false,previewVerified:false,backupPath:'/private/document',receipts:[{oldId:2,newId:10,phase:'pasted',name:'private',oldSignature:'secret'}]};
  const r:any=toolFailureResponse(raw,'wps_ppt_set_corner_radius');
  assert.equal(r.operationId,'op_8');assert.equal(r.state,'partial');assert.deepEqual(r.receipts,[{oldId:2,newId:10,phase:'pasted'}]);assert.equal(r.requiresReadbackBeforeRetry,true);
  assert.doesNotMatch(JSON.stringify(r),/secret|private|Signature/);assert.deepEqual(toolFailureResponse(Object.assign(new Error(r.error),r)),r);
});
test('classified errors preserve execution uncertainty through safe JSON, never leak cause/stack',()=>{
  const err=new BridgeError('failed','宿主失败',{channel:'wps-addon',method:'ppt_manage_slides',cause:{token:'secret',document:'private'}});
  const r=JSON.parse(JSON.stringify(toolFailureResponse(err)));
  assert.equal(r.error,'宿主失败');assert.equal(r.errorDetails.executed,'unknown');assert.equal(r.errorDetails.kind,'failed');
  assert.deepEqual(Object.keys(r.errorDetails).sort(),['channel','executed','kind','message','method','unsafe']);
  assert.doesNotMatch(JSON.stringify(r),/secret|private|stack|cause/);
});
test('plain exceptions and unclassified failure results stay unknown; explicit remote classification survives',()=>{
  assert.equal(toolFailureResponse(new Error('timeout')).errorDetails.executed,'unknown');
  assert.equal(toolFailureResponse({success:false,error:'failed',executed:'no'}).errorDetails.executed,'unknown');
  const r=toolFailureResponse({success:false,error:'failed',errorDetails:{kind:'failed',channel:'wps-addon',executed:'unknown',token:'secret'},partial:{completedShapeIds:[3,4],completedCount:2,document:'private',arbitrary:{token:'secret'}}});
  assert.deepEqual(r.partial,{completedShapeIds:[3,4],completedCount:2});assert.equal(r.errorDetails.executed,'unknown');assert.doesNotMatch(JSON.stringify(r),/secret|private/);
  const twice=toolFailureResponse(Object.assign(new Error(r.error),r));assert.deepEqual(twice,r);
  assert.equal(toolFailureResponse({error:'拒绝',errorDetails:{kind:'rejected'}}).errorDetails.executed,'unknown');
});

test('actual host layout failure keeps applied/before/after and safe categorized reasons',async()=>{
  const {layoutFailure}=await import('./fixtures/ppt-style/layout-failure.js');
  const raw=layoutFailure();assert.equal(raw.success,false);assert.equal(raw.after.slideWidth,1000);
  const r:any=toolFailureResponse(raw,'wps_ppt_configure_layout');
  assert.equal(r.applied.slideWidth,1000);assert.equal(r.before.slideWidth,900);assert.equal(r.after.slideWidth,1000);assert.equal(r.after.slideHeight,600);
  assert.equal(r.warningDetails[0].field,'slideHeight');assert.equal(r.warningDetails[0].code,'write_failed');assert.equal(r.warningDetails[0].reason,'host_rejected');
  assert.doesNotMatch(JSON.stringify(r),/secret|private-document|permission denied|layoutNames/);
  assert.deepEqual(toolFailureResponse(Object.assign(new Error(r.error),r)),r,'safe diagnostic survives a second stdio/MCP projection');
});
