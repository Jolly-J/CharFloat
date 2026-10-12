// 收尾：只操作本批两份测试文件，保存/关闭/重开，复制产物并采集真实预览。
import os from 'node:os';import fs from 'node:fs';import path from 'node:path';import crypto from 'node:crypto';import assert from 'node:assert/strict';import {fileURLToPath} from 'node:url';
import {Client} from '@modelcontextprotocol/sdk/client/index.js';import {StdioClientTransport} from '@modelcontextprotocol/sdk/client/stdio.js';
const root=fileURLToPath(new URL('../../../../',import.meta.url)),dir=path.dirname(fileURLToPath(import.meta.url));
const runtime=fs.readFileSync('/tmp/charfloat-m1h-runtime-path','utf8').trim(),book='CharFloat_M1H_20261012.xlsx',control='CharFloat_M1H_ViewControl_20261012.xlsx';
const expand=value=>String(value).replace('<user-home>',os.homedir()).replace('<repo>',root.replace(/\/$/,''));
const file=expand(JSON.parse(fs.readFileSync(path.join(dir,'07-create-workbook.json'))).data.savedPath);
const controlFile=expand(JSON.parse(fs.readFileSync(path.join(dir,'208-view-control-book.json'))).data.savedPath);
const client=new Client({name:'CharFloat M1-H Final Readback',version:'1.0.0'});
await client.connect(new StdioClientTransport({command:path.join(root,'node_modules/electron/dist/Electron.app/Contents/MacOS/Electron'),args:[path.join(root,'dist/bridge/cli.cjs')],cwd:root,env:{...process.env,ELECTRON_RUN_AS_NODE:'1',WPS_BRIDGE_HOME:runtime,WPS_BRIDGE_PORT:'19890',WPS_BRIDGE_RESOURCES:root},stderr:'inherit'}));
let seq=300;
async function call(tool,args,tag){const response=await client.callTool({name:tool,arguments:args},undefined,{timeout:60000});const text=response.content.find(c=>c.type==='text')?.text;let data;try{data=JSON.parse(text);}catch{data={error:text};}
 const label=`${seq++}-${tag}`;for(const image of response.content.filter(c=>c.type==='image')){fs.writeFileSync(path.join(dir,'final-preview.png'),Buffer.from(image.data,'base64'));data.previewArtifact='final-preview.png';}
 fs.writeFileSync(path.join(dir,label+'.json'),JSON.stringify({tool,arguments:args,isError:response.isError===true,data,atUTC:new Date().toISOString()},null,2)+'\n');
 if(response.isError||data.success===false)throw Error(data.error||data.message);return data;}
const script=(code,name=book,readOnly=true)=>call('wps_execute_script',{component:'excel',workbookName:name,readOnly,code},'native-readback');
try{
 const existing=await call('wps_get_charts',{workbookName:book,sheetName:'验收数据'},'baseline-chart-cleanup-read');
 for(const name of ['Chart 6','Chart 7'])if(existing.charts.some(c=>c.shapeName===name))await call('wps_delete_chart',{workbookName:book,sheetName:'验收数据',chartName:name},'delete-baseline-chart');
 const chart=await call('wps_get_charts',{workbookName:book,sheetName:'验收数据',detail:true},'chart-final');assert.equal(chart.chartCount,1);assert.equal(chart.charts[0].title,'数量验收-已更新');assert.equal(chart.charts[0].chartType,4);assert.equal(chart.charts[0].seriesCount,1);
 await call('wps_save_workbook',{workbookName:control},'save-control');
 if(!process.env.CHARFLOAT_LAYOUT_REFRESH){const controlBefore=await call('wps_manage_workbook_views',{workbookName:control,action:'list'},'control-views');assert.equal(controlBefore.viewCount,1);
 await script(`if(wb.Name!==${JSON.stringify(control)})throw Error('目标保护');wb.Close(false);return {closed:true};`,control,false);
 await call('wps_execute_script',{component:'excel',readOnly:false,params:{filePath:controlFile},code:"const b=app.Workbooks.Open(params.filePath);return {name:b.Name};"},'reopen-control');
 const controlAfter=await call('wps_manage_workbook_views',{workbookName:control,action:'list'},'control-views-persisted');assert.deepEqual(controlAfter.views,controlBefore.views);}
 const picture=await call('wps_manage_pictures',{workbookName:book,sheetName:'验收数据',action:'list'},'picture-final');assert.ok(JSON.stringify(picture).includes('M1H_Icon'));
 const shapes=await call('wps_list_shapes',{workbookName:book,sheetName:'验收数据'},'shapes-final');assert.ok(shapes.shapes.find(x=>x.name==='M1H_Group'&&x.groupItemCount===2));
 const props=await call('wps_manage_document_properties',{workbookName:book,action:'read'},'properties-final');assert.ok(JSON.stringify(props).includes('M1-H 独立CLI验收'));
 const print=await call('wps_configure_print_layout',{workbookName:book,sheetName:'验收数据',action:'read'},'print-final');assert.equal(print.settings.printArea.replaceAll('$',''),'A1:S28');assert.equal(print.settings.orientation,2);
 const frozen=await call('wps_freeze_panes',{workbookName:book,sheetName:'验收数据',action:'read'},'freeze-final');assert.equal(frozen.splitRow,1);assert.equal(frozen.splitColumn,0);
 await call('wps_save_workbook',{workbookName:book},'save-main');
 const pre=(await script("const s=wb.Worksheets.Item('验收数据');return {saved:Boolean(wb.Saved),data:s.Range('A1:H6').Value2,formulas:s.Range('A1:H6').Formula,shapes:s.Shapes.Count,sheets:wb.Worksheets.Count,validationType:s.Range('E2:E6').Validation.Type,conditionCount:s.Range('D2:D6').FormatConditions.Count,names:wb.Names.Count,links:s.Hyperlinks.Count};")).returnValue;assert.equal(pre.saved,true);
 await script(`if(wb.Name!==${JSON.stringify(book)})throw Error('目标保护');wb.Close(false);return {closed:true};`,book,false);
 await call('wps_execute_script',{component:'excel',readOnly:false,params:{filePath:file},code:"const b=app.Workbooks.Open(params.filePath);return {name:b.Name};"},'reopen-main');
 const post=(await script("const s=wb.Worksheets.Item('验收数据');return {saved:Boolean(wb.Saved),data:s.Range('A1:H6').Value2,formulas:s.Range('A1:H6').Formula,shapes:s.Shapes.Count,sheets:wb.Worksheets.Count,validationType:s.Range('E2:E6').Validation.Type,conditionCount:s.Range('D2:D6').FormatConditions.Count,names:wb.Names.Count,links:s.Hyperlinks.Count};")).returnValue;
 assert.deepEqual(post,pre,'保存重开必须保留数据、公式和对象');
 const back=await call('wps_get_charts',{workbookName:book,sheetName:'验收数据',detail:true},'chart-persisted');assert.equal(back.charts[0].title,chart.charts[0].title);
 await script("const s=wb.Worksheets.Item('验收数据');s.Activate();s.Range('B4').Select();return {selected:app.Selection.Address()};",book,false);
 const pdf=await call('wps_export_sheet_pdf',{workbookName:book,sheetName:'验收数据',scope:'sheet'},'pdf-final');assert.equal(pdf.verifiedOnDisk,true);
 const exported=await call('wps_export_chart_image',{workbookName:book,sheetName:'验收数据',chartName:'Chart 1',outputPath:path.join(path.dirname(file),'CharFloat_M1H_chart_final.png'),format:'PNG'},'chart-export-final');assert.equal(exported.fileWritten,true);
 const preview=await call('wps_capture_sheet_preview',{workbookName:book,sheetName:'验收数据',address:'A1:S28'},'preview-final');assert.ok(preview.previewArtifact);
 const caps=await call('bridge_get_capabilities',{},'runtime-final');const expected=JSON.parse(fs.readFileSync(path.join(dir,'session.json'))).expectedBuildFingerprint;assert.equal(caps.connection.components.excel.buildFingerprint,expected);
 const files=[];for(const [source,target] of [[file,'CharFloat_M1H_20261012.xlsx'],[controlFile,'CharFloat_M1H_ViewControl_20261012.xlsx'],[pdf.outputPath,'acceptance.pdf'],[exported.outputPath,'chart.png']]){
  fs.copyFileSync(source,path.join(dir,target));files.push({file:target,size:fs.statSync(path.join(dir,target)).size,sha256:crypto.createHash('sha256').update(fs.readFileSync(path.join(dir,target))).digest('hex')});}
 const result={before:pre,after:post,dataFormulasObjectsPersisted:true,controlViewsPersisted:true,print,frozen,preview,runningCandidateMatches:true,originalWorkbookStillOpen:caps.connection.components.excel.summary.openWorkbooks.some(x=>x.name==='工作簿1'),files};
 fs.writeFileSync(path.join(dir,'final-readback.json'),JSON.stringify(result,null,2)+'\n');console.log(JSON.stringify({passed:true,dataFormulasObjectsPersisted:true,controlViewsPersisted:true,files}));
}finally{await client.close();}
