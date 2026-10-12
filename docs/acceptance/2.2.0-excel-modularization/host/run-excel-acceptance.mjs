// 真实 WPS 验收：只写本任务创建的工作簿；串行执行，失败后停止依赖步骤，不重放未知结果。
import os from 'node:os';import fs from 'node:fs';import path from 'node:path';import assert from 'node:assert/strict';import crypto from 'node:crypto';import {fileURLToPath} from 'node:url';
import {Client} from '@modelcontextprotocol/sdk/client/index.js';import {StdioClientTransport} from '@modelcontextprotocol/sdk/client/stdio.js';
const root=fileURLToPath(new URL('../../../../',import.meta.url)),dir=path.dirname(fileURLToPath(import.meta.url));
const runtime=fs.readFileSync('/tmp/charfloat-m1h-runtime-path','utf8').trim();
const expand=value=>String(value).replace('<user-home>',os.homedir()).replace('<repo>',root.replace(/\/$/,''));
const fileName='CharFloat_M1H_20261012.xlsx',sheet='验收数据';
const created=JSON.parse(fs.readFileSync(path.join(dir,'07-create-workbook.json'))).data;
assert.equal(created.workbookName,fileName);assert.ok(created.success);
const filePath=expand(created.savedPath),container=path.dirname(filePath);
const picturePath=path.join(container,'CharFloat_M1H_icon.png');fs.copyFileSync(path.join(root,'wps-addon/ribbon-icon@2x.png'),picturePath);
const client=new Client({name:'CharFloat M1-H Suite',version:'1.0.0'});
const transport=new StdioClientTransport({command:path.join(root,'node_modules/electron/dist/Electron.app/Contents/MacOS/Electron'),args:[path.join(root,'dist/bridge/cli.cjs')],cwd:root,env:{...process.env,ELECTRON_RUN_AS_NODE:'1',WPS_BRIDGE_HOME:runtime,WPS_BRIDGE_PORT:'19890',WPS_BRIDGE_RESOURCES:root},stderr:'inherit'});
await client.connect(transport);
const schema=new Map((await client.listTools()).tools.map(t=>[t.name,t.inputSchema]));let sequence=100;let stopped=false;
const results=[];
function write(name,data){fs.writeFileSync(path.join(dir,name),JSON.stringify(data,null,2)+'\n');}
async function call(name,args={},options={}){
 const definition=schema.get(name);assert.ok(definition,'缺少已注册工具 '+name);
 const parameters={...(definition.properties?.workbookName?{workbookName:fileName}:{}),...args};
 for(const key of Object.keys(parameters))assert.ok(definition.properties?.[key],name+' 未声明参数 '+key);
 const tag=`${sequence++}-${options.tag||name.replace(/^wps_/, '')}`,start=new Date().toISOString();let response;
 try{response=await client.callTool({name,arguments:parameters},undefined,{timeout:60000});}catch(error){write(tag+'.json',{tool:name,arguments:parameters,startUTC:start,endUTC:new Date().toISOString(),transportError:error.message,executed:'unknown; no replay'});stopped=true;throw error;}
 const text=response.content.find(c=>c.type==='text')?.text||'';let data;try{data=JSON.parse(text);}catch{data={error:text};}
 for(const [index,image] of response.content.filter(c=>c.type==='image').entries()){const extension=image.mimeType==='image/jpeg'?'jpg':'png';const target=`${tag}-${index}.${extension}`;fs.writeFileSync(path.join(dir,target),Buffer.from(image.data,'base64'));data.previewArtifact=target;}
 write(tag+'.json',{tool:name,arguments:parameters,startUTC:start,endUTC:new Date().toISOString(),isError:response.isError===true,data});
 if(options.expectError){assert.equal(response.isError,true,'预期在执行前拒绝');return data;}
 if(response.isError||data?.success===false){const error=new Error(data?.error||data?.message||text);if(/超时|连接断开|Connection closed/.test(error.message))stopped=true;throw error;}
 return data;
}
const read=(address,custom={})=>call('wps_read_range',{sheetName:sheet,address,includeFormulas:true,...custom});
const script=(code,readOnly=true)=>call('wps_execute_script',{component:'excel',code,readOnly});
async function group(name,fn){if(stopped)return;const start=new Date().toISOString();try{const detail=await fn();results.push({name,status:'passed',detail,startUTC:start,endUTC:new Date().toISOString()});console.log('通过 '+name);}catch(error){results.push({name,status:'failed',error:error.message,startUTC:start,endUTC:new Date().toISOString()});console.log('未通过 '+name+': '+error.message);}write('suite-results.json',{fileName,results,stoppedOnUnknownOutcome:stopped});}
try{
const capabilities=await call('bridge_get_capabilities');const expected=JSON.parse(fs.readFileSync(path.join(dir,'session.json'))).expectedBuildFingerprint;
assert.equal(capabilities.connection.components.excel.buildFingerprint,expected,'不得验收旧加载项');
await group('混合值公式与非法维度保护',async()=>{
 await call('wps_patch_cells',{sheetName:sheet,address:'G1:H1',values:[[7,'保留值']],formulas:[['=SUM(C2:C6)','']],reason:'混合矩阵回归'});
 const before=await read('G1:H1');assert.deepEqual(before.values,[[10,'保留值']]);assert.equal(before.formulas[0][0],'=SUM(C2:C6)');
 const rejected=await call('wps_patch_cells',{sheetName:sheet,address:'G1:H1',values:[[9]]},{expectError:true});assert.match(rejected.error,/尺寸/);
 const after=await read('G1:H1');assert.deepEqual(after.values,before.values);return {values:after.values,invalidWritePreserved:true};
});
await group('审计与回滚',async()=>{
 const update=await call('wps_patch_cells',{sheetName:sheet,address:'J1',values:[['仅用于回滚']],reason:'独立验收回滚'});assert.ok(update.auditId);
 await call('wps_rollback',{auditId:update.auditId});const back=await read('J1');assert.equal(back.values[0][0],null);return {auditId:update.auditId,restored:true};
});
await group('格式写入、行高与样式读回',async()=>{
 await call('wps_format_cells',{sheetName:sheet,address:'A1:E1',fontName:'微软雅黑',fontSize:14,bold:true,fontColor:'#FFFFFF',backgroundColor:'#1F4E78',verticalAlignment:'top',rowHeight:28,wrapText:true});
 const back=await call('wps_get_range_styles',{sheetName:sheet,address:'A1:E1',include:['fontSize','bold','fontColor','backgroundColor','verticalAlignment','rowHeight']});
 assert.equal(back.styles.fontSize,14);assert.equal(back.styles.bold,true);assert.equal(back.styles.verticalAlignment,'top');assert.equal(back.styles.rowHeight,28);return back.styles;
});
await group('富文本片段',async()=>{
 const result=await call('wps_format_text_segment',{sheetName:sheet,address:'H1',find:'保留',bold:true,fontColor:'#0066FF'});
 const back=(await script("const s=wb.Worksheets.Item('验收数据').Range('H1');const c=s.Characters(1,2);return {text:String(s.Value2),bold:Boolean(c.Font.Bold),color:Number(c.Font.Color)};")).returnValue;
 assert.equal(back.text,'保留值');assert.equal(back.bold,true);assert.equal(back.color,16737792);return {result,back};
});
await group('数据有效性写后读回与拒绝保护',async()=>{
 await call('wps_set_data_validation',{sheetName:sheet,address:'E2:E6',validationType:'list',listItems:['通过','不通过'],promptMessage:'请选择验收状态'});
 const back=await call('wps_get_range_styles',{sheetName:sheet,address:'E2:E6',include:['validation','validationViolations']});assert.equal(back.styles.validation.type,3);
 const rejected=await call('wps_set_data_validation',{sheetName:sheet,address:'E2:E6',validationType:'list'},{expectError:true});assert.match(rejected.error,/listItems/);
 const after=await call('wps_get_range_styles',{sheetName:sheet,address:'E2:E6',include:['validation']});assert.equal(after.styles.validation.formula1,back.styles.validation.formula1);
 await call('wps_set_data_validation',{sheetName:sheet,address:'C2:C6',validationType:'number_range',operator:'between',minVal:0,maxVal:9});
 return {list:after.styles.validation,invalidWritePreserved:true,check:back.validationCheck};
});
await group('条件格式与设计语言读取',async()=>{
 await call('wps_add_conditional_formatting',{sheetName:sheet,address:'D2:D6',ruleType:'cell_value',operator:'greater_than',formula1:'100',backgroundColor:'#FFE8E5',clearExisting:true});
 const back=await call('wps_add_conditional_formatting',{sheetName:sheet,address:'D2:D6',action:'read'});assert.ok(JSON.stringify(back).includes('100'));
 const design=await call('wps_get_style_token',{sheetName:sheet,sampleAddress:'A1:E6'});assert.ok(design);return {rules:back,designKeys:Object.keys(design)};
});
await group('列宽、隐藏行与冻结窗格',async()=>{
 await call('wps_auto_fit_columns',{sheetName:sheet,address:'A1:E6'});
 await call('wps_modify_rows_columns',{sheetName:sheet,targetType:'row',action:'hide',index:30,count:2});
 const hidden=(await script("const s=wb.Worksheets.Item('验收数据');return {rows:[Boolean(s.Rows.Item(30).Hidden),Boolean(s.Rows.Item(31).Hidden)],width:Number(s.Columns.Item(2).ColumnWidth)};")).returnValue;assert.deepEqual(hidden.rows,[true,true]);
 await call('wps_modify_rows_columns',{sheetName:sheet,targetType:'row',action:'unhide',index:30,count:2});
 await call('wps_freeze_panes',{sheetName:sheet,unfreeze:true});
 await call('wps_freeze_panes',{sheetName:sheet,freezeRowIndex:2,freezeColumnIndex:1});
 const frozen=await call('wps_freeze_panes',{sheetName:sheet,action:'read'});assert.equal(frozen.splitRow,1);assert.equal(frozen.splitColumn,0);return {hidden,frozen};
});
await group('复制、命名区域与超链接',async()=>{
 await call('wps_copy_range',{sheetName:sheet,sourceRange:'A1:E6',destRange:'N1:R6',copyType:'values'});
 const copied=await read('N1:R6');assert.equal(copied.values[1][2],0);assert.equal(copied.values[5][3],150);
 await call('wps_manage_named_range',{action:'add',name:'M1H_Data',refersTo:"='验收数据'!$A$1:$E$6"});
 const names=await call('wps_manage_named_range',{action:'list'});assert.ok(JSON.stringify(names).includes('M1H_Data'));
 await call('wps_manage_hyperlink',{sheetName:sheet,action:'add',address:'H3',url:'https://example.com',displayText:'验收链接'});
 const links=await call('wps_manage_hyperlink',{sheetName:sheet,action:'list'});assert.ok(JSON.stringify(links).includes('https://example.com'));return {copiedRows:copied.rowCount,names,links};
});
await group('批注写入与读回',async()=>{
 await call('wps_manage_cell_comments',{sheetName:sheet,address:'E2',action:'add',text:'独立CLI验收批注'});
 const back=await call('wps_manage_cell_comments',{sheetName:sheet,address:'E2',action:'read'});assert.ok(JSON.stringify(back).includes('独立CLI验收批注'));return back;
});
await group('筛选排序与越界保护',async()=>{
 const before=await read('A1:E6');
 await call('wps_set_filter_and_sort',{sheetName:sheet,range:'A1:E6',enableAutoFilter:true,sortRules:[{colIndex:9,order:'asc'}]},{expectError:true});
 assert.deepEqual((await read('A1:E6')).values,before.values);
 await call('wps_set_filter_and_sort',{sheetName:sheet,range:'A1:E6',enableAutoFilter:true,sortRules:[{colIndex:4,order:'asc'}]});
 const after=await read('A1:E6');assert.deepEqual(after.values.slice(1).map(row=>row[3]),[50,80,120,150,200]);
 const state=await call('wps_set_filter_and_sort',{sheetName:sheet,action:'read'});assert.equal(state.autoFilterOn,true);return {amounts:after.values.slice(1).map(row=>row[3]),state};
});
await group('结构化表与工作表生命周期',async()=>{
 await call('wps_create_sheet',{sheetName:'对象验收'});
 await call('wps_patch_cells',{sheetName:'对象验收',address:'A1:C4',values:[['项目','数量','金额'],['测试A',1,10],['测试B',2,20],['测试C',3,30]]});
 await call('wps_manage_table',{sheetName:'对象验收',action:'apply',address:'A1:C4',tableName:'M1H_Table',styleName:'TableStyleMedium2'});
 const tables=await call('wps_manage_table',{sheetName:'对象验收',action:'list'});assert.ok(JSON.stringify(tables).includes('M1H_Table'));
 await call('wps_duplicate_sheet',{sourceSheetName:'对象验收',newSheetName:'对象副本'});
 await call('wps_manage_sheet',{sheetName:'对象副本',action:'rename',newName:'对象副本_验收'});
 await call('wps_manage_sheet',{sheetName:'对象副本_验收',action:'move',targetIndex:2});
 const outline=await call('wps_get_sheet_outline',{sheetName:'对象副本_验收'});assert.equal(outline.sheetName,'对象副本_验收');return {tables,outline};
});
await group('透视表创建与字段读回',async()=>{
 await call('wps_create_sheet',{sheetName:'透视验收'});
 const pivot=await call('wps_create_pivot_table',{sourceSheetName:sheet,sourceRange:'A1:E6',destSheetName:'透视验收',destCell:'A1',rowFields:['类别'],dataFields:[{fieldName:'金额',summaryFunction:'sum',caption:'金额汇总'}]});
 const back=await call('wps_create_pivot_table',{action:'read',destSheetName:'透视验收'});assert.ok(JSON.stringify(back).includes('金额'));
 const values=await read('A1:C8',{sheetName:'透视验收'});assert.ok(values.values.flat().includes(600),'透视汇总必须为 600');return {pivot,back,values:values.values};
});
await group('统计图表创建、更新、读回与导图',async()=>{
 await call('wps_add_chart',{sheetName:sheet,dataRange:'B1:C6',chartType:'column',title:'数量验收',position:{leftCell:'G5',width:360,height:220},replaceExisting:false,hasLegend:false});
 const charts=await call('wps_get_charts',{sheetName:sheet,detail:true});const raw=charts.charts||charts.items||[];const chart=raw.find(x=>JSON.stringify(x).includes('数量验收'));assert.ok(chart,'新图表必须能按标题读回');
 const name=chart.shapeName||chart.name||chart.chartName;
 await call('wps_update_chart',{sheetName:sheet,chartName:name,title:'数量验收-已更新',chartType:'line',position:{leftCell:'G5',width:360,height:220}});
 const back=await call('wps_get_charts',{sheetName:sheet,detail:true});assert.ok(JSON.stringify(back).includes('数量验收-已更新'));
 const exported=await call('wps_export_chart_image',{sheetName:sheet,chartName:name,outputPath:path.join(container,'CharFloat_M1H_chart.png'),format:'PNG'});
 assert.ok(fs.existsSync(exported.outputPath),'导图必须真实落盘');fs.copyFileSync(exported.outputPath,path.join(dir,'chart.png'));return {chartName:name,back,exported};
});
await group('矢量形状、编组、层级与图片',async()=>{
 await call('wps_add_shape',{sheetName:sheet,name:'M1H_BoxA',shapeType:'rounded_rectangle',left:410,top:335,width:145,height:48,text:'模块化验收',fillColor:'#1F4E78',fontColor:'#FFFFFF',textAlign:'center',textVAlign:'middle'});
 await call('wps_add_shape',{sheetName:sheet,name:'M1H_BoxB',shapeType:'rectangle',left:570,top:335,width:145,height:48,text:'独立CLI通过',fillColor:'#E2F0D9',fontColor:'#1F4E78'});
 await call('wps_update_shape',{sheetName:sheet,name:'M1H_BoxA',left:415});
 const groupResult=await call('wps_group_shapes',{sheetName:sheet,shapeNames:['M1H_BoxA','M1H_BoxB'],groupName:'M1H_Group'});assert.equal(groupResult.memberCount,2);
 await call('wps_set_shape_zorder',{sheetName:sheet,name:'M1H_Group',zOrder:'bringToFront'});
 await call('wps_manage_pictures',{sheetName:sheet,action:'insert',filePath:picturePath,pictureName:'M1H_Icon',left:735,top:335,width:48,height:48});
 const back=await call('wps_list_shapes',{sheetName:sheet});assert.ok(JSON.stringify(back).includes('M1H_Group'));assert.ok(JSON.stringify(back).includes('M1H_Icon'));return {groupResult,back};
});
await group('工作簿属性与自定义视图',async()=>{
 await call('wps_manage_document_properties',{action:'apply',properties:{Title:'M1-H 独立CLI验收',M1H_Candidate:'2.2.0-excel-modularization'}});
 const props=await call('wps_manage_document_properties',{action:'read'});assert.ok(JSON.stringify(props).includes('M1-H 独立CLI验收'));
 await call('wps_manage_workbook_views',{action:'add',viewName:'M1H_View'});const views=await call('wps_manage_workbook_views',{action:'list'});assert.ok(JSON.stringify(views).includes('M1H_View'));return {props,views};
});
await group('变更检测与目标选择事件',async()=>{
 await call('wps_manage_sheet_changes',{sheetName:sheet,action:'capture',maxCells:2000});
 await call('wps_patch_cells',{sheetName:sheet,address:'G3',values:[['M1-H落盘标记']],reason:'变更检测与持久化标记'});
 const diff=await call('wps_manage_sheet_changes',{sheetName:sheet,action:'diff',maxCells:2000});assert.ok(JSON.stringify(diff).includes('M1-H落盘标记'));
 const selection=(await script("const s=wb.Worksheets.Item('验收数据');s.Activate();s.Range('B4').Select();return {workbookName:wb.Name,address:app.Selection.Address()};",false)).returnValue;assert.equal(selection.address,'$B$4');return {diff,selection};
});
await group('打印设置与真实PDF导出',async()=>{
 await call('wps_configure_print_layout',{sheetName:sheet,printArea:'A1:S28',orientation:'landscape',paperSize:'a4',fitToPagesWide:1,fitToPagesTall:1,centerHorizontally:true,printGridlines:false,centerHeader:'M1-H 独立CLI验收'});
 const readback=await call('wps_configure_print_layout',{sheetName:sheet,action:'read'});assert.equal(readback.settings.printArea.replaceAll('$',''),'A1:S28');
 const pdf=await call('wps_export_sheet_pdf',{sheetName:sheet,scope:'sheet'});assert.equal(pdf.verifiedOnDisk,true);assert.ok(fs.statSync(pdf.outputPath).size>1000);fs.copyFileSync(pdf.outputPath,path.join(dir,'acceptance.pdf'));return {readback,pdf};
});
await group('真实区域预览',async()=>{
 const preview=await call('wps_capture_sheet_preview',{sheetName:sheet,address:'A1:S28'});assert.ok(preview.previewArtifact,'必须收到真实图片，不接受单纯成功标志');return preview;
});
await group('保存、关闭重开与对象持久化',async()=>{
 await call('wps_save_workbook');const before=await read('A1:H6');
 const pre=(await script("if(wb.Name!=='CharFloat_M1H_20261012.xlsx')throw Error('目标保护');const s=wb.Worksheets.Item('验收数据');return {name:wb.Name,fullName:wb.FullName,saved:Boolean(wb.Saved),sheetCount:wb.Worksheets.Count,shapeCount:s.Shapes.Count,namesCount:wb.Names.Count};")).returnValue;assert.equal(pre.saved,true);
 await script("if(wb.Name!=='CharFloat_M1H_20261012.xlsx')throw Error('目标保护');const name=wb.Name;wb.Close(false);return {closed:name};",false);
 const reopened=await call('wps_execute_script',{component:'excel',readOnly:false,params:{filePath},code:"if(params.filePath.indexOf('CharFloat_M1H_20261012.xlsx')<0)throw Error('目标保护');const opened=app.Workbooks.Open(params.filePath);if(!opened)throw Error('宿主未打开测试文件');return {name:opened.Name,fullName:opened.FullName,sheetCount:opened.Worksheets.Count};"});assert.equal(reopened.returnValue.name,fileName);
 const after=await read('A1:H6');assert.deepEqual(after.values,before.values);assert.deepEqual(after.formulas,before.formulas);
 const post=(await script("const s=wb.Worksheets.Item('验收数据');return {sheetCount:wb.Worksheets.Count,shapeCount:s.Shapes.Count,namesCount:wb.Names.Count};")).returnValue;assert.equal(post.sheetCount,pre.sheetCount);assert.equal(post.shapeCount,pre.shapeCount);assert.equal(post.namesCount,pre.namesCount);
 assert.ok(fs.statSync(filePath).size>1000);fs.copyFileSync(filePath,path.join(dir,'CharFloat_M1H_20261012.xlsx'));
 return {pre,post,valuesAndFormulasPersisted:true,fileSize:fs.statSync(filePath).size,sha256:crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex')};
});
}finally{write('suite-results.json',{fileName,results,stoppedOnUnknownOutcome:stopped,summary:{passed:results.filter(r=>r.status==='passed').length,failed:results.filter(r=>r.status==='failed').length},scope:'macOS WPS 12.1.28496，独立 CLI，专属测试工作簿'});await client.close();}
