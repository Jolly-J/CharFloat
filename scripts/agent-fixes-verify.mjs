import fs from 'node:fs';
const token = fs.readFileSync(process.env.HOME + '/.wps-bridge/token', 'utf8').trim();
async function call(n, a) {
  const r = await fetch('http://127.0.0.1:19890/api/v1/tool/call', { method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: n, arguments: a }) });
  return r.json().catch(() => ({}));
}
async function host(code) {
  const r = await call('wps_execute_script', { component: 'excel', code, readOnly: true });
  return r.success ? r.data?.returnValue : null;
}
const WB = 'AgentExcelProbe.xlsx', SH = 'S1';
const B = { host: 'wps', workbookName: WB, sheetName: SH };
let pass = 0, total = 0;
const ck = (n, ok, d) => { total++; if (ok) pass++; console.log(`  ${ok ? '✔' : '✖'} ${n}${ok ? '' : '   → ' + String(d).slice(0, 120)}`); };

console.log('── H 类（高）──');
let r = await call('excel_format_cells', { ...B, address: 'A1', verticalAlignment: 'bottom' });
let g = await host(`return Number(wb.Worksheets.Item('${SH}').Range('A1').VerticalAlignment);`);
ck('H-2 verticalAlignment=bottom → -4107', Number(g) === -4107, `读回 ${g}`);

r = await call('excel_format_cells', { ...B, address: 'A1', backgroundColor: '不是颜色' });
ck('L-1 非法颜色写 warnings', r.success === true && r.data?.warnings?.length > 0, JSON.stringify(r.data?.warnings)?.slice(0,90));

await call('excel_patch_cells', { ...B, address: 'M1:M3', values: [['特价'],['普通'],['特价品']] });
await call('wps_execute_script', { component: 'excel', code: `wb.Worksheets.Item('${SH}').Range('A1').Select(); return 1;`, readOnly: true });
await call('excel_add_conditional_formatting', { ...B, address: 'M1:M3', ruleType: 'text_contains', containsText: '特价' });
g = await host(`const fc=wb.Worksheets.Item('${SH}').Range('M1:M3').FormatConditions; return fc.Item(fc.Count).Formula1;`);
ck('H-1 text_contains 锚定 M1', String(g).indexOf('M1') >= 0 && String(g).indexOf('Y1') < 0, `公式 ${g}`);

console.log('── M 类（中）──');
r = await call('excel_create_sheet', { ...B, sheetName: 'x/非法' });
ck('M-4 非法表名被拒', r.success === false && /非法字符/.test(String(r.error)), String(r.error).slice(0,80));
r = await call('excel_set_data_validation', { ...B, address: 'A1', validationType: 'list' });
ck('M-5 缺 listItems 被拒', r.success === false && /listItems/.test(String(r.error)), String(r.error).slice(0,80));
r = await call('excel_freeze_panes', { ...B, freezeRowIndex: 0 });
ck('L-2 freezeRowIndex:0 被拒', r.success === false, String(r.error).slice(0,80));
r = await call('excel_manage_sheet', { ...B, action: 'rename', newName: SH });
ck('M-1 rename 到已存在名被拒', r.success === false && /已被占用/.test(String(r.error)), String(r.error).slice(0,80));
r = await call('wps_manage_named_range', { workbookName: WB, action: 'add', name: 'fz_verify', refersTo: `${SH}!$A$1`, comment: '备注' });
ck('M-10 comment 不生效时如实告警', r.success === true && (r.data?.warnings||[]).some(w=>/comment 未生效/.test(w)), JSON.stringify(r.data)?.slice(0,130));
await call('wps_manage_named_range', { workbookName: WB, action: 'delete', name: 'fz_verify' });

console.log('── L 类（低）──');
r = await call('excel_save_workbook', { host: 'wps', workbookName: WB });
ck('L-4 save_workbook 读回 saved', r.success === true && r.data?.saved !== undefined, JSON.stringify(r.data)?.slice(0,90));
r = await call('excel_manage_hyperlink', { ...B, action: 'add', address: 'P1', targetAddress: `${SH}!A1` });
ck('L-5 单独 targetAddress 可用', r.success === true, String(r.error).slice(0,90));
await call('excel_manage_hyperlink', { ...B, action: 'delete', address: 'P1' });

console.log('── M-9 readOnly 脚本不再挡回滚 ──');
r = await call('wps_execute_script', { component: 'excel', code: 'return 1;', readOnly: true });
ck('M-9 readOnly 参数被接受', r.success === true, String(r.error).slice(0,90));

// 清理
await call('excel_clear_range', { ...B, address: 'A1:C3' });
await call('wps_execute_script', { component: 'excel', code: `const ws=wb.Worksheets.Item('${SH}'); try{ws.Range('M1:M3').FormatConditions.Delete();}catch(e){} ws.Range('M1:M3').Clear(); return 1;`, readOnly: false });

console.log(`\n════════ 本轮全量自测：${pass}/${total} 通过 ════════`);
process.exit(pass === total ? 0 : 1);
