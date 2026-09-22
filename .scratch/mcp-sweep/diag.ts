import { getTools } from '../../src/bridge/catalog.js';
const tools = getTools() as any[];
console.log('工具数:', tools.length);
const missingType: string[] = [];
const nonStd: string[] = [];
const walk = (name: string, schema: any, path: string) => {
  for (const [k, v] of Object.entries<any>(schema.properties || {})) {
    const p = `${name}.${path}${k}`;
    if (v.type === undefined) missingType.push(`${p} => ${JSON.stringify(Object.keys(v))}`);
    else if (!['string','number','integer','boolean','array','object','null'].includes(v.type)) nonStd.push(`${p} => ${JSON.stringify(v.type)}`);
    if (v.properties) walk(name, v, `${path}${k}.`);
    if (v.items && v.items.properties) walk(name, v.items, `${path}${k}[].`);
  }
};
for (const t of tools) walk(t.name, t.inputSchema, '');
console.log('\n缺少 type 的参数 (%d):', missingType.length);
missingType.forEach(x => console.log('  ' + x));
console.log('\n非标准 type (%d):', nonStd.length);
nonStd.forEach(x => console.log('  ' + x));
const zeroReq = tools.filter(t => !t.annotations?.readOnlyHint && (t.inputSchema.required || []).length === 0);
console.log('\n零必填写工具 (%d):', zeroReq.length);
zeroReq.forEach(t => console.log('  ' + t.name));
console.log('\n描述字符统计:');
let total = 0;
for (const t of tools) total += t.description.length;
console.log('  合计', total, '均值', Math.round(total/tools.length));
const excel = tools.filter(t => t.name.startsWith('excel_'));
console.log('\nexcel_* 数量', excel.length, ' wps_* 数量', tools.filter(t=>t.name.startsWith('wps_')).length);
console.log('\n样例 excel_add_chart 描述:\n', tools.find(t=>t.name==='excel_add_chart')?.description);
console.log('\noffice_execute_script 描述:\n', JSON.stringify(tools.find(t=>t.name==='office_execute_script')?.description));
