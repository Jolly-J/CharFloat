import { getTools } from '../../src/bridge/catalog.js';
const tools = getTools() as any[];
const bySuffix = new Map<string, any>();
for (const t of tools) { if (t.name.startsWith('excel_')) continue; bySuffix.set(t.name, t); }
function check(name: string, schema: any, desc: string, path = '') {
  for (const [k, v] of Object.entries<any>(schema.properties || {})) {
    if (v.enum) {
      const missing = v.enum.filter((m: any) => !String(v.description || '').includes(String(m)) && !String(desc).includes(String(m)));
      if (missing.length) console.log(`${name}.${path}${k} 枚举未在说明中出现: ${missing.join(', ')}`);
    }
    if (!v.description) console.log(`${name}.${path}${k} ⚠️ 无参数说明  (type=${JSON.stringify(v.type)})`);
    if (v.properties) check(name, v, desc, `${path}${k}.`);
    if (v.items && v.items.properties) check(name, v.items, desc, `${path}${k}[].`);
  }
}
for (const t of bySuffix.values()) check(t.name, t.inputSchema, t.description);
