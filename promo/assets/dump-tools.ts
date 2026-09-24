/**
 * 导出运行时工具契约（只读，不监听端口、不连接宿主）。
 *
 * 用途：施工脚本必须按**真实**工具名与参数写，不能凭记忆。本脚本从 catalog 读出
 * 权威工具清单，落到 promo/assets/tool-schema.json，供人工核对与脚本生成。
 *
 * 用法：node --import tsx promo/assets/dump-tools.ts
 */
import fs from 'node:fs';
import path from 'node:path';
import { getTools, capabilities } from '../../src/bridge/catalog.js';

const out = path.resolve(import.meta.dirname, 'tool-schema.json');
const tools = getTools() as Array<{ function?: { name?: string } } & Record<string, unknown>>;
const names = tools.map(t => t.function?.name ?? '(unknown)').filter(Boolean);

fs.writeFileSync(out, JSON.stringify({ tools, capabilities: capabilities() }, null, 2), 'utf8');
console.log(`工具总数: ${names.length}`);
console.log(`已写入: ${out}`);
console.log('\n--- 表格 wps_* ---');
console.log(names.filter(n => n.startsWith('wps_')).join('\n'));
console.log('\n--- 文档 word_* ---');
console.log(names.filter(n => n.includes('word')).join('\n'));
console.log('\n--- 演示 ppt_* ---');
console.log(names.filter(n => n.includes('ppt')).join('\n'));
