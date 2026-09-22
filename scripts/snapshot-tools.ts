/**
 * 只读契约快照导出器。
 *
 * 用途：从真实实现（catalog.getTools / catalog.capabilities）导出对外工具名、入参 schema、
 * 能力声明与逐工具 SHA-256，用于改造前后比对兼容差异（P0.2）。
 *
 * 副作用：无。不监听端口、不连接宿主、不读写用户文档，只读取源码模块并输出 JSON。
 *
 * 用法：
 *   npm run snapshot:tools                    # 输出到 stdout
 *   npm run snapshot:tools -- <输出文件路径>   # 写到指定文件
 *
 * 易变字段（如连接状态 lastUpdated）会被剥离，保证同一源码重复导出结果一致。
 */
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { getTools, capabilities, EXCEL_METHODS } from '../src/bridge/catalog.js';

/** 递归移除运行时易变字段，避免快照因时间戳产生假差异。 */
function stable(value: any): any {
  if (Array.isArray(value)) return value.map(stable);
  if (value && typeof value === 'object') {
    const out: Record<string, any> = {};
    for (const key of Object.keys(value).sort()) {
      if (key === 'lastUpdated' || key === 'pid' || key === 'port') continue;
      out[key] = stable(value[key]);
    }
    return out;
  }
  return value;
}

function hash(value: unknown): string {
  return crypto.createHash('sha256').update(JSON.stringify(value ?? null)).digest('hex').slice(0, 16);
}

const tools = getTools() as any[];
const names = tools.map(t => t.name);
const duplicates = names.filter((name, i) => names.indexOf(name) !== i);
if (duplicates.length) throw new Error(`工具名重复，快照不可信: ${[...new Set(duplicates)].join(', ')}`);

const entries = tools.map(tool => ({
  name: tool.name,
  readOnlyHint: tool.annotations?.readOnlyHint === true,
  description: tool.description,
  inputSchema: stable(tool.inputSchema),
  schemaHash: hash(stable(tool.inputSchema)),
  descriptionHash: hash(tool.description)
}));

const snapshot = {
  generator: 'scripts/snapshot-tools.ts',
  toolCount: entries.length,
  readOnlyCount: entries.filter(e => e.readOnlyHint).length,
  writeCount: entries.filter(e => !e.readOnlyHint).length,
  excelMethods: [...EXCEL_METHODS],
  capabilities: stable(capabilities()),
  tools: entries
};

const json = JSON.stringify(snapshot, null, 2) + '\n';
const outPath = process.argv[2];
if (outPath) {
  fs.mkdirSync(path.dirname(path.resolve(outPath)), { recursive: true });
  fs.writeFileSync(path.resolve(outPath), json);
  console.log(`快照已写入 ${outPath}：${entries.length} 个工具，其中只读 ${snapshot.readOnlyCount} 个。`);
} else {
  process.stdout.write(json);
}
