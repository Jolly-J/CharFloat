import { isMap, isSeq, isScalar, parseDocument, type YAMLMap } from 'yaml';
import path from 'node:path';
import type { McpServerEntry } from './mcp-adapters.js';
import { readConfigText, writeConfigText } from './config-file.js';

const PLUGIN = '@deepseek-ai/dsh-mcp-client';
const ID = 'mcp-charfloat';
function document(raw: string) {
  let doc = parseDocument(raw, { uniqueKeys: true, logLevel: 'silent' });
  if (doc.errors.length || (doc.contents && !isSeq(doc.contents))) throw new Error('DSH 配置不是有效的补丁列表，未修改。');
  // Unknown tags (including !!js) are retained as YAML nodes. Never evaluate expressions.
  if (!doc.contents) {
    doc = parseDocument(raw + '\n[]\n', { uniqueKeys: true, logLevel: 'silent' });
    (doc.contents as any).flow = false;
  }
  doc.toJS({ maxAliasCount: 50 });
  return doc;
}
function entries(doc: ReturnType<typeof document>) {
  const found: YAMLMap[] = [];
  for (const patch of (doc.contents as any).items) {
    if (!isMap(patch)) throw new Error('DSH 补丁格式无效，未修改。');
    const insert = patch.get('insert', true);
    if (!insert) continue;
    if (!isSeq(insert)) throw new Error('DSH insert 格式无效，未修改。');
    for (const item of insert.items) {
      if (!isMap(item)) continue;
      if (item.get('id') === ID || (item.get('name') === PLUGIN && item.getIn(['config', 'serverName']) === 'charfloat')) found.push(item);
    }
  }
  if (found.length > 1) throw new Error('DSH 中存在重复的字浮连接，请先在客户端整理。');
  return found;
}
export function readDshMcp(file: string) {
  try {
    const doc = document(readConfigText(file)), item = entries(doc)[0];
    if (!item) return undefined;
    if (item.get('name') !== PLUGIN) throw new Error();
    const disabledNode = item.get('disabled', true);
    if (disabledNode && (!isScalar(disabledNode) || disabledNode.tag || typeof disabledNode.value !== 'boolean')) throw new Error();
    const value = item.toJSON() as any;
    let disabled = value.disabled === true;
    // Explicit user overrides are authoritative. Unknown transforms cannot be called verified.
    for (const patch of (doc.contents as any).items) {
      if (isMap(patch) && patch.get('id') === value.id) {
        const override = patch.toJSON() as any;
        if (Object.keys(override).some(key => !['id', 'disabled'].includes(key))) throw new Error();
        const node = patch.get('disabled', true);
        if (node && (!isScalar(node) || node.tag || typeof node.value !== 'boolean')) throw new Error();
        if (typeof override.disabled === 'boolean') disabled = override.disabled;
      }
    }
    return { ...value.config, disabled };
  } catch { throw new Error('DSH 连接配置无法安全确认，原文件已保留。'); }
}
export function mergeDshMcp(file: string, entry: McpServerEntry) {
  const raw = readConfigText(file), doc = document(raw), previous = entries(doc)[0];
  if (previous) {
    readDshMcp(file); // Ambiguous overrides must fail before any write.
    if (previous.get('name') !== PLUGIN) throw new Error('DSH 中同名条目属于其他插件，未覆盖。');
    const config = previous.get('config', true);
    if (!isMap(config)) throw new Error('DSH 字浮条目的配置格式无效，未覆盖。');
    config.set('transport', 'stdio'); config.set('serverName', 'charfloat'); config.set('command', entry.command); config.set('args', entry.args);
    for (const key of ['url', 'headers']) config.delete(key);
    const env = config.get('env', true);
    if (env && !isMap(env)) throw new Error('DSH 字浮环境变量不是映射，未覆盖。');
    if (!env) config.set('env', entry.env);
    else for (const [key, value] of Object.entries(entry.env)) env.set(key, value);
  } else {
    (doc.contents as any).add(doc.createNode({ insert: [{ id: ID, name: PLUGIN, config: { serverName: 'charfloat', transport: 'stdio', ...entry } }] }));
  }
  const output = doc.toString({ lineWidth: 0 });
  document(output);
  writeConfigText(file, raw, output);
}

export function retargetDshMcp(file: string, previousExecutable: string, replace: (value: string) => string) {
  const raw = readConfigText(file), doc = document(raw), item = entries(doc)[0];
  if (!item || item.get('name') !== PLUGIN) return false;
  const config = item.get('config', true);
  if (!isMap(config)) return false;
  const command = config.get('command');
  if (typeof command !== 'string' || path.win32.normalize(command).toLowerCase() !== path.win32.normalize(previousExecutable).toLowerCase()) return false;
  let changed = false;
  const visit = (node: any) => {
    if (isScalar(node) && typeof node.value === 'string' && (!node.tag || node.tag === 'tag:yaml.org,2002:str')) {
      const value = replace(node.value); changed ||= value !== node.value; node.value = value;
    } else if (isMap(node)) for (const pair of node.items) visit(pair.value);
    else if (isSeq(node)) for (const value of node.items) visit(value);
  };
  visit(config);
  if (!changed) return false;
  const output = doc.toString({ lineWidth: 0 }); document(output); writeConfigText(file, raw, output); return true;
}
