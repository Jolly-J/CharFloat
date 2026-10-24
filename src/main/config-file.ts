import fs from 'node:fs';
import { atomicWrite } from '../bridge/runtime.js';
import { parse, parseTree, type ParseError, type Node as JsonNode } from 'jsonc-parser';

/** Config text may contain credentials: errors deliberately omit source text and parser excerpts. */
export function readConfigText(file: string) {
  try {
    const stat = fs.lstatSync(file);
    if (!stat.isFile() || stat.isSymbolicLink() || stat.nlink !== 1 || stat.size > 4 * 1024 ** 2) throw new Error('配置文件类型或大小异常，未修改。');
    return fs.readFileSync(file, 'utf8');
  } catch (error: any) { if (error.code === 'ENOENT') return ''; throw error; }
}
export function parseConfigJson(raw: string, comments = false): Record<string, any> {
  const text = raw.replace(/^\uFEFF/, '');
  if (!text.trim()) return {};
  try {
    if (!comments) JSON.parse(text);
    const errors: ParseError[] = [];
    const value = parse(text, errors, { allowTrailingComma: comments, disallowComments: !comments });
    if (errors.length || !value || typeof value !== 'object' || Array.isArray(value)) throw new Error();
    const walk = (node?: JsonNode) => {
      if (node?.type === 'object') {
        const keys = node.children?.map(property => property.children?.[0].value) || [];
        if (new Set(keys).size !== keys.length) throw new Error();
      }
      node?.children?.forEach(walk);
    };
    walk(parseTree(text)); return value;
  } catch { throw new Error('现有 JSON 无法解析或含重复字段，未覆盖。'); }
}
export function writeConfigText(file: string, before: string, after: string) {
  if (readConfigText(file) !== before) throw new Error('配置已被其他操作修改，请刷新后重试。');
  atomicWrite(file, after, true);
  if (readConfigText(file) !== after) throw new Error('配置写入后读回不一致，请保留备份并重试。');
}
