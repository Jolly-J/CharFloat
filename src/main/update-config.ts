import fs from 'node:fs';
import path from 'node:path';
import { atomicWrite } from '../bridge/runtime.js';
import type { McpConfigKind } from './mcp-adapters.js';
import { mergeCodexToml } from './mcp-adapters.js';
import * as TOML from '@iarna/toml';
import { applyEdits, modify } from 'jsonc-parser';
import { parseConfigJson } from './config-file.js';
import { retargetDshMcp } from './dsh-mcp.js';

/** Retarget only path literals within the existing current-brand entry. Preserve arguments, consent and all unrelated entries. */
export function retargetUpdateConfig(file: string, kind: McpConfigKind, previous: string, next: string) {
  if (!fs.existsSync(file)) return false;
  const st = fs.lstatSync(file);
  if (!st.isFile() || st.isSymbolicLink() || st.size > 4 * 1024 * 1024) throw new Error('客户端配置类型或大小无效，未迁移。');
  const raw = fs.readFileSync(file, 'utf8');
  const replace = (value: string) => {
    const suffix = path.win32.relative(previous, value);
    if (!path.win32.isAbsolute(value) || !suffix || suffix === '..' || suffix.startsWith('..\\') || path.win32.isAbsolute(suffix)) return value;
    return path.win32.join(next, suffix);
  };
  let output = raw;
  const walk = (value: any): any => typeof value === 'string' ? replace(value) : Array.isArray(value) ? value.map(walk) : value && typeof value === 'object' ? Object.fromEntries(Object.entries(value).map(([key, item]) => [key, walk(item)])) : value;
  if (kind === 'yaml-dsh') return retargetDshMcp(file, path.win32.join(previous, '字浮 CharFloat.exe'), replace);
  if (kind === 'toml-codex') {
    let parsed: any;
    try { parsed = TOML.parse(raw); } catch { throw new Error('客户端 TOML 配置无法解析，未迁移。'); }
    const entry = parsed.mcp_servers?.charfloat, command = entry?.command;
    if (typeof command !== 'string') return false;
    if (path.win32.normalize(command).toLowerCase() !== path.win32.join(previous, '字浮 CharFloat.exe').toLowerCase()) return false;
    mergeCodexToml(file, walk(entry)); return true;
  } else {
    const parsed = parseConfigJson(raw, kind === 'json-opencode' || kind === 'json-vscode');
    const key = kind === 'json-opencode' ? 'mcp' : kind === 'json-vscode' ? 'servers' : 'mcpServers';
    const entry = parsed?.[key]?.charfloat;
    if (!entry || typeof entry !== 'object' || Array.isArray(entry)) return false;
    // Match the executable too; a custom charfloat entry that merely mentions an old directory is not ours.
    const command = kind === 'json-opencode' ? entry.command?.[0] : entry.command;
    if (typeof command !== 'string' || path.win32.normalize(command).toLowerCase() !== path.win32.join(previous, '字浮 CharFloat.exe').toLowerCase()) return false;
    const result = walk(entry);
    if (JSON.stringify(result) === JSON.stringify(entry)) return false;
    output = applyEdits(raw, modify(raw, [key, 'charfloat'], result, { formattingOptions: { insertSpaces: true, tabSize: 2 } }));
    parseConfigJson(output, kind === 'json-opencode' || kind === 'json-vscode');
  }
  if (output === raw) return false;
  if (fs.readFileSync(file, 'utf8') !== raw) throw new Error('客户端配置在迁移时被修改，未覆盖。');
  atomicWrite(file, output, true); return true;
}
