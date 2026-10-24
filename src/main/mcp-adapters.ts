import fs from 'node:fs';
import path from 'node:path';
import { isDeepStrictEqual } from 'node:util';
import * as TOML from '@iarna/toml';
import { applyEdits, modify } from 'jsonc-parser';
import { legacyServerKey } from './legacy-brand.js';
import { readConfigText, parseConfigJson, writeConfigText } from './config-file.js';
import { readDshMcp, mergeDshMcp } from './dsh-mcp.js';

/**
 * 各客户端的 MCP 配置写入适配器。
 *
 * 使用各客户端实际读取的格式：
 * - `json-mcp`：绝大多数客户端，`{"mcpServers": {"charfloat": {...}}}`；
 * - `toml-codex`：Codex CLI 只读 `~/.codex/config.toml` 的 `[mcp_servers.<name>]`；
 * - `json-opencode`：opencode 的 `{"mcp": {"<name>": {type:"local", command:[...]}}}`。
 * - `json-vscode`：原生用户配置的 `servers`；`yaml-dsh`：DSH 插件补丁。
 *
 * 所有写入都遵循同一条安全约束：先读、只改本产品条目、解析失败直接抛错且不覆盖原文件。
 */

export type McpConfigKind = 'json-mcp' | 'json-vscode' | 'toml-codex' | 'json-opencode' | 'yaml-dsh';

export interface McpServerEntry {
  command: string;
  args: string[];
  env: Record<string, string>;
}

/** 本产品写入配置时使用的服务名（固定新品牌名，不随旧名变化）。 */
export const SERVER_NAME = 'charfloat';

function readJsonObject(file: string, comments = false): any { return parseConfigJson(readConfigText(file), comments); }
function record(value: any): value is Record<string, any> { return Boolean(value && typeof value === 'object' && !Array.isArray(value)); }
function mergeEntry(previous: any, entry: any) {
  const merged = { ...(record(previous) ? previous : {}), ...entry };
  if (typeof entry?.command === 'string') {
    for (const key of ['url', 'serverUrl', 'httpUrl', 'headers', 'http_headers', 'bearer_token_env_var', 'oauth']) delete merged[key];
    if (merged.type && merged.type !== 'stdio') delete merged.type;
  }
  if (entry.env) merged.env = { ...(record(previous?.env) ? previous.env : {}), ...entry.env };
  return merged;
}
function editJson(raw: string, keys: string[], value: any) {
  raw = raw.replace(/^\uFEFF/, '');
  return applyEdits(raw.trim() ? raw : '{}\n', modify(raw.trim() ? raw : '{}\n', keys, value, { getInsertionIndex: () => 0, formattingOptions: { insertSpaces: true, tabSize: 2, eol: raw.includes('\r\n') ? '\r\n' : '\n' } }));
}

/** 判断 mcpServers 中是否已有本产品条目（兼容旧品牌名）。 */
function hasOwnServer(servers: any): boolean {
  return Boolean(servers && typeof servers === 'object' && (servers[SERVER_NAME] || servers[legacyServerKey] || servers['wps-bridge']));
}

export function mergeMcpConfig(file: string, entry: unknown) {
  const raw = readConfigText(file), config = parseConfigJson(raw);
  if (config.mcpServers !== undefined && (!config.mcpServers || typeof config.mcpServers !== 'object' || Array.isArray(config.mcpServers))) {
    throw new Error('mcpServers 格式无效，未覆盖');
  }
  // 只写当前品牌名称。迁移规则：只清掉确属本产品的遗留条目（命令/参数与原条目一致），
  // 别人的同名条目（若有）一律保留——安全性质与"合并写入、不覆盖他人"保持一致。
  const LEGACY_KEYS = [legacyServerKey, 'wps-bridge'];
  const next: Record<string, unknown> = { ...(config.mcpServers || {}) };
  let previousCurrent = next[SERVER_NAME];
  let output = raw;
  for (const leg of LEGACY_KEYS) {
    const prev = next[leg];
    const legacyIsOurs = prev
      && typeof prev === 'object'
      && !Array.isArray(prev)
      && Boolean(entry && typeof entry === 'object'
        && (prev as any).command === (entry as any).command
        && JSON.stringify((prev as any).args) === JSON.stringify((entry as any).args)
        && isDeepStrictEqual((prev as any).env || {}, (entry as any).env || {}));
    if (legacyIsOurs) { previousCurrent ??= prev; delete next[leg]; output = editJson(output, ['mcpServers', leg], undefined); }
  }
  next[SERVER_NAME] = mergeEntry(previousCurrent, entry);
  output = editJson(output, ['mcpServers', SERVER_NAME], next[SERVER_NAME]);
  parseConfigJson(output);
  writeConfigText(file, raw, output);
}

export function isJsonMcpConfigured(file: string): boolean {
  try {
    const config = readJsonObject(file);
    return [SERVER_NAME, legacyServerKey, 'wps-bridge'].some(name => validStdio(config.mcpServers?.[name]) && config.mcpServers[name].disabled !== true && config.mcpServers[name].enabled !== false);
  } catch {
    return false;
  }
}

/** TOML 字符串字面量：JSON 字符串的转义集合是 TOML basic string 的合法子集。 */
function tomlString(value: string): string {
  return JSON.stringify(value);
}

function tomlKey(key: string): string {
  return /^[A-Za-z0-9_-]+$/.test(key) ? key : JSON.stringify(key);
}

export function renderCodexBlock(entry: McpServerEntry): string[] {
  return [
    `[mcp_servers.${SERVER_NAME}]`,
    `command = ${tomlString(entry.command)}`,
    `args = [${entry.args.map(entry => tomlString(entry)).join(', ')}]`,
    '',
    `[mcp_servers.${SERVER_NAME}.env]`,
    ...Object.entries(entry.env).map(([key, value]) => `${tomlKey(key)} = ${tomlString(value)}`),
  ];
}

function isTomlSection(line: string): boolean {
  return /^\s*\[/.test(line);
}

function tomlSectionName(line: string): string {
  const match = line.match(/^\s*\[+([^\]]+)\]+/);
  return match ? match[1].trim() : '';
}

/**
 * 合并 Codex 的 config.toml：
 * - 已有 `[mcp_servers.charfloat]` → 替换该块（连同 `[mcp_servers.charfloat.*]` 子表）；
 * - 没有 → 追加到文件末尾；
 * - 其他段落、注释、空行原样保留。
 */
export function mergeCodexToml(file: string, entry: McpServerEntry) {
  const raw = readConfigText(file);
  let before: any;
  try { before = TOML.parse(raw); } catch { throw new Error('现有 TOML 无法解析，未覆盖。'); }
  const lines = raw.length ? raw.split(/\r?\n/) : [];
  if (before.mcp_servers !== undefined && !record(before.mcp_servers)) throw new Error('mcp_servers 不是配置表，未覆盖。');
  const merged = mergeEntry(before.mcp_servers?.[SERVER_NAME], entry);
  const block = TOML.stringify({ mcp_servers: { [SERVER_NAME]: merged } });
  const outputLines: string[] = []; let own = false, inserted = false;
  for (const line of lines) {
    if (isTomlSection(line)) {
      let section: any;
      try { section = TOML.parse(line); } catch { section = {}; }
      own = record(section.mcp_servers) && Object.hasOwn(section.mcp_servers, SERVER_NAME);
      if (own && !inserted) { outputLines.push(block.trimEnd()); inserted = true; }
    }
    if (!own || line.trimStart().startsWith('#')) outputLines.push(line);
  }
  if (!inserted) outputLines.push('', block.trimEnd());
  const output = outputLines.join('\n').replace(/\n*$/, '\n');
  // A header-looking line inside a multiline string must never cause collateral edits.
  let after: any;
  try { after = TOML.parse(output); } catch { throw new Error('无法安全合并当前 TOML 结构，原文件已保留。'); }
  const withoutOwn = (value: any) => { const copy = structuredClone(value); if (copy.mcp_servers) { delete copy.mcp_servers[SERVER_NAME]; if (!Object.keys(copy.mcp_servers).length) delete copy.mcp_servers; } return copy; };
  if (!isDeepStrictEqual(withoutOwn(before), withoutOwn(after)) || !isDeepStrictEqual(after.mcp_servers?.[SERVER_NAME], merged)) throw new Error('配置包含无法安全合并的自定义结构，原文件已保留。');
  writeConfigText(file, raw, output);
}

export function isCodexTomlConfigured(file: string): boolean {
  if (!fs.existsSync(file)) return false;
  try {
    const value: any = TOML.parse(readConfigText(file));
    return validStdio(value.mcp_servers?.[SERVER_NAME]) && value.mcp_servers[SERVER_NAME].enabled !== false;
  } catch {
    return false;
  }
}

/**
 * 合并 opencode 的 opencode.json：
 * 条目写在 `mcp.<name>` 下，本地服务为 `type: "local"`，command 是「命令 + 参数」数组，
 * 环境变量键名是 `environment`。其他配置项一律保留。
 */
export function mergeOpencodeConfig(file: string, entry: McpServerEntry) {
  const raw = readConfigText(file), config = parseConfigJson(raw, true);
  if (config.mcp !== undefined && (!config.mcp || typeof config.mcp !== 'object' || Array.isArray(config.mcp))) {
    throw new Error('mcp 格式无效，未覆盖');
  }
  const mcp: Record<string, unknown> = { ...(config.mcp || {}) };
  const previous = record(mcp[SERVER_NAME]) ? mcp[SERVER_NAME] : {};
  for (const key of ['url', 'headers', 'oauth']) delete previous[key];
  mcp[SERVER_NAME] = {
    ...previous,
    type: 'local',
    command: [entry.command, ...entry.args],
    environment: { ...(record(previous.environment) ? previous.environment : {}), ...entry.env },
    enabled: previous.enabled !== false,
  };
  const output = editJson(raw, ['mcp', SERVER_NAME], mcp[SERVER_NAME]); parseConfigJson(output, true);
  writeConfigText(file, raw, output);
}

export function isOpencodeConfigured(file: string): boolean {
  try {
    const config = readJsonObject(file, true), value = config.mcp?.[SERVER_NAME];
    return value?.type === 'local' && Array.isArray(value.command) && value.command.length > 0 && value.command.every((v: unknown) => typeof v === 'string') && value.enabled !== false;
  } catch {
    return false;
  }
}

/** 按客户端格式写入 MCP 条目。 */
export function applyMcpEntry(kind: McpConfigKind, file: string, entry: McpServerEntry) {
  switch (kind) {
    case 'json-mcp': return mergeMcpConfig(file, entry);
    case 'toml-codex': return mergeCodexToml(file, entry);
    case 'json-opencode': return mergeOpencodeConfig(file, entry);
    case 'yaml-dsh': return mergeDshMcp(file, entry);
    case 'json-vscode': {
      const raw = readConfigText(file), config = parseConfigJson(raw, true);
      if (config.servers !== undefined && !record(config.servers)) throw new Error('servers 格式无效，未覆盖。');
      const output = editJson(raw, ['servers', SERVER_NAME], { ...mergeEntry(config.servers?.[SERVER_NAME], entry), type: 'stdio' });
      parseConfigJson(output, true); return writeConfigText(file, raw, output);
    }
    default: throw new Error(`未知的配置格式: ${String(kind)}`);
  }
}

/** 按客户端格式判断是否已写入本产品条目。 */
export function isMcpConfigured(kind: McpConfigKind, file: string): boolean {
  switch (kind) {
    case 'json-mcp': return isJsonMcpConfigured(file);
    case 'toml-codex': return isCodexTomlConfigured(file);
    case 'json-opencode': return isOpencodeConfigured(file);
    case 'json-vscode': case 'yaml-dsh': return inspectMcpConfig(kind, file).status === 'valid';
    default: return false;
  }
}

function validStdio(value: any) {
  return record(value) && typeof value.command === 'string' && value.command.length > 0 && (value.args === undefined || (Array.isArray(value.args) && value.args.every((arg: unknown) => typeof arg === 'string'))) && (value.env === undefined || (record(value.env) && Object.values(value.env).every(v => typeof v === 'string')));
}
export type McpConfigInspection = { status: 'valid' | 'missing' | 'outdated' | 'disabled' | 'invalid'; detail: string };
export function inspectMcpEnablement(file: string | undefined, inspection: McpConfigInspection): McpConfigInspection {
  if (!file || !['valid', 'disabled'].includes(inspection.status) || !fs.existsSync(file)) return inspection;
  try {
    const value = readJsonObject(file)[SERVER_NAME];
    if (value !== undefined && (!record(value) || typeof value.enabled !== 'boolean')) throw new Error();
    return value?.enabled === false ? { status: 'disabled', detail: '客户端已持久停用此连接' } : inspection;
  } catch { return { status: 'invalid', detail: '客户端停用记录无法读取，未更改其设置' }; }
}
export function inspectMcpConfig(kind: McpConfigKind, file: string, expected?: McpServerEntry): McpConfigInspection {
  try {
    if (!fs.existsSync(file)) return { status: 'missing', detail: '尚未配置' };
    let value: any, root: any = {};
    if (kind === 'yaml-dsh') value = readDshMcp(file);
    else if (kind === 'toml-codex') { root = TOML.parse(readConfigText(file)); value = root.mcp_servers?.[SERVER_NAME]; }
    else {
      root = readJsonObject(file, kind === 'json-vscode' || kind === 'json-opencode');
      value = kind === 'json-opencode' ? root.mcp?.[SERVER_NAME] : kind === 'json-vscode' ? root.servers?.[SERVER_NAME] : root.mcpServers?.[SERVER_NAME];
      if (kind === 'json-opencode' && value) value = { ...value, command: value.command?.[0], args: Array.isArray(value.command) ? value.command.slice(1) : null, env: value.environment, invalidType: value.type !== 'local' };
    }
    if (!value) return { status: 'missing', detail: '尚未配置' };
    if (!validStdio(value) || value.invalidType || value.url || value.serverUrl || value.httpUrl || (kind === 'yaml-dsh' && value.transport !== 'stdio')) return { status: 'invalid', detail: '连接配置格式异常或不是本地连接' };
    const glob = (pattern: string) => new RegExp('^' + pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&').replaceAll('*', '.*').replaceAll('?', '.') + '$').test(SERVER_NAME);
    const excluded = Array.isArray(root.mcp?.excluded) && root.mcp.excluded.some((p: unknown) => typeof p === 'string' && glob(p));
    const disallowed = Array.isArray(root.mcp?.allowed) && !root.mcp.allowed.some((p: unknown) => typeof p === 'string' && glob(p));
    if (value.disabled === true || value.enabled === false || excluded || disallowed) return { status: 'disabled', detail: '客户端已停用此连接' };
    if (expected) {
      const equalPath = (a: string, b: string) => process.platform === 'win32' ? path.normalize(a).toLowerCase() === path.normalize(b).toLowerCase() : path.normalize(a) === path.normalize(b);
      if (!equalPath(value.command, expected.command) || !isDeepStrictEqual(value.args || [], expected.args) || Object.entries(expected.env).some(([key, v]) => value.env?.[key] !== v)) return { status: 'outdated', detail: '配置与当前字浮不一致' };
      if (!fs.existsSync(value.command) || !fs.existsSync(expected.args[0])) return { status: 'outdated', detail: '配置中的程序文件已失效' };
    }
    return { status: 'valid', detail: '配置已读回；客户端加载状态需在客户端确认' };
  } catch { return { status: 'invalid', detail: '配置无法解析或读取，原文件未改变' }; }
}
