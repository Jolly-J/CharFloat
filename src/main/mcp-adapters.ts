import fs from 'node:fs';
import { atomicWrite } from '../bridge/runtime.js';
import { legacyServerKey } from './legacy-brand.js';

/**
 * 各客户端的 MCP 配置写入适配器。
 *
 * 目前有三类真实格式：
 * - `json-mcp`：绝大多数客户端，`{"mcpServers": {"charfloat": {...}}}`；
 * - `toml-codex`：Codex CLI 只读 `~/.codex/config.toml` 的 `[mcp_servers.<name>]`；
 * - `json-opencode`：opencode 的 `{"mcp": {"<name>": {type:"local", command:[...]}}}`。
 *
 * 所有写入都遵循同一条安全约束：先读、只改本产品条目、解析失败直接抛错且不覆盖原文件。
 */

export type McpConfigKind = 'json-mcp' | 'toml-codex' | 'json-opencode';

export interface McpServerEntry {
  command: string;
  args: string[];
  env: Record<string, string>;
}

/** 本产品写入配置时使用的服务名（固定新品牌名，不随旧名变化）。 */
export const SERVER_NAME = 'charfloat';

function readJsonObject(file: string): any {
  if (!fs.existsSync(file)) return {};
  const raw = fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, '');
  // 空文件（例如 Antigravity 生成的 0 字节 mcp_config.json）属于"尚未配置"，不是损坏
  if (!raw.trim()) return {};
  let config: any;
  try { config = JSON.parse(raw); } catch { throw new Error('现有 JSON 无法解析，未覆盖。请先修复原配置。'); }
  if (!config || typeof config !== 'object' || Array.isArray(config)) throw new Error('现有配置不是 JSON 对象，未覆盖');
  return config;
}

/** 判断 mcpServers 中是否已有本产品条目（兼容旧品牌名）。 */
function hasOwnServer(servers: any): boolean {
  return Boolean(servers && typeof servers === 'object' && (servers[SERVER_NAME] || servers[legacyServerKey] || servers['wps-bridge']));
}

export function mergeMcpConfig(file: string, entry: unknown) {
  const config = readJsonObject(file);
  if (config.mcpServers !== undefined && (!config.mcpServers || typeof config.mcpServers !== 'object' || Array.isArray(config.mcpServers))) {
    throw new Error('mcpServers 格式无效，未覆盖');
  }
  // 只写当前品牌名称。迁移规则：只清掉确属本产品的遗留条目（命令/参数与原条目一致），
  // 别人的同名条目（若有）一律保留——安全性质与"合并写入、不覆盖他人"保持一致。
  const LEGACY_KEYS = [legacyServerKey, 'wps-bridge'];
  const next: Record<string, unknown> = { ...(config.mcpServers || {}) };
  for (const leg of LEGACY_KEYS) {
    const prev = next[leg];
    const legacyIsOurs = prev
      && typeof prev === 'object'
      && !Array.isArray(prev)
      && Boolean(entry && typeof entry === 'object'
        && (prev as any).command === (entry as any).command
        && JSON.stringify((prev as any).args) === JSON.stringify((entry as any).args));
    if (legacyIsOurs) delete next[leg];
  }
  next[SERVER_NAME] = entry;
  config.mcpServers = next;
  atomicWrite(file, JSON.stringify(config, null, 2) + '\n', true);
}

export function isJsonMcpConfigured(file: string): boolean {
  try {
    const config = readJsonObject(file);
    return hasOwnServer(config.mcpServers);
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
  const raw = fs.existsSync(file) ? fs.readFileSync(file, 'utf8') : '';
  const lines = raw.length ? raw.split(/\r?\n/) : [];
  const block = renderCodexBlock(entry);
  const start = lines.findIndex(line => /^\s*\[mcp_servers\.charfloat\]\s*$/.test(line));
  let nextLines: string[];

  if (start >= 0) {
    let end = lines.length;
    for (let i = start + 1; i < lines.length; i += 1) {
      if (!isTomlSection(lines[i])) continue;
      if (!tomlSectionName(lines[i]).startsWith(`mcp_servers.${SERVER_NAME}.`)) {
        end = i;
        break;
      }
    }
    const before = lines.slice(0, start);
    while (before.length && before[before.length - 1].trim() === '') before.pop();
    const after = lines.slice(end);
    while (after.length && after[0].trim() === '') after.shift();
    nextLines = [...before, ...(before.length ? [''] : []), ...block, ...(after.length ? ['', ...after] : [])];
  } else {
    const prefix = raw.replace(/[\s]+$/, '');
    nextLines = prefix ? [...prefix.split(/\r?\n/), '', ...block] : [...block];
  }

  atomicWrite(file, nextLines.join('\n') + '\n', true);
}

export function isCodexTomlConfigured(file: string): boolean {
  if (!fs.existsSync(file)) return false;
  try {
    return new RegExp(`^\\s*\\[mcp_servers\\.${SERVER_NAME}\\]\\s*$`, 'm').test(fs.readFileSync(file, 'utf8'));
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
  const config = readJsonObject(file);
  if (config.mcp !== undefined && (!config.mcp || typeof config.mcp !== 'object' || Array.isArray(config.mcp))) {
    throw new Error('mcp 格式无效，未覆盖');
  }
  const mcp: Record<string, unknown> = { ...(config.mcp || {}) };
  mcp[SERVER_NAME] = {
    type: 'local',
    command: [entry.command, ...entry.args],
    environment: entry.env,
    enabled: true,
  };
  config.mcp = mcp;
  atomicWrite(file, JSON.stringify(config, null, 2) + '\n', true);
}

export function isOpencodeConfigured(file: string): boolean {
  try {
    const config = readJsonObject(file);
    return Boolean(config.mcp && typeof config.mcp === 'object' && config.mcp[SERVER_NAME]);
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
    default: throw new Error(`未知的配置格式: ${String(kind)}`);
  }
}

/** 按客户端格式判断是否已写入本产品条目。 */
export function isMcpConfigured(kind: McpConfigKind, file: string): boolean {
  switch (kind) {
    case 'json-mcp': return isJsonMcpConfigured(file);
    case 'toml-codex': return isCodexTomlConfigured(file);
    case 'json-opencode': return isOpencodeConfigured(file);
    default: return false;
  }
}
