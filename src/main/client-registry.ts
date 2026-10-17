import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';

/**
 * AI 客户端接入注册表：一个客户端一行。
 *
 * 新增客户端时只需要在这里补一项：
 * - `region` 决定它出现在「国内」还是「国外」分类里；
 * - `configKind` 决定写入格式（见 mcp-adapters.ts），目前三类真实格式 + `manual`（配置在应用内，无法用文件写入）；
 * - `paths()` 返回配置文件、技能目录和检测路径。
 */

export type ClientRegion = 'cn' | 'global';
export type ClientConfigKind = 'json-mcp' | 'toml-codex' | 'json-opencode' | 'manual';

export interface ClientPathContext {
  home: string;
  /** macOS: ~/Library/Application Support；Windows: %APPDATA% */
  appData: string;
  platform: NodeJS.Platform;
}

export interface ClientRuntimePaths {
  /** MCP 配置文件绝对路径；`manual` 类型可不提供 */
  configPath?: string;
  /** 需要同时写入的多个配置文件（例如 Windsurf 的 codeium 与 devin 两套目录）；不提供时只用 configPath */
  configPaths?: string[];
  /** 技能目录；不提供表示该客户端不支持技能同步 */
  skillsPath?: string;
  /** 任一存在即视为「已安装」（用于区分未安装与未配置） */
  detectPaths: string[];
}

export interface ClientDefinition {
  id: string;
  name: string;
  vendor: string;
  region: ClientRegion;
  configKind: ClientConfigKind;
  recommended?: boolean;
  /** 无法通过写文件接入时展示给用户的说明 */
  manualHint?: string;
  paths(ctx: ClientPathContext): ClientRuntimePaths;
}

export function clientPathContext(overrides: { home?: string; appData?: string; platform?: NodeJS.Platform } = {}): ClientPathContext {
  const home = overrides.home || os.homedir();
  const platform = overrides.platform || process.platform;
  const appData = overrides.appData || (platform === 'darwin'
    ? path.join(home, 'Library/Application Support')
    : process.env.APPDATA || path.join(home, 'AppData/Roaming'));
  return { home, appData, platform };
}

/** 豆包工作的配置在应用数据目录里，未安装时退回 ~/.doubao。 */
function resolveDoubao(ctx: ClientPathContext): ClientRuntimePaths {
  const base = path.join(ctx.appData, 'DoubaoWork');
  const workDir = path.join(base, 'Default/.doubaowork');
  const detected = fs.existsSync(workDir) || fs.existsSync(base);
  return detected
    ? {
      configPath: path.join(workDir, 'mcp.json'),
      skillsPath: path.join(workDir, 'agent_mode/workspace/.user_skills'),
      detectPaths: [base],
    }
    : {
      configPath: path.join(ctx.home, '.doubao/mcp.json'),
      skillsPath: path.join(ctx.home, '.doubao/skills'),
      detectPaths: [path.join(ctx.home, '.doubao'), base],
    };
}

/** opencode 遵循 XDG 风格：macOS/Linux 在 ~/.config，Windows 在 %APPDATA%。 */
function resolveOpencode(ctx: ClientPathContext): ClientRuntimePaths {
  const configDir = ctx.platform === 'win32' ? path.join(ctx.appData, 'opencode') : path.join(ctx.home, '.config/opencode');
  return {
    configPath: path.join(configDir, 'opencode.json'),
    skillsPath: path.join(configDir, 'skills'),
    detectPaths: [configDir],
  };
}

/**
 * Windsurf 正在从 Codeium 旧目录迁移到 Devin 配置：两套目录都可能存在。
 * 已装旧版就写旧位置，否则写官方文档里的新位置，两处都参与检测。
 */
function resolveWindsurf(ctx: ClientPathContext): ClientRuntimePaths {
  const legacyDir = path.join(ctx.home, '.codeium/windsurf');
  const devinDir = ctx.platform === 'win32' ? path.join(ctx.appData, 'devin') : path.join(ctx.home, '.config/devin');
  // 旧版编辑器侧文件与新 Cascade 托管文件都写，避免只写一份导致客户端读不到
  const configPaths: string[] = [];
  if (fs.existsSync(legacyDir)) configPaths.push(path.join(legacyDir, 'mcp_config.json'));
  if (fs.existsSync(devinDir)) configPaths.push(path.join(devinDir, 'mcp_config.json'));
  if (!configPaths.length) configPaths.push(path.join(legacyDir, 'mcp_config.json'));
  return { configPath: configPaths[0], configPaths, detectPaths: [legacyDir, devinDir] };
}

/** Antigravity 有 IDE 与独立版两个配置目录，优先写已存在的那个。 */
function resolveAntigravity(ctx: ClientPathContext): ClientRuntimePaths {
  const mainDir = path.join(ctx.home, '.gemini/antigravity');
  const ideDir = path.join(ctx.home, '.gemini/antigravity-ide');
  const configPaths: string[] = [];
  if (fs.existsSync(mainDir)) configPaths.push(path.join(mainDir, 'mcp_config.json'));
  if (fs.existsSync(ideDir)) configPaths.push(path.join(ideDir, 'mcp_config.json'));
  if (!configPaths.length) configPaths.push(path.join(mainDir, 'mcp_config.json'));
  return {
    configPath: configPaths[0],
    configPaths,
    detectPaths: [mainDir, ideDir, path.join(ctx.appData, 'Antigravity')],
  };
}

export const CLIENT_DEFINITIONS: ClientDefinition[] = [
  // ===== 国内 =====
  {
    id: 'workbuddy',
    name: 'WorkBuddy',
    vendor: '腾讯',
    region: 'cn',
    configKind: 'json-mcp',
    recommended: true,
    paths: ctx => ({
      configPath: path.join(ctx.home, '.workbuddy/mcp.json'),
      skillsPath: path.join(ctx.home, '.workbuddy/skills'),
      detectPaths: [path.join(ctx.home, '.workbuddy'), path.join(ctx.appData, 'WorkBuddy')],
    }),
  },
  {
    id: 'doubao',
    name: '豆包工作',
    vendor: '字节跳动',
    region: 'cn',
    configKind: 'json-mcp',
    paths: resolveDoubao,
  },
  {
    id: 'qwen',
    name: '千问办公',
    vendor: '阿里巴巴',
    region: 'cn',
    configKind: 'json-mcp',
    paths: ctx => ({
      // Qwen Code 与 Gemini CLI 同构，MCP 写在 settings.json 的 mcpServers 里
      configPath: path.join(ctx.home, '.qwen/settings.json'),
      skillsPath: path.join(ctx.home, '.qwen/skills'),
      detectPaths: [path.join(ctx.home, '.qwen')],
    }),
  },
  {
    id: 'kimi',
    name: 'Kimi',
    vendor: '月之暗面',
    region: 'cn',
    configKind: 'json-mcp',
    paths: ctx => ({
      configPath: path.join(ctx.home, '.kimi-code/mcp.json'),
      skillsPath: path.join(ctx.home, '.kimi-code/skills'),
      detectPaths: [path.join(ctx.home, '.kimi-code')],
    }),
  },
  {
    id: 'trae-cn',
    name: 'Trae CN',
    vendor: '字节跳动',
    region: 'cn',
    configKind: 'manual',
    manualHint: 'Trae 只支持项目级 .trae/mcp.json 或 trae:// 导入链接，需在客户端内手动添加',
    paths: ctx => ({
      detectPaths: [path.join(ctx.home, '.trae'), path.join(ctx.home, '.trae-cn'), path.join(ctx.appData, 'Trae CN')],
    }),
  },
  {
    id: 'lingma',
    name: '通义灵码',
    vendor: '阿里巴巴',
    region: 'cn',
    configKind: 'manual',
    manualHint: '通义灵码（现 Qoder CN）只能在其 IDE 内添加 MCP，无法通过写文件接入',
    paths: ctx => ({
      detectPaths: [path.join(ctx.home, '.lingma'), path.join(ctx.home, '.qoder-cn')],
    }),
  },

  // ===== 国外 =====
  {
    id: 'claude-code',
    name: 'Claude Code',
    vendor: 'Anthropic',
    region: 'global',
    configKind: 'json-mcp',
    paths: ctx => ({
      configPath: path.join(ctx.home, '.claude.json'),
      skillsPath: path.join(ctx.home, '.claude/skills'),
      detectPaths: [path.join(ctx.home, '.claude'), path.join(ctx.home, '.claude.json')],
    }),
  },
  {
    id: 'codex',
    name: 'Codex',
    vendor: 'OpenAI',
    region: 'global',
    configKind: 'toml-codex',
    paths: ctx => ({
      configPath: path.join(ctx.home, '.codex/config.toml'),
      skillsPath: path.join(ctx.home, '.codex/skills'),
      detectPaths: [path.join(ctx.home, '.codex')],
    }),
  },
  {
    id: 'claude-desktop',
    name: 'Claude Desktop',
    vendor: 'Anthropic',
    region: 'global',
    configKind: 'json-mcp',
    paths: ctx => ({
      configPath: path.join(ctx.appData, 'Claude/claude_desktop_config.json'),
      detectPaths: [path.join(ctx.appData, 'Claude')],
    }),
  },
  {
    id: 'cursor',
    name: 'Cursor',
    vendor: 'Anysphere',
    region: 'global',
    configKind: 'json-mcp',
    paths: ctx => ({
      configPath: path.join(ctx.home, '.cursor/mcp.json'),
      detectPaths: [path.join(ctx.home, '.cursor'), path.join(ctx.appData, 'Cursor')],
    }),
  },
  {
    id: 'windsurf',
    name: 'Windsurf',
    vendor: 'Cognition',
    region: 'global',
    configKind: 'json-mcp',
    paths: resolveWindsurf,
  },
  {
    id: 'gemini-cli',
    name: 'Gemini CLI',
    vendor: 'Google',
    region: 'global',
    configKind: 'json-mcp',
    paths: ctx => ({
      configPath: path.join(ctx.home, '.gemini/settings.json'),
      detectPaths: [path.join(ctx.home, '.gemini')],
    }),
  },
  {
    id: 'antigravity',
    name: 'Antigravity',
    vendor: 'Google',
    region: 'global',
    configKind: 'json-mcp',
    paths: resolveAntigravity,
  },
  {
    id: 'opencode',
    name: 'opencode',
    vendor: 'SST',
    region: 'global',
    configKind: 'json-opencode',
    paths: resolveOpencode,
  },
  {
    id: 'vscode',
    name: 'VS Code',
    vendor: 'Microsoft',
    region: 'global',
    configKind: 'json-mcp',
    paths: ctx => ({
      // 官方文档：用户级配置在 $COPILOT_HOME/mcp-config.json（默认 ~/.copilot），
      // 顶层键为 mcpServers，VS Code 与 Copilot CLI 共用。
      configPath: path.join(ctx.home, '.copilot/mcp-config.json'),
      detectPaths: [path.join(ctx.home, '.copilot'), path.join(ctx.appData, 'Code')],
    }),
  },
];

export function findClient(id: string): ClientDefinition | undefined {
  return CLIENT_DEFINITIONS.find(client => client.id === id);
}
