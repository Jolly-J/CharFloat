import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import type { McpConfigKind } from './mcp-adapters.js';

/**
 * AI 客户端接入注册表：一个客户端一行。
 *
 * 新增客户端时只需要在这里补一项：
 * - `region` 决定它出现在「国内」还是「国外」分类里；
 * - `configKind` 决定写入格式（见 mcp-adapters.ts）；`manual` 保留应用内接入；
 * - `paths()` 返回配置文件、技能目录和检测路径。
 */

export type ClientRegion = 'cn' | 'global';
export type ClientConfigKind = McpConfigKind | 'manual';

export interface ClientPathContext {
  home: string;
  /** macOS: ~/Library/Application Support；Windows: %APPDATA% */
  appData: string;
  platform: NodeJS.Platform;
  env?: NodeJS.ProcessEnv;
}

export interface ClientRuntimePaths {
  configKind?: McpConfigKind;
  configTargets?: { file: string; kind: McpConfigKind }[];
  /** 独立的持久停用记录；只检查，不替用户启用。 */
  enablementPath?: string;
  /** MCP 配置文件绝对路径；`manual` 类型可不提供 */
  configPath?: string;
  /** 需要同时写入的多个配置文件（例如 Windsurf 的 codeium 与 devin 两套目录）；不提供时只用 configPath */
  configPaths?: string[];
  /** 技能目录；不提供表示该客户端不支持技能同步 */
  skillsPath?: string;
  skillsPaths?: string[];
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
  documentation?: string;
  paths(ctx: ClientPathContext): ClientRuntimePaths;
}

export function clientPathContext(overrides: { home?: string; appData?: string; platform?: NodeJS.Platform; env?: NodeJS.ProcessEnv } = {}): ClientPathContext {
  const platform = overrides.platform || process.platform;
  const env = overrides.env || (overrides.home ? {} : process.env);
  const ops = platform === 'win32' ? path.win32 : path.posix;
  const home = overrides.home || (platform === 'win32' && env.USERPROFILE ? env.USERPROFILE : os.homedir());
  const appData = overrides.appData || (platform === 'darwin' ? ops.join(home, 'Library/Application Support')
    : platform === 'win32' ? env.APPDATA || ops.join(home, 'AppData/Roaming') : env.XDG_CONFIG_HOME || ops.join(home, '.config'));
  return { home, appData, platform, env };
}
const joinClientPath = (ctx: ClientPathContext, ...parts: string[]) => (ctx.platform === 'win32' ? path.win32 : path.posix).join(...parts);
const absoluteClientPath = (ctx: ClientPathContext, value: string) => (ctx.platform === 'win32' ? path.win32 : path.posix).isAbsolute(value);
const homeRoot = (ctx: ClientPathContext, key: string, fallback: string) => {
  const value = ctx.env?.[key];
  if (value && !absoluteClientPath(ctx, value)) throw new Error(`${key} 不是绝对路径，请核对客户端自定义目录。`);
  return value || joinClientPath(ctx, ctx.home, fallback);
};

/** 豆包工作的配置在应用数据目录里，未安装时退回 ~/.doubao。 */
function resolveDoubao(ctx: ClientPathContext): ClientRuntimePaths {
  const base = joinClientPath(ctx, ctx.appData, 'DoubaoWork');
  const workDir = joinClientPath(ctx, base, 'Default/.doubaowork');
  const detected = fs.existsSync(workDir) || fs.existsSync(base);
  return detected
    ? {
      configPath: joinClientPath(ctx, workDir, 'mcp.json'),
      skillsPath: joinClientPath(ctx, workDir, 'agent_mode/workspace/.user_skills'),
      detectPaths: [base],
    }
    : {
      configPath: joinClientPath(ctx, ctx.home, '.doubao/mcp.json'),
      skillsPath: joinClientPath(ctx, ctx.home, '.doubao/skills'),
      detectPaths: [joinClientPath(ctx, ctx.home, '.doubao'), base],
    };
}

/** OpenCode 的用户配置在 XDG config 或 ~/.config；Windows 也不是 Electron 的 APPDATA 缓存目录。 */
function resolveOpencode(ctx: ClientPathContext): ClientRuntimePaths {
  const configDir = joinClientPath(ctx, homeRoot(ctx, 'XDG_CONFIG_HOME', '.config'), 'opencode');
  const custom = ctx.env?.OPENCODE_CONFIG;
  const config = custom && absoluteClientPath(ctx, custom) ? custom : joinClientPath(ctx, configDir, fs.existsSync(joinClientPath(ctx, configDir, 'opencode.jsonc')) ? 'opencode.jsonc' : 'opencode.json');
  if (custom && !absoluteClientPath(ctx, custom)) throw new Error('OPENCODE_CONFIG 不是绝对路径，请核对客户端自定义配置。');
  const existing = ['opencode.json', 'opencode.jsonc'].map(name => joinClientPath(ctx, configDir, name)).filter(file => fs.existsSync(file));
  return {
    configPath: config,
    configPaths: custom ? [custom] : existing.length ? existing : [config],
    skillsPath: joinClientPath(ctx, configDir, 'skills'),
    detectPaths: [configDir],
  };
}

/**
 * Windsurf 正在从 Codeium 旧目录迁移到 Devin 配置：两套目录都可能存在。
 * 已装旧版就写旧位置，否则写官方文档里的新位置，两处都参与检测。
 */
function resolveWindsurf(ctx: ClientPathContext): ClientRuntimePaths {
  const legacyDir = joinClientPath(ctx, ctx.home, '.codeium/windsurf');
  const devinDir = ctx.platform === 'win32' ? joinClientPath(ctx, ctx.appData, 'devin') : joinClientPath(ctx, homeRoot(ctx, 'XDG_CONFIG_HOME', '.config'), 'devin');
  // 旧版编辑器侧文件与新 Cascade 托管文件都写，避免只写一份导致客户端读不到
  const configPaths: string[] = [];
  if (fs.existsSync(legacyDir)) configPaths.push(joinClientPath(ctx, legacyDir, 'mcp_config.json'));
  if (fs.existsSync(devinDir) || !configPaths.length) configPaths.push(joinClientPath(ctx, devinDir, 'mcp_config.json'));
  return { configPath: configPaths[0], configPaths, skillsPath: joinClientPath(ctx, fs.existsSync(legacyDir) && !fs.existsSync(devinDir) ? legacyDir : devinDir, 'skills'), detectPaths: [legacyDir, devinDir] };
}

/** 当前 Antigravity 使用 .gemini/config；仅为已存在的旧入口保留兼容写入。 */
function resolveAntigravity(ctx: ClientPathContext): ClientRuntimePaths {
  const mainDir = joinClientPath(ctx, ctx.home, '.gemini/antigravity');
  const ideDir = joinClientPath(ctx, ctx.home, '.gemini/antigravity-ide');
  const currentDir = joinClientPath(ctx, ctx.home, '.gemini/config');
  const configPaths: string[] = [];
  if (fs.existsSync(currentDir)) configPaths.push(joinClientPath(ctx, currentDir, 'mcp_config.json'));
  if (fs.existsSync(mainDir)) configPaths.push(joinClientPath(ctx, mainDir, 'mcp_config.json'));
  if (fs.existsSync(ideDir)) configPaths.push(joinClientPath(ctx, ideDir, 'mcp_config.json'));
  if (!configPaths.length) configPaths.push(joinClientPath(ctx, currentDir, 'mcp_config.json'));
  return {
    configPath: configPaths[0],
    configPaths,
    skillsPath: joinClientPath(ctx, currentDir, 'skills'),
    skillsPaths: [joinClientPath(ctx, currentDir, 'skills'), ...(fs.existsSync(joinClientPath(ctx, ctx.home, '.gemini/antigravity-cli')) ? [joinClientPath(ctx, ctx.home, '.gemini/antigravity-cli/skills')] : [])],
    detectPaths: [currentDir, mainDir, ideDir, joinClientPath(ctx, ctx.appData, 'Antigravity')],
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
    documentation: 'https://www.codebuddy.cn/docs/workbuddy/From-Beginner-to-Expert-Guide/Function-Description/MCP-Guide',
    paths: ctx => ({
      configPath: joinClientPath(ctx, ctx.home, '.workbuddy/mcp.json'),
      skillsPath: joinClientPath(ctx, ctx.home, '.workbuddy/skills'),
      detectPaths: [joinClientPath(ctx, ctx.home, '.workbuddy'), joinClientPath(ctx, ctx.appData, 'WorkBuddy')],
    }),
  },
  {
    id: 'doubao',
    name: '豆包工作',
    vendor: '字节跳动',
    region: 'cn',
    configKind: 'manual',
    manualHint: '在豆包工作中添加本地连接器。',
    paths: resolveDoubao,
  },
  {
    id: 'qwen',
    name: 'Qwen Code',
    vendor: '阿里巴巴',
    region: 'cn',
    configKind: 'json-mcp',
    documentation: 'https://github.com/QwenLM/qwen-code/blob/main/docs/users/features/mcp.md',
    paths: ctx => ({
      // Qwen Code 与 Gemini CLI 同构，MCP 写在 settings.json 的 mcpServers 里
      configPath: joinClientPath(ctx, ctx.home, '.qwen/settings.json'),
      skillsPath: joinClientPath(ctx, ctx.home, '.qwen/skills'),
      detectPaths: [joinClientPath(ctx, ctx.home, '.qwen')],
    }),
  },
  {
    id: 'qwen-work', name: '千问办公', vendor: '阿里巴巴', region: 'cn', configKind: 'manual',
    manualHint: '扩展 → 连接器 → 添加，粘贴配置。',
    documentation: 'https://help.aliyun.com/zh/qwenwork/desktop-connectors',
    paths: ctx => ({ detectPaths: [joinClientPath(ctx, ctx.appData, 'QwenWork')] }),
  },
  {
    id: 'deepseek-harness', name: 'DeepSeek Harness', vendor: 'DeepSeek', region: 'cn', configKind: 'yaml-dsh',
    documentation: 'https://github.com/deepseek-ai/deepseek-harness/blob/master/packages/mcp/mcp-client/README.md',
    paths: ctx => { const dir = homeRoot(ctx, 'DSH_HOME', '.dsh'); return { configPath: joinClientPath(ctx, dir, 'cordis.patch.yml'), skillsPath: joinClientPath(ctx, dir, 'skills'), detectPaths: [dir, joinClientPath(ctx, ctx.appData, 'DeepSeek Harness')] }; },
  },
  {
    id: 'kimi',
    name: 'Kimi',
    vendor: '月之暗面',
    region: 'cn',
    configKind: 'json-mcp',
    documentation: 'https://www.kimi.com/code/docs/en/kimi-code-desktop/settings-and-extensions.html',
    paths: ctx => { const dir = homeRoot(ctx, 'KIMI_CODE_HOME', '.kimi-code'); return { configPath: joinClientPath(ctx, dir, 'mcp.json'), skillsPath: joinClientPath(ctx, dir, 'skills'), detectPaths: [dir] }; },
  },
  {
    id: 'trae-cn',
    name: 'Trae CN',
    vendor: '字节跳动',
    region: 'cn',
    configKind: 'manual',
    manualHint: '在 TRAE 设置中添加 MCP，或使用项目配置。',
    documentation: 'https://docs.trae.ai/ide/model-context-protocol',
    paths: ctx => ({
      detectPaths: [joinClientPath(ctx, ctx.home, '.trae'), joinClientPath(ctx, ctx.home, '.trae-cn'), joinClientPath(ctx, ctx.appData, 'Trae CN')],
    }),
  },
  {
    id: 'lingma',
    name: '通义灵码',
    vendor: '阿里巴巴',
    region: 'cn',
    configKind: 'manual',
    manualHint: '在客户端 MCP 页面粘贴配置。',
    documentation: 'https://docs.qoder.com/zh/user-guide/chat/model-context-protocol',
    paths: ctx => ({
      detectPaths: [joinClientPath(ctx, ctx.home, '.lingma'), joinClientPath(ctx, ctx.home, '.qoder-cn')],
    }),
  },

  // ===== 国外 =====
  {
    id: 'claude-code',
    name: 'Claude Code',
    vendor: 'Anthropic',
    region: 'global',
    configKind: 'json-mcp',
    documentation: 'https://code.claude.com/docs/en/mcp',
    paths: ctx => ({
      configPath: joinClientPath(ctx, ctx.home, '.claude.json'),
      skillsPath: joinClientPath(ctx, ctx.home, '.claude/skills'),
      detectPaths: [joinClientPath(ctx, ctx.home, '.claude'), joinClientPath(ctx, ctx.home, '.claude.json')],
    }),
  },
  {
    id: 'codex',
    name: 'Codex',
    vendor: 'OpenAI',
    region: 'global',
    configKind: 'toml-codex',
    documentation: 'https://learn.chatgpt.com/docs/extend/mcp?surface=cli',
    paths: ctx => ({
      configPath: joinClientPath(ctx, homeRoot(ctx, 'CODEX_HOME', '.codex'), 'config.toml'),
      skillsPath: joinClientPath(ctx, ctx.home, '.agents/skills'),
      detectPaths: [homeRoot(ctx, 'CODEX_HOME', '.codex')],
    }),
  },
  {
    id: 'claude-desktop',
    name: 'Claude Desktop',
    vendor: 'Anthropic',
    region: 'global',
    configKind: 'json-mcp',
    documentation: 'https://py.sdk.modelcontextprotocol.io/get-started/real-host/',
    paths: ctx => ({
      configPath: joinClientPath(ctx, ctx.appData, 'Claude/claude_desktop_config.json'),
      detectPaths: [joinClientPath(ctx, ctx.appData, 'Claude')],
    }),
  },
  {
    id: 'cursor',
    name: 'Cursor',
    vendor: 'Anysphere',
    region: 'global',
    configKind: 'json-mcp',
    documentation: 'https://cursor.com/docs/mcp',
    paths: ctx => ({
      configPath: joinClientPath(ctx, ctx.home, '.cursor/mcp.json'),
      skillsPath: joinClientPath(ctx, ctx.home, '.cursor/skills'),
      detectPaths: [joinClientPath(ctx, ctx.home, '.cursor'), joinClientPath(ctx, ctx.appData, 'Cursor')],
    }),
  },
  {
    id: 'windsurf',
    name: 'Windsurf / Devin',
    vendor: 'Cognition',
    region: 'global',
    configKind: 'json-mcp',
    documentation: 'https://docs.devin.ai/cli/extensibility/mcp/configuration',
    paths: resolveWindsurf,
  },
  {
    id: 'gemini-cli',
    name: 'Gemini CLI',
    vendor: 'Google',
    region: 'global',
    configKind: 'json-mcp',
    documentation: 'https://geminicli.com/docs/tools/mcp-server/',
    paths: ctx => ({
      configPath: joinClientPath(ctx, ctx.home, '.gemini/settings.json'),
      enablementPath: joinClientPath(ctx, ctx.home, '.gemini/mcp-server-enablement.json'),
      skillsPath: joinClientPath(ctx, ctx.home, '.gemini/skills'),
      detectPaths: [joinClientPath(ctx, ctx.home, '.gemini')],
    }),
  },
  {
    id: 'antigravity',
    name: 'Antigravity',
    vendor: 'Google',
    region: 'global',
    configKind: 'json-mcp',
    documentation: 'https://antigravity.google/docs/mcp',
    paths: resolveAntigravity,
  },
  {
    id: 'opencode',
    name: 'opencode',
    vendor: 'SST',
    region: 'global',
    configKind: 'json-opencode',
    documentation: 'https://opencode.ai/docs/mcp-servers/',
    paths: resolveOpencode,
  },
  {
    id: 'vscode',
    name: 'VS Code',
    vendor: 'Microsoft',
    region: 'global',
    configKind: 'json-vscode',
    documentation: 'https://code.visualstudio.com/docs/agents/reference/mcp-configuration',
    paths: ctx => {
      const native = joinClientPath(ctx, ctx.appData, 'Code/User/mcp.json');
      const portableRoot = homeRoot(ctx, 'COPILOT_HOME', '.copilot'), portable = joinClientPath(ctx, portableRoot, 'mcp-config.json');
      // Preserve an existing official portable setup. Never write servers into mcpServers or vice versa.
      const targets: { file: string; kind: McpConfigKind }[] = [];
      if (fs.existsSync(native)) targets.push({ file: native, kind: 'json-vscode' });
      if (fs.existsSync(portable)) targets.push({ file: portable, kind: 'json-mcp' });
      if (!targets.length) targets.push({ file: native, kind: 'json-vscode' });
      return { configPath: targets[0].file, configPaths: targets.map(target => target.file), configTargets: targets, configKind: targets[0].kind, skillsPath: joinClientPath(ctx, ctx.home, '.agents/skills'), detectPaths: [portableRoot, joinClientPath(ctx, ctx.appData, 'Code')] };
    },
  },
];

export function findClient(id: string): ClientDefinition | undefined {
  return CLIENT_DEFINITIONS.find(client => client.id === id);
}
