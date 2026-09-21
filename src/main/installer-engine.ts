import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { atomicWrite, resourcePath, runtimeHome, runtimePort } from '../bridge/runtime.js';
import { AddonInstaller } from './addon-installer.js';

export function mergeMcpConfig(file: string, entry: unknown) {
  let config: any = {};
  if (fs.existsSync(file)) {
    const raw = fs.readFileSync(file, 'utf8');
    try { config = JSON.parse(raw); } catch { throw new Error('现有 JSON 无法解析，未覆盖。请先修复原配置。'); }
    if (!config || typeof config !== 'object' || Array.isArray(config)) throw new Error('现有配置不是 JSON 对象，未覆盖');
  }
  if (config.mcpServers !== undefined && (!config.mcpServers || typeof config.mcpServers !== 'object' || Array.isArray(config.mcpServers))) throw new Error('mcpServers 格式无效，未覆盖');
  config.mcpServers = { ...config.mcpServers, 'office-agent-bridge': entry, 'wps-bridge': entry };
  atomicWrite(file, JSON.stringify(config, null, 2) + '\n', true);
}
export class InstallerEngine {
  static runtimeEntry = '';
  static getSystemTargetDir() { return runtimeHome(); }
  static configEntry() {
    return { command: process.execPath, args: [this.runtimeEntry || resourcePath('dist/bridge/cli.cjs')], env: { ELECTRON_RUN_AS_NODE: '1', OFFICE_AGENT_BRIDGE_HOME: runtimeHome(), WPS_BRIDGE_HOME: runtimeHome(), WPS_BRIDGE_PORT: String(runtimePort()), WPS_BRIDGE_RESOURCES: resourcePath('package.json').replace(/[\\/]package\.json$/, '') } };
  }
  static detectEnvironment() {
    const home = os.homedir();
    const appData = process.platform === 'darwin' ? path.join(home, 'Library/Application Support') : process.env.APPDATA || path.join(home, 'AppData/Roaming');
    const entries = [
      ['doubao', '豆包桌面端', path.join(home, '.doubao/mcp.json'), path.join(home, '.doubao/skills')],
      ['workbuddy', 'WorkBuddy', path.join(home, '.workbuddy/mcp.json'), path.join(home, '.workbuddy/skills')],
      ['qwen', '千问办公', path.join(home, '.qwen/mcp.json'), path.join(home, '.qwen/skills')],
      ['kimi', 'Kimi', path.join(home, '.kimi-code/mcp.json'), path.join(home, '.kimi-code/skills')],
      ['claude-code', 'Claude Code', path.join(home, '.claude.json'), path.join(home, '.claude/skills')],
      ['codex', 'Codex', path.join(home, '.codex/mcp.json'), path.join(home, '.codex/skills')]
    ];
    const agents = entries.map(([id, name, configPath, skillsPath]) => {
      let configured = false, detail = '';
      try {
        if (fs.existsSync(configPath)) {
          const srv = JSON.parse(fs.readFileSync(configPath, 'utf8')).mcpServers;
          configured = Boolean(srv?.['office-agent-bridge'] || srv?.['wps-bridge']);
        }
      } catch { detail = '配置 JSON 无法解析，修复前不会覆盖'; }
      const detected = fs.existsSync(path.dirname(configPath));
      return { id, name, configPath, skillsPath, detected, status: configured ? 'configured' : detected ? 'installed' : 'not_found', detail };
    });
    return { platform: process.platform, systemTargetDir: runtimeHome(), addon: AddonInstaller.checkStatus(), agents };
  }
  static async executeInstall(options: { agents?: string[]; addon?: boolean; skills?: boolean } = {}) {
    const logs: { step: string; status: string; detail: string }[] = [];
    if (options.addon) { const r = AddonInstaller.install(); logs.push({ step: 'WPS 加载项', status: r.success ? 'success' : 'error', detail: r.message }); }
    for (const agent of this.detectEnvironment().agents.filter(a => options.agents?.includes(a.id))) {
      try {
        mergeMcpConfig(agent.configPath, this.configEntry());
        if (options.skills && agent.skillsPath) {
          const source = resourcePath('skills');
          for (const name of fs.readdirSync(source)) {
            const dir = path.join(source, name);
            if (fs.existsSync(path.join(dir, 'SKILL.md'))) this.copySkills(dir, path.join(agent.skillsPath, name));
          }
        }
        logs.push({ step: agent.name, status: 'success', detail: '已合并 MCP 配置；支持技能目录的客户端已按选择同步技能。' });
      } catch (e: any) { logs.push({ step: agent.name, status: 'error', detail: e.message }); }
    }
    return { success: logs.every(l => l.status === 'success'), message: logs.length ? '所选配置处理完成，请查看每项结果。' : '未选择需要安装的项目。', logs };
  }
  private static copySkills(source: string, dest: string) {
    for (const entry of fs.readdirSync(source, { withFileTypes: true })) {
      if (entry.isSymbolicLink()) continue;
      const from = path.join(source, entry.name), to = path.join(dest, entry.name);
      if (entry.isDirectory()) this.copySkills(from, to); else atomicWrite(to, fs.readFileSync(from, 'utf8'), true);
    }
  }
}
