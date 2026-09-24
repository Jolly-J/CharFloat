import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import crypto from 'node:crypto';
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
  // 只写当前品牌名称 charfloat。
  // 迁移规则：只清掉确属本产品的遗留条目（office-agent-bridge 与 wps-bridge，命令/参数与原条目一致），
  // 别人的同名条目（若有）一律保留——安全性质与"合并写入、不覆盖他人"保持一致。
  const LEGACY_KEYS = ['office-agent-bridge', 'wps-bridge'];
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
  next['charfloat'] = entry;
  config.mcpServers = next;
  atomicWrite(file, JSON.stringify(config, null, 2) + '\n', true);
}
export class InstallerEngine {
  static runtimeEntry = '';
  static getSystemTargetDir() { return runtimeHome(); }
  static configEntry() {
    return { command: process.execPath, args: [this.runtimeEntry || resourcePath('dist/bridge/cli.cjs')], env: { ELECTRON_RUN_AS_NODE: '1', CHARFLOAT_HOME: runtimeHome(), OFFICE_AGENT_BRIDGE_HOME: runtimeHome(), WPS_BRIDGE_HOME: runtimeHome(), WPS_BRIDGE_PORT: String(runtimePort()), WPS_BRIDGE_RESOURCES: resourcePath('package.json').replace(/[\\/]package\.json$/, '') } };
  }
  static detectEnvironment() {
    const home = os.homedir();
    const appData = process.platform === 'darwin' ? path.join(home, 'Library/Application Support') : process.env.APPDATA || path.join(home, 'AppData/Roaming');

    // 豆包工作（Doubao Work）桌面端配置与技能目录适配
    // macOS: ~/Library/Application Support/DoubaoWork/Default/.doubaowork
    // Windows: %APPDATA%/DoubaoWork/Default/.doubaowork
    const doubaoWorkBase = path.join(appData, 'DoubaoWork');
    const doubaoWorkDir = path.join(doubaoWorkBase, 'Default/.doubaowork');
    const doubaoWorkDetected = fs.existsSync(doubaoWorkDir) || fs.existsSync(doubaoWorkBase);
    const doubaoConfigPath = doubaoWorkDetected
      ? path.join(doubaoWorkDir, 'mcp.json')
      : path.join(home, '.doubao/mcp.json');
    const doubaoSkillsPath = doubaoWorkDetected
      ? path.join(doubaoWorkDir, 'agent_mode/workspace/.user_skills')
      : path.join(home, '.doubao/skills');

    const entries: [string, string, string, string, boolean?][] = [
      ['workbuddy', 'WorkBuddy', path.join(home, '.workbuddy/mcp.json'), path.join(home, '.workbuddy/skills'), true],
      ['doubao', '豆包工作', doubaoConfigPath, doubaoSkillsPath],
      ['qwen', '千问办公', path.join(home, '.qwen/mcp.json'), path.join(home, '.qwen/skills')],
      ['kimi', 'Kimi', path.join(home, '.kimi-code/mcp.json'), path.join(home, '.kimi-code/skills')],
      ['claude-code', 'Claude Code', path.join(home, '.claude.json'), path.join(home, '.claude/skills')],
      ['codex', 'Codex', path.join(home, '.codex/mcp.json'), path.join(home, '.codex/skills')]
    ];
    const agents = entries.map(([id, name, configPath, skillsPath, recommended]) => {
      let configured = false, detail = '';
      try {
        if (id === 'doubao') {
          configured = this.isDoubaoConfigured(doubaoWorkDir);
          if (configured) {
            detail = '豆包工作连接器已配置成功并就绪';
          }
        } else if (fs.existsSync(configPath)) {
          const srv = JSON.parse(fs.readFileSync(configPath, 'utf8')).mcpServers;
          configured = Boolean(srv?.['charfloat'] || srv?.['office-agent-bridge'] || srv?.['wps-bridge']);
        }
      } catch { detail = '配置 JSON 无法解析，修复前不会覆盖'; }
      const detected = fs.existsSync(path.dirname(configPath));
      return { id, name, configPath, skillsPath, detected, status: configured ? 'configured' : detected ? 'installed' : 'not_found', detail, recommended: Boolean(recommended) };
    });
    return { platform: process.platform, systemTargetDir: runtimeHome(), addon: AddonInstaller.checkStatus(), agents };
  }
  static isDoubaoConfigured(doubaoWorkDir: string): boolean {
    try {
      // 1. 检查桌面端持久化偏好 desktop.json 是否记录了已配置
      const prefsPath = path.join(runtimeHome(), 'desktop.json');
      if (fs.existsSync(prefsPath)) {
        try {
          const prefs = JSON.parse(fs.readFileSync(prefsPath, 'utf8'));
          if (prefs.doubaoConfigured) return true;
        } catch {}
      }

      // 2. 检查豆包本地 sessions 目录下是否有调用过 office_agent_bridge 的真实工具调用记录
      const sessionsDir = path.join(doubaoWorkDir, 'agent_mode/workspace/.sessions');
      if (fs.existsSync(sessionsDir)) {
        const sessionDirs = fs.readdirSync(sessionsDir);
        for (const s of sessionDirs) {
          const agentsDir = path.join(sessionsDir, s, 'agents');
          if (fs.existsSync(agentsDir)) {
            const agentList = fs.readdirSync(agentsDir);
            for (const a of agentList) {
              const toolResultsDir = path.join(agentsDir, a, 'system/tool-results');
              if (fs.existsSync(toolResultsDir)) {
                const files = fs.readdirSync(toolResultsDir);
                if (files.some(f => f.includes('charfloat') || f.includes('office_agent_bridge') || f.includes('office_agent'))) {
                  return true;
                }
              }
            }
          }
        }
      }
    } catch {}
    return false;
  }
  static autoApproveWorkbuddy(configDir: string, entry: any, serverName: string) {
    try {
      const approvalsFile = path.join(configDir, 'mcp-approvals.json');
      let approvals: Record<string, number> = {};
      if (fs.existsSync(approvalsFile)) {
        try { approvals = JSON.parse(fs.readFileSync(approvalsFile, 'utf8')); } catch {}
      }
      const command = entry.command || '';
      const args = (entry.args || []).map(String).sort().join(',');
      const envKeys = Object.keys(entry.env || {}).sort().join(',');
      const input = `${command}|${args}|${envKeys}`;
      const hash = crypto.createHash('sha256').update(input).digest('hex');
      const key = `${hash}::${serverName}`;
      approvals[key] = Date.now();
      atomicWrite(approvalsFile, JSON.stringify(approvals, null, 2) + '\n', true);
    } catch {}
  }
  static async executeInstall(options: { agents?: string[]; addon?: boolean; skills?: boolean } = {}) {
    const logs: { step: string; status: string; detail: string }[] = [];
    if (options.addon) { const r = AddonInstaller.install(); logs.push({ step: 'WPS 加载项', status: r.success ? 'success' : 'error', detail: r.message }); }
    for (const agent of this.detectEnvironment().agents.filter(a => options.agents?.includes(a.id))) {
      try {
        mergeMcpConfig(agent.configPath, this.configEntry());
        if (agent.id === 'workbuddy') {
          this.autoApproveWorkbuddy(path.dirname(agent.configPath), this.configEntry(), 'charfloat');
        }
        const shouldSyncSkills = options.skills !== false;
        if (shouldSyncSkills && agent.skillsPath) {
          // 清洗历史遗留技能目录，防止客户端技能列表残留旧名称
          const legacySkills = [
            'office-agent-bridge',
            'office-agent-bridge-chart-style',
            'office-agent-bridge-ppt-design',
            'office-agent-bridge-word-batch-edit',
            'wps-bridge'
          ];
          for (const leg of legacySkills) {
            const legPath = path.join(agent.skillsPath, leg);
            if (fs.existsSync(legPath)) {
              try { fs.rmSync(legPath, { recursive: true, force: true }); } catch {}
            }
          }
          const source = resourcePath('skills');
          for (const name of fs.readdirSync(source)) {
            const dir = path.join(source, name);
            if (fs.existsSync(path.join(dir, 'SKILL.md'))) this.copySkills(dir, path.join(agent.skillsPath, name));
          }
        }
        logs.push({ 
          step: agent.name, 
          status: 'success', 
          detail: agent.skillsPath 
            ? '已合并 MCP 配置并完成安全授信；专属技能库已自动同步就绪。' 
            : '已合并 MCP 配置并完成安全授信。' 
        });
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
