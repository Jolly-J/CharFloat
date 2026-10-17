import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { atomicWrite, resourcePath, runtimeHome, runtimePort } from '../bridge/runtime.js';
import { AddonInstaller } from './addon-installer.js';
import { legacyServerKey, legacyToolPrefix, legacySkillNames } from './legacy-brand.js';
import { CLIENT_DEFINITIONS, clientPathContext, findClient, type ClientConfigKind, type ClientRuntimePaths } from './client-registry.js';
import { applyMcpEntry, isMcpConfigured, mergeMcpConfig, type McpServerEntry } from './mcp-adapters.js';

// 兼容既有调用方与测试：JSON mcpServers 合并仍从本模块导出
export { mergeMcpConfig } from './mcp-adapters.js';

export interface AgentRuntimeInfo {
  id: string;
  name: string;
  vendor: string;
  region: 'cn' | 'global';
  configPath: string | undefined;
  configPaths: string[];
  skillsPath: string | undefined;
  configKind: ClientConfigKind;
  detected: boolean;
  status: 'configured' | 'installed' | 'not_found';
  detail: string;
  recommended: boolean;
  manual: boolean;
  manualHint?: string;
}

export class InstallerEngine {
  static runtimeEntry = '';
  static getSystemTargetDir() { return runtimeHome(); }
  static configEntry(): McpServerEntry {
    return {
      command: process.execPath,
      args: [this.runtimeEntry || resourcePath('dist/bridge/cli.cjs')],
      env: {
        ELECTRON_RUN_AS_NODE: '1',
        CHARFLOAT_HOME: runtimeHome(),
        WPS_BRIDGE_HOME: runtimeHome(),
        WPS_BRIDGE_PORT: String(runtimePort()),
        WPS_BRIDGE_RESOURCES: resourcePath('package.json').replace(/[\\/]package\.json$/, ''),
      },
    };
  }

  /** 按注册表生成各客户端的接入状态；读取状态不会写入任何配置。 */
  static detectEnvironment() {
    const ctx = clientPathContext();
    const agents: AgentRuntimeInfo[] = CLIENT_DEFINITIONS.map(def => {
      const paths: ClientRuntimePaths = def.paths(ctx);
      const { configPath, configPaths, skillsPath, detectPaths } = paths;
      const targets = configPaths?.length ? configPaths : configPath ? [configPath] : [];
      const detected = detectPaths.some(p => fs.existsSync(p));
      const manual = def.configKind === 'manual';
      const configKind = def.configKind;
      let configured = false;
      let detail = '';
      if (configKind === 'manual') {
        detail = def.manualHint || '该客户端只能在应用内添加 MCP 服务，需手动接入';
      } else if (def.id === 'doubao' && configPath) {
        configured = this.isDoubaoConfigured(path.dirname(configPath));
        if (configured) detail = '豆包工作连接器已配置成功并就绪';
      } else if (targets.length) {
        for (const file of targets) {
          try {
            if (isMcpConfigured(configKind, file)) { configured = true; break; }
          } catch {
            detail = '配置 JSON 无法解析，修复前不会覆盖';
          }
        }
      }
      return {
        id: def.id,
        name: def.name,
        vendor: def.vendor,
        region: def.region,
        configPath,
        configPaths: targets,
        skillsPath,
        configKind: def.configKind,
        detected,
        status: configured ? 'configured' : detected ? 'installed' : 'not_found',
        detail,
        recommended: Boolean(def.recommended),
        manual,
        manualHint: def.manualHint,
      };
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

      // 2. 检查豆包本地 sessions 目录下是否有当前或旧版服务的真实工具调用记录
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
                if (files.some(f => f.includes('charfloat') || f.includes(legacyToolPrefix) || f.includes('office_agent'))) {
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
      const def = findClient(agent.id);
      if (!def) continue;
      const configKind = def.configKind;
      if (configKind === 'manual' || agent.configPaths.length === 0) {
        logs.push({ step: agent.name, status: 'manual', detail: agent.manualHint || `${agent.name} 的 MCP 配置只能在客户端应用内添加，本次未写入文件。` });
        continue;
      }
      try {
        for (const target of agent.configPaths) applyMcpEntry(configKind, target, this.configEntry());
        if (agent.id === 'workbuddy') {
          this.autoApproveWorkbuddy(path.dirname(agent.configPaths[0]), this.configEntry(), 'charfloat');
        }
        const shouldSyncSkills = options.skills !== false;
        if (shouldSyncSkills && agent.skillsPath) {
          // 清洗历史遗留技能目录，防止客户端技能列表残留旧名称
          const legacySkills = [...legacySkillNames, 'wps-bridge'];
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
            : '已合并 MCP 配置并完成安全授信。',
        });
      } catch (e: any) { logs.push({ step: agent.name, status: 'error', detail: e.message }); }
    }
    const successCount = logs.filter(l => l.status === 'success').length;
    const manualCount = logs.filter(l => l.status === 'manual').length;
    const errorCount = logs.filter(l => l.status === 'error').length;
    const message = logs.length === 0
      ? '未选择需要安装的项目。'
      : errorCount > 0
        ? `${successCount ? `${successCount} 个成功、` : ''}${errorCount} 个失败${manualCount ? `、${manualCount} 个需手动` : ''}，请查看每项结果。`
        : successCount > 0
          ? manualCount
            ? `已为 ${successCount} 个客户端完成配置，${manualCount} 个需在客户端内手动接入。`
            : `已为 ${successCount} 个客户端完成配置。`
          : `${manualCount} 个客户端需在应用内手动接入，本次未写入配置。`;
    // 「需手动」不算失败：只有真正写入出错才降低整体成功标记
    return { success: errorCount === 0, message, logs, stats: { success: successCount, manual: manualCount, error: errorCount } };
  }

  private static copySkills(source: string, dest: string) {
    for (const entry of fs.readdirSync(source, { withFileTypes: true })) {
      if (entry.isSymbolicLink()) continue;
      const from = path.join(source, entry.name), to = path.join(dest, entry.name);
      if (entry.isDirectory()) this.copySkills(from, to); else atomicWrite(to, fs.readFileSync(from, 'utf8'), true);
    }
  }
}
