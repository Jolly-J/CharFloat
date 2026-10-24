import fs from 'node:fs';
import path from 'node:path';
import { resourcePath, runtimeHome, runtimePort } from '../bridge/runtime.js';
import { AddonInstaller } from './addon-installer.js';
import { CLIENT_DEFINITIONS, clientPathContext, findClient, type ClientConfigKind, type ClientRuntimePaths } from './client-registry.js';
import { applyMcpEntry, inspectMcpConfig, inspectMcpEnablement, type McpServerEntry, type McpConfigInspection, type McpConfigKind } from './mcp-adapters.js';
import { inspectSkillLinks, linkClientSkills, syncManagedSkills, type SkillLinkInspection } from './skill-links.js';

// 兼容既有调用方与测试：JSON mcpServers 合并仍从本模块导出
export { mergeMcpConfig } from './mcp-adapters.js';

export interface AgentRuntimeInfo {
  id: string;
  name: string;
  vendor: string;
  region: 'cn' | 'global';
  configPath: string | undefined;
  configPaths: string[];
  configTargets: { file: string; kind: McpConfigKind }[];
  enablementPath?: string;
  skillsPath: string | undefined;
  skillsPaths: string[];
  configKind: ClientConfigKind;
  detected: boolean;
  status: 'configured' | 'installed' | 'not_found';
  detail: string;
  recommended: boolean;
  manual: boolean;
  manualHint?: string;
  documentation?: string;
  mcp: McpConfigInspection | { status: 'manual'; detail: string };
  skills: SkillLinkInspection;
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
    const ctx = clientPathContext(), expected = this.configEntry();
    const agents: AgentRuntimeInfo[] = CLIENT_DEFINITIONS.map(def => {
      let paths: ClientRuntimePaths, pathError = '';
      try { paths = def.paths(ctx); } catch (error: any) { paths = { detectPaths: [] }; pathError = error.message || '无法确定客户端目录'; }
      const { configPath, configPaths, skillsPath, detectPaths } = paths;
      const targets = configPaths?.length ? configPaths : configPath ? [configPath] : [];
      const detected = detectPaths.some(p => fs.existsSync(p));
      const configKind = paths.configKind || def.configKind;
      const manual = configKind === 'manual' || Boolean(pathError);
      const configTargets = paths.configTargets || (configKind === 'manual' ? [] : targets.map(file => ({ file, kind: configKind })));
      const inspections = configTargets.map(target => inspectMcpEnablement(paths.enablementPath, inspectMcpConfig(target.kind, target.file, expected)));
      const mcp: AgentRuntimeInfo['mcp'] = pathError ? { status: 'invalid', detail: pathError } : configKind === 'manual' ? { status: 'manual', detail: def.manualHint || '在客户端中添加连接' }
        : inspections.find(value => value.status !== 'valid') || inspections[0] || { status: 'missing', detail: '尚未配置' };
      const configured = mcp.status === 'valid';
      const skillTargets = paths.skillsPaths?.length ? paths.skillsPaths : skillsPath ? [skillsPath] : [];
      const skillChecks = manual ? [] : skillTargets.map(dir => inspectSkillLinks(dir));
      const skills = skillChecks.find(value => value.status !== 'linked') || skillChecks[0] || inspectSkillLinks();
      return {
        id: def.id,
        name: def.name,
        vendor: def.vendor,
        region: def.region,
        configPath,
        configPaths: targets,
        configTargets,
        enablementPath: paths.enablementPath,
        skillsPath,
        skillsPaths: skillTargets,
        configKind,
        detected,
        status: configured ? 'configured' : detected ? 'installed' : 'not_found',
        detail: mcp.detail,
        recommended: Boolean(def.recommended),
        manual,
        manualHint: pathError || def.manualHint,
        documentation: def.documentation,
        mcp,
        skills,
      };
    });
    return { platform: process.platform, systemTargetDir: runtimeHome(), addon: AddonInstaller.checkStatus(), agents };
  }

  static async executeInstall(options: { agents?: string[]; addon?: boolean; skills?: boolean } = {}) {
    const logs: { step: string; status: string; detail: string }[] = [];
    if (options.addon) { const r = AddonInstaller.install(); logs.push({ step: 'WPS 加载项', status: r.success ? 'success' : 'error', detail: r.message }); }
    let skillSource: ReturnType<typeof syncManagedSkills> | undefined;
    for (const agent of this.detectEnvironment().agents.filter(a => options.agents?.includes(a.id))) {
      const def = findClient(agent.id);
      if (!def) continue;
      const configKind = agent.configKind;
      if (configKind === 'manual' || agent.configPaths.length === 0) {
        logs.push({ step: agent.name, status: 'manual', detail: agent.manualHint || `${agent.name} 的 MCP 配置只能在客户端应用内添加，本次未写入文件。` });
        continue;
      }
      let writtenTargets = 0;
      try {
        const shouldSyncSkills = options.skills !== false;
        if (shouldSyncSkills && agent.skillsPaths.length) skillSource ||= syncManagedSkills();
        const entry = this.configEntry();
        let disabled = false;
        for (const target of agent.configTargets) {
          applyMcpEntry(target.kind, target.file, entry);
          writtenTargets++;
          const verified = inspectMcpEnablement(agent.enablementPath, inspectMcpConfig(target.kind, target.file, entry));
          if (!['valid', 'disabled'].includes(verified.status)) throw new Error('配置写入后未通过读回检查，请保留备份并重试。');
          disabled ||= verified.status === 'disabled';
        }
        let linked = 0, backups = 0;
        if (shouldSyncSkills && skillSource) for (const destination of agent.skillsPaths) {
          const result = linkClientSkills(agent.id, destination, skillSource); linked += result.count; backups += result.backups;
        }
        logs.push({
          step: agent.name,
          status: disabled ? 'manual' : 'success',
          detail: `MCP 配置已保存并读回${linked ? `，${linked} 项技能已链接` : ''}${backups ? '，原技能副本已备份' : ''}。${disabled ? '客户端原有停用设置已保留，请在客户端启用。' : '请在客户端刷新连接；其授权设置保持不变。'}`,
        });
      } catch (e: any) { logs.push({ step: agent.name, status: 'error', detail: `${writtenTargets ? '部分配置已写入；' : ''}${e.message}` }); }
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

}
