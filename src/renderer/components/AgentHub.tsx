import React, { useEffect, useRef, useState } from 'react';
import { 
  Cable, 
  CheckCircle2, 
  AlertCircle, 
  Copy, 
  ChevronDown, 
  LoaderCircle,
  Check,
  Bot,
  X,
  ExternalLink
} from 'lucide-react';
import appLogo from '../assets/app-logo.png';
import brandGreeting from '../assets/brand-greeting.png';
import doubaoLogo from '../assets/doubao-work.png';
import workbuddyLogo from '../assets/workbuddy.svg';
import qwenLogo from '../assets/qwen.png';
import kimiLogo from '../assets/kimi.webp';
import claudeLogo from '../assets/claude-color.svg';
import claudeDesktopLogo from '../assets/claude-desktop.png';
import codexLogo from '../assets/codex.png';
import traeLogo from '../assets/trae.png';
import lingmaLogo from '../assets/lingma.png';
import cursorLogo from '../assets/cursor.png';
import windsurfLogo from '../assets/windsurf.svg';
import geminiLogo from '../assets/gemini.png';
import antigravityLogo from '../assets/antigravity.png';
import opencodeLogo from '../assets/opencode.png';
import vscodeLogo from '../assets/vscode.png';
import { AnimatePresence, motion, useReducedMotion } from 'motion/react';

const agentLogos: Record<string, string> = {
  // 国内
  doubao: doubaoLogo,
  workbuddy: workbuddyLogo,
  qwen: qwenLogo,
  kimi: kimiLogo,
  'trae-cn': traeLogo,
  lingma: lingmaLogo,
  // 国外
  'claude-code': claudeLogo,
  codex: codexLogo,
  'claude-desktop': claudeDesktopLogo,
  cursor: cursorLogo,
  windsurf: windsurfLogo,
  'gemini-cli': geminiLogo,
  antigravity: antigravityLogo,
  opencode: opencodeLogo,
  vscode: vscodeLogo,
};

const REGIONS = [
  { id: 'cn' as const, label: '国内' },
  { id: 'global' as const, label: '国外' },
];

interface AgentHubProps {
  agents: any[];
  selectedAgents: string[];
  busy: string;
  installLogs: any[];
  mcpConfig: any;
  onToggleAgent: (id: string) => void;
  onToggleSelectAll: () => void;
  onExecuteInstall: () => void;
  onCopyText: (text: string) => void;
  onOpenDoubaoGuide?: () => void;
  onOpenAgentApp?: (id: string) => void;
  onOpenDocumentation?: (id: string) => void;
}

export const AgentHub: React.FC<AgentHubProps> = ({
  agents = [],
  selectedAgents,
  busy,
  installLogs = [],
  mcpConfig,
  onToggleAgent,
  onToggleSelectAll,
  onExecuteInstall,
  onCopyText,
  onOpenDoubaoGuide,
  onOpenAgentApp,
  onOpenDocumentation,
}) => {
  const reduced = useReducedMotion();
  const [manualAgent, setManualAgent] = useState<any>(null);
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [copiedType, setCopiedType] = useState<'prompt' | 'json' | null>(null);
  const [activeRegion, setActiveRegion] = useState<'cn' | 'global'>('cn');
  const regionDefaultedRef = useRef(false);
  const isInstalling = busy === 'install';
  const logStats = installLogs.reduce(
    (acc: { success: number; manual: number; error: number }, log: any) => {
      if (log?.status === 'success') acc.success += 1;
      else if (log?.status === 'manual') acc.manual += 1;
      else acc.error += 1;
      return acc;
    },
    { success: 0, manual: 0, error: 0 },
  );

  const regionOf = (agent: any): 'cn' | 'global' => (agent?.region === 'cn' ? 'cn' : 'global');
  const regionAgents = (region: 'cn' | 'global') => agents.filter((agent: any) => regionOf(agent) === region);
  const visibleAgents = regionAgents(activeRegion);

  // 默认停在「国内」；只有国内一个都没配置、国外已有配置时才自动切到国外。用户手动切换后不再自动跳转
  useEffect(() => {
    if (regionDefaultedRef.current || agents.length === 0) return;
    regionDefaultedRef.current = true;
    const configuredCount = (region: 'cn' | 'global') =>
      regionAgents(region).filter((agent: any) => agent.status === 'configured').length;
    if (configuredCount('cn') === 0 && configuredCount('global') > 0) setActiveRegion('global');
  }, [agents]);

  const getAiPrompt = () => {
    const jsonStr = JSON.stringify(mcpConfig, null, 2);
    return `请帮我把本地「字浮 CharFloat」办公自动化服务配置到你（当前 AI 助手）的 MCP (Model Context Protocol) 列表中。

stdio 服务参数如下，请按当前客户端官方配置格式合并，不要覆盖其他服务：
\`\`\`json
${jsonStr}
\`\`\`

配置执行说明：
1. 核对你所在客户端的官方位置和格式。Codex 使用 TOML，DeepSeek Harness 使用 Cordis YAML，VS Code 使用 servers；不能全部按 mcpServers JSON 写入。
2. 保留其他服务、自定义选项、禁用和授权设置，解析失败不要覆盖。
3. 写入后读回核对命令、参数和路径；分别说明配置已保存与客户端实际已连接，不能混为一谈。`;
  };

  const handleCopyPrompt = () => {
    onCopyText(getAiPrompt());
    setCopiedType('prompt');
    setTimeout(() => setCopiedType(null), 2000);
  };

  const handleCopyJson = () => {
    onCopyText(JSON.stringify(mcpConfig, null, 2));
    setCopiedType('json');
    setTimeout(() => setCopiedType(null), 2000);
  };

  return (
    <><section className="agent-hub-card">
      {/* 头部：标题与快捷操作栏 */}
      <div className="agent-hub-header">
        <div className="hub-title-group">
          <div className="hub-icon-wrap">
              <img src={brandGreeting} alt="" className="hub-greeting-img" />
          </div>
          <div>
            <h3>AI 接入</h3>
          </div>
        </div>

        <div className="hub-actions-bar">
          <button
            className="install-cta-btn"
            disabled={Boolean(busy) || selectedAgents.length === 0}
            onClick={onExecuteInstall}
          >
            {isInstalling ? <LoaderCircle size={13} className="spin" /> : <Cable size={13} />}
            <span>配置所选 ({selectedAgents.length})</span>
          </button>
        </div>
      </div>

      {/* 分类切换器：国内 / 国外 */}
      <div className="agent-region-bar">
        <div className="agent-region-switch" role="tablist" aria-label="AI 客户端分类">
          {REGIONS.map(region => {
            const list = regionAgents(region.id);
            const configured = list.filter((agent: any) => agent.status === 'configured').length;
            return (
              <button
                key={region.id}
                type="button"
                role="tab"
                aria-selected={activeRegion === region.id}
                className={`region-tab ${activeRegion === region.id ? 'active' : ''}`}
                title={`${region.label}客户端 ${list.length} 个，已配置 ${configured} 个`}
                onClick={() => setActiveRegion(region.id)}
              >
                <span>{region.label}</span>
                <span className="region-count">{list.length}</span>
              </button>
            );
          })}
        </div>
      </div>

      {/* AI 客户端网格（按分类切换） */}
      <div className="agent-grid">
        {visibleAgents.map((agent) => {
          const isSelected = selectedAgents.includes(agent.id);
          const isConfigured = agent.status === 'configured';
          const isDetected = agent.detected;
          const logo = agentLogos[agent.id];

          return (
            <div
              key={agent.id}
              className={`agent-card ${isSelected ? 'selected' : ''}`}
              onClick={() => onToggleAgent(agent.id)}
            >
              <div className="agent-card-top">
                <div className="agent-top-left">
                  <div className="agent-logo-box">
                    {logo ? (
                      <img src={logo} alt={agent.name} className="agent-logo-img" />
                    ) : (
                      <span className={`agent-logo-fallback ${agent.id === 'deepseek-harness' ? 'deepseek-mark' : ''}`}>{agent.id === 'deepseek-harness' ? 'DS' : agent.name[0]}</span>
                    )}
                  </div>
                  {agent.id === 'workbuddy' && onOpenAgentApp && (
                    <button 
                      type="button" 
                      className="doubao-guide-trigger"
                      onClick={(e) => {
                        e.stopPropagation();
                        onOpenAgentApp('workbuddy');
                      }}
                      title="呼出 WorkBuddy 的 MCP 服务管理界面"
                    >
                      <span>打开管理</span>
                    </button>
                  )}
                  {agent.id === 'doubao' && (
                    <div style={{ display: 'flex', gap: '4px', alignItems: 'center' }}>
                      {onOpenDoubaoGuide && (
                        <button 
                          type="button" 
                          className="doubao-guide-trigger"
                          onClick={(e) => {
                            e.stopPropagation();
                            onOpenDoubaoGuide();
                          }}
                          title="点击查看豆包工作连接器添加指引"
                        >
                          <span>连接器指引</span>
                        </button>
                      )}
                      {onOpenAgentApp && (
                        <button 
                          type="button" 
                          className="doubao-guide-trigger"
                          onClick={(e) => {
                            e.stopPropagation();
                            onOpenAgentApp('doubao');
                          }}
                          title="呼出豆包工作桌面端"
                        >
                          <span>呼出软件</span>
                        </button>
                      )}
                    </div>
                  )}
                </div>

                <div className="agent-card-checkbox">
                  <input
                    type="checkbox"
                    checked={isSelected}
                    onChange={() => {}}
                    aria-label={`选择 ${agent.name}`}
                  />
                </div>
              </div>

              <div className="agent-card-info">
                <div className="agent-name-line">
                  <div className="agent-title-group">
                    <span className="agent-card-title">{agent.name}</span>
                    {agent.recommended && (
                      <span className="agent-badge recommended">
                        <span>推荐</span>
                      </span>
                    )}
                  </div>
                  {/* 状态徽章：严格单行并排 */}
                  {isConfigured ? (
                    <span className="agent-badge active">
                      <CheckCircle2 size={10} />
                      <span>已保存</span>
                    </span>
                  ) : agent.manual ? (
                    <span className="agent-badge gray">
                      <span>应用内连接</span>
                    </span>
                  ) : isDetected ? (
                    <span className="agent-badge pending">
                      <span>{agent.mcp?.status === 'disabled' ? '已停用' : agent.mcp?.status === 'outdated' ? '待更新' : agent.mcp?.status === 'invalid' ? '待修复' : '待配置'}</span>
                    </span>
                  ) : (
                    <span className="agent-badge gray">
                      <span>未检测到</span>
                    </span>
                  )}
                </div>
                <div className="agent-config-footnote" title={`${agent.detail || ''}\n${agent.skills?.detail || ''}`}>
                  <span>{agent.manual ? '在客户端完成接入' : agent.skills?.status === 'linked' ? '技能已链接' : agent.skills?.status === 'copied' ? '技能为旧副本' : isConfigured ? '在客户端刷新连接' : '选择后配置'}</span>
                  {agent.manual ? <button onClick={event => { event.stopPropagation(); agent.id === 'doubao' && onOpenDoubaoGuide ? onOpenDoubaoGuide() : setManualAgent(agent); }}>连接<ChevronDown size={11} /></button>
                    : agent.documentation && <button aria-label={`${agent.name}官方接入说明`} onClick={event => { event.stopPropagation(); onOpenDocumentation?.(agent.id); }}><ExternalLink size={12} /></button>}
                </div>
              </div>
            </div>
          );
        })}
        {visibleAgents.length === 0 && (
          <div className="agent-grid-empty">该分类下暂无客户端</div>
        )}
      </div>

      {/* 安装反馈结果面板 */}
      {installLogs.length > 0 && (
        <div className="install-feedback-panel">
          <div className="feedback-title">
            {logStats.error > 0 ? (
              <AlertCircle size={13} className="text-rose-500" />
            ) : logStats.manual > 0 ? (
              <AlertCircle size={13} className="text-amber-500" />
            ) : (
              <CheckCircle2 size={13} className="text-emerald-500" />
            )}
            <span>配置更新结果</span>
            <span className="feedback-summary">
              {logStats.success > 0 && <span className="ok">成功 {logStats.success}</span>}
              {logStats.manual > 0 && <span className="manual">需手动 {logStats.manual}</span>}
              {logStats.error > 0 && <span className="fail">失败 {logStats.error}</span>}
            </span>
          </div>
          <div className="feedback-list">
            {installLogs.map((log, idx) => (
              <div className={`feedback-item ${log.status === 'error' ? 'error' : ''}`} key={idx}>
                {log.status === 'success' ? (
                  <CheckCircle2 size={12} className="text-emerald-500" />
                ) : (
                  <AlertCircle size={12} className="text-amber-500" />
                )}
                <span>
                  <strong>{log.step}</strong>: {log.detail}
                  {log.step === '豆包工作' && onOpenDoubaoGuide && (
                    <button 
                      type="button" 
                      className="doubao-guide-trigger" 
                      style={{ marginLeft: '8px', padding: '1px 6px', fontSize: '9px', verticalAlign: 'middle' }}
                      onClick={onOpenDoubaoGuide}
                    >
                      <span>连接器指引</span>
                    </button>
                  )}
                </span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 让 AI 助手自动配置自身（复制 Prompt 给 Cursor / Windsurf 等） */}
      <div className="advanced-mcp-fold">
        <button
          type="button"
          className="fold-toggle-btn"
          onClick={() => setShowAdvanced(!showAdvanced)}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <Bot size={13} className="text-muted" />
            <span>手动接入</span>
          </div>
          <ChevronDown size={13} className={showAdvanced ? 'rotate-180' : ''} />
        </button>

        {showAdvanced && (
          <div className="fold-content" style={{ marginTop: '10px' }}>
            <div className="mcp-prompt-guide-box">
              <div className="mcp-prompt-header">
                <div>
                  <span className="mcp-prompt-tip-title">交给 AI 配置</span>
                  <p className="sub-caption" style={{ margin: '2px 0 0' }}>
                    复制指令发给当前客户端。
                  </p>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 }}>
                  <button
                    type="button"
                    className="copy-prompt-cta-btn"
                    onClick={handleCopyPrompt}
                  >
                    {copiedType === 'prompt' ? (
                      <>
                        <Check size={12} className="text-emerald-500" />
                        <span>已复制</span>
                      </>
                    ) : (
                      <>
                        <Copy size={12} />
                        <span>复制接入指令</span>
                      </>
                    )}
                  </button>

                  <button
                    type="button"
                    className="copy-json-mini-btn"
                    onClick={handleCopyJson}
                    title="仅复制纯 JSON 配置片段"
                  >
                    {copiedType === 'json' ? <Check size={11} className="text-emerald-500" /> : <Copy size={11} />}
                    <span>仅 JSON</span>
                  </button>
                </div>
              </div>

              <pre className="code-block mcp-prompt-text">{getAiPrompt()}</pre>
            </div>
          </div>
        )}
      </div>
    </section>
    <AnimatePresence>{manualAgent && <motion.div className="modal-backdrop manual-client-backdrop" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} transition={{ duration: reduced ? 0 : .2 }} onMouseDown={event => event.target === event.currentTarget && setManualAgent(null)}>
      <motion.div className="modal-window manual-client-window" role="dialog" aria-modal="true" aria-label={`连接 ${manualAgent.name}`} initial={{ y: reduced ? 0 : 10 }} animate={{ y: 0 }} exit={{ y: reduced ? 0 : 10 }} transition={{ duration: reduced ? 0 : .2 }}>
        <div className="modal-header"><h3>连接 {manualAgent.name}</h3><button className="modal-close-btn" aria-label="关闭接入说明" onClick={() => setManualAgent(null)}><X size={17} /></button></div>
        <div className="modal-body"><p>{manualAgent.manualHint || '在客户端的 MCP 页面粘贴配置。'}</p></div>
        <div className="modal-footer">{manualAgent.documentation && <button className="dialog-btn secondary" onClick={() => onOpenDocumentation?.(manualAgent.id)}>官方说明</button>}<button className="drawer-action-btn update-primary" onClick={() => onCopyText(JSON.stringify(mcpConfig, null, 2))}><Copy size={13} />复制配置</button></div>
      </motion.div>
    </motion.div>}</AnimatePresence></>
  );
};
