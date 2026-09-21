import React, { useState } from 'react';
import { 
  Cable, 
  CheckCircle2, 
  AlertCircle, 
  Copy, 
  ChevronDown, 
  LoaderCircle,
  FolderSync
} from 'lucide-react';
import appLogo from '../assets/app-logo.png';
import doubaoLogo from '../assets/doubao-color.svg';
import workbuddyLogo from '../assets/workbuddy.svg';
import qwenLogo from '../assets/qwen-color.svg';
import kimiLogo from '../assets/kimi.webp';
import claudeLogo from '../assets/claude-color.svg';
import openaiLogo from '../assets/openai.svg';

const agentLogos: Record<string, string> = {
  doubao: doubaoLogo,
  workbuddy: workbuddyLogo,
  qwen: qwenLogo,
  kimi: kimiLogo,
  'claude-code': claudeLogo,
  codex: openaiLogo,
};

interface AgentHubProps {
  agents: any[];
  selectedAgents: string[];
  syncSkills: boolean;
  busy: string;
  installLogs: any[];
  mcpConfig: any;
  onToggleAgent: (id: string) => void;
  onToggleSelectAll: () => void;
  onToggleSyncSkills: (val: boolean) => void;
  onExecuteInstall: () => void;
  onCopyText: (text: string) => void;
}

export const AgentHub: React.FC<AgentHubProps> = ({
  agents = [],
  selectedAgents,
  syncSkills,
  busy,
  installLogs = [],
  mcpConfig,
  onToggleAgent,
  onToggleSyncSkills,
  onExecuteInstall,
  onCopyText,
}) => {
  const [showAdvanced, setShowAdvanced] = useState(false);
  const isInstalling = busy === 'install';

  return (
    <section className="agent-hub-card">
      {/* 头部：标题与快捷操作栏 */}
      <div className="agent-hub-header">
        <div className="hub-title-group">
          <div className="hub-icon-wrap">
            <img src={appLogo} alt="Logo" className="brand-logo-icon-sm" />
          </div>
          <div>
            <h3>AI 助手授权中心</h3>
            <p className="sub-caption">选择你需要打通的 AI 桌面客户端，一键授权 Bridge 办公工具集</p>
          </div>
        </div>

        <div className="hub-actions-bar">
          <label className="sync-skills-checkbox" title="将 Office/WPS 结构化操作指南同步至 AI 客户端技能库">
            <input
              type="checkbox"
              checked={syncSkills}
              onChange={(e) => onToggleSyncSkills(e.target.checked)}
            />
            <FolderSync size={12} />
            <span>同步 AI 专属技能库</span>
          </label>

          <button
            className="install-cta-btn"
            disabled={Boolean(busy) || selectedAgents.length === 0}
            onClick={onExecuteInstall}
          >
            {isInstalling ? <LoaderCircle size={13} className="spin" /> : <Cable size={13} />}
            <span>一键配置已选助手 ({selectedAgents.length})</span>
          </button>
        </div>
      </div>

      {/* AI 客户端网格 */}
      <div className="agent-grid">
        {agents.map((agent) => {
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
                <div className="agent-logo-box">
                  {logo ? (
                    <img src={logo} alt={agent.name} className="agent-logo-img" />
                  ) : (
                    <span className="agent-logo-fallback">{agent.name[0]}</span>
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
                  <span className="agent-card-title">{agent.name}</span>
                  {/* 状态徽章：严格单行并排 */}
                  {isConfigured ? (
                    <span className="agent-badge active">
                      <CheckCircle2 size={10} />
                      <span>已配置</span>
                    </span>
                  ) : isDetected ? (
                    <span className="agent-badge pending">
                      <span>待授权</span>
                    </span>
                  ) : (
                    <span className="agent-badge gray">
                      <span>未检测到</span>
                    </span>
                  )}
                </div>
                <p className="agent-detail-text">
                  {agent.detail || (isConfigured ? '配置已是最新，可直接在客户端对话' : isDetected ? '已在本机检测到客户端配置目录' : '选择后将自动创建接入配置')}
                </p>
              </div>
            </div>
          );
        })}
      </div>

      {/* 安装反馈结果面板 */}
      {installLogs.length > 0 && (
        <div className="install-feedback-panel">
          <div className="feedback-title">
            <CheckCircle2 size={13} className="text-emerald-500" />
            <span>配置更新结果</span>
          </div>
          <div className="feedback-list">
            {installLogs.map((log, idx) => (
              <div className="feedback-item" key={idx}>
                {log.status === 'success' ? (
                  <CheckCircle2 size={12} className="text-emerald-500" />
                ) : (
                  <AlertCircle size={12} className="text-amber-500" />
                )}
                <span><strong>{log.step}</strong>: {log.detail}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 专业开发者 / 手动 MCP 配置折叠栏 */}
      <div className="advanced-mcp-fold">
        <button
          className="fold-toggle-btn"
          onClick={() => setShowAdvanced(!showAdvanced)}
        >
          <span>为 Cursor / Windsurf / 自定义客户端手动配置 MCP</span>
          <ChevronDown size={13} className={showAdvanced ? 'rotate-180' : ''} />
        </button>

        {showAdvanced && (
          <div className="fold-content">
            <div className="mcp-code-header">
              <span className="sub-caption">标准 stdio MCP 配置文件片段（免额外装 Node.js）：</span>
              <button
                className="copy-code-btn"
                onClick={() => onCopyText(JSON.stringify(mcpConfig, null, 2))}
              >
                <Copy size={11} />
                <span>复制完整 JSON</span>
              </button>
            </div>
            <pre className="code-block">{JSON.stringify(mcpConfig, null, 2)}</pre>
          </div>
        )}
      </div>
    </section>
  );
};
