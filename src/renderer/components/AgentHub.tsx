import React, { useState } from 'react';
import { 
  Cable, 
  CheckCircle2, 
  AlertCircle, 
  Copy, 
  ChevronDown, 
  LoaderCircle,
  Check,
  Bot
} from 'lucide-react';
import appLogo from '../assets/app-logo.png';
import brandGreeting from '../assets/brand-greeting.png';
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
  busy: string;
  installLogs: any[];
  mcpConfig: any;
  onToggleAgent: (id: string) => void;
  onToggleSelectAll: () => void;
  onExecuteInstall: () => void;
  onCopyText: (text: string) => void;
  onOpenDoubaoGuide?: () => void;
  onOpenAgentApp?: (id: string) => void;
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
}) => {
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [copiedType, setCopiedType] = useState<'prompt' | 'json' | null>(null);
  const isInstalling = busy === 'install';

  const getAiPrompt = () => {
    const jsonStr = JSON.stringify(mcpConfig, null, 2);
    return `请帮我把本地「字浮 CharFloat」办公自动化服务配置到你（当前 AI 助手）的 MCP (Model Context Protocol) 列表中。

服务配置参数如下（请将其安全合并到你的 mcpServers 配置文件中，不要覆盖已有其他服务）：
\`\`\`json
${jsonStr}
\`\`\`

配置执行说明：
1. 检查并定位你所在客户端的 MCP 配置文件（如 Cursor 用户的 ~/.cursor/mcp.json 或当前项目的 .cursor/mcp.json，Windsurf 用户的 ~/.codeium/windsurf/mcp_config.json 等）；
2. 读取现有配置内容，并将上述 "charfloat" 服务条目合并进 "mcpServers" 对象；
3. 保存后请告知我已完成配置，并提醒我重载或刷新 MCP 服务。`;
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
    <section className="agent-hub-card">
      {/* 头部：标题与快捷操作栏 */}
      <div className="agent-hub-header">
        <div className="hub-title-group">
          <div className="hub-icon-wrap">
            <img src={brandGreeting} alt="AI 助手授权中心" className="hub-greeting-img" />
          </div>
          <div>
            <h3>AI 助手授权中心</h3>
            <p className="sub-caption">选择你需要打通的 AI 桌面客户端，一键授权 Bridge 办公工具集</p>
          </div>
        </div>

        <div className="hub-actions-bar">
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
                <div className="agent-top-left">
                  <div className="agent-logo-box">
                    {logo ? (
                      <img src={logo} alt={agent.name} className="agent-logo-img" />
                    ) : (
                      <span className="agent-logo-fallback">{agent.name[0]}</span>
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
            <span>让 AI 助手为自己配置 MCP（直接发指令给 Cursor / Windsurf 等）</span>
          </div>
          <ChevronDown size={13} className={showAdvanced ? 'rotate-180' : ''} />
        </button>

        {showAdvanced && (
          <div className="fold-content" style={{ marginTop: '10px' }}>
            <div className="mcp-prompt-guide-box">
              <div className="mcp-prompt-header">
                <div>
                  <span className="mcp-prompt-tip-title">无需手动翻找配置文件</span>
                  <p className="sub-caption" style={{ margin: '2px 0 0' }}>
                    复制下方指令发给你的 AI 编程助手，它会自动定位配置文件并完成写入：
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
                        <span>已复制指令，快发给 AI 吧</span>
                      </>
                    ) : (
                      <>
                        <Copy size={12} />
                        <span>复制给 AI 的配置指令</span>
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
  );
};
