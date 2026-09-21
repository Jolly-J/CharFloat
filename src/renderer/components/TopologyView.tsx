import React, { useState, useEffect } from 'react';
import { 
  Check, 
  ShieldCheck, 
  ChevronRight,
  Bot,
  CheckCircle2,
  ChevronDown
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import wpsLogo from '../assets/wps-logo.png';
import officeLogo from '../assets/office-logo.png';
import appLogo from '../assets/app-logo.png';

interface TopologyViewProps {
  online: boolean;
  connectedCount: number;
  activeWorkbook?: string;
  configuredAgentsCount: number;
  isWpsReady: boolean;
  isOfficeReady: boolean;
  onStartService: () => void;
  onGoToAgents: () => void;
  onRepairWps: () => void;
}

export const TopologyView: React.FC<TopologyViewProps> = ({
  online,
  connectedCount,
  configuredAgentsCount,
  isWpsReady,
  isOfficeReady,
  onStartService,
  onGoToAgents,
  onRepairWps,
}) => {
  const step = !online ? 1 : (!isWpsReady && !isOfficeReady) ? 2 : configuredAgentsCount === 0 ? 3 : 4;
  
  // 维护向导完成后的平滑展示与折叠状态，避免瞬间蒸发
  const [guideDismissed, setGuideDismissed] = useState(false);

  // 如果从未完成进入完成态，延时 3 秒后再平滑折叠收起
  useEffect(() => {
    if (step === 4) {
      const timer = setTimeout(() => {
        setGuideDismissed(true);
      }, 3500);
      return () => clearTimeout(timer);
    } else {
      setGuideDismissed(false);
    }
  }, [step]);

  return (
    <motion.section layout className="clean-topology-card">
      {/* 头部标题与极简说明 */}
      <div className="topology-title-bar">
        <div className="title-left">
          <h2>实时连通拓扑</h2>
          <span className="title-sub">办公套件与 AI 助手通过本地加密管道双向交互</span>
        </div>
        <div className="title-right">
          <ShieldCheck size={13} className="text-muted" />
          <span>本地安全隔离</span>
        </div>
      </div>

      {/* 连通流水线 */}
      <div className="topology-track">
        {/* 节点 1: 办公套件 */}
        <div className={`track-node ${connectedCount > 0 ? 'active' : ''}`}>
          <div className="node-icon-row">
            <img src={wpsLogo} alt="WPS" className="app-icon-img" />
            <img src={officeLogo} alt="Office" className="app-icon-img" />
          </div>
          <div className="node-title">办公软件</div>
          <div className="node-detail">
            {connectedCount > 0 
              ? `${connectedCount} 个应用已连接` 
              : isWpsReady || isOfficeReady 
              ? '加载项已就绪' 
              : '待配置加载项'}
          </div>
          <div className="node-status-badge">
            <span className={`badge-dot ${connectedCount > 0 ? 'green' : isWpsReady || isOfficeReady ? 'blue' : 'gray'}`} />
            <span>{connectedCount > 0 ? '实时协同' : isWpsReady || isOfficeReady ? '已就绪' : '待配置'}</span>
          </div>
        </div>

        {/* Apple 极简流光微导管 1 */}
        <div className={`conduit-connector ${online && (connectedCount > 0 || isWpsReady || isOfficeReady) ? 'active' : ''}`}>
          <span className="conduit-anchor start" />
          <div className="conduit-track">
            <div className="conduit-base" />
            {online && (connectedCount > 0 || isWpsReady || isOfficeReady) && (
              <div className={`conduit-beam ${connectedCount > 0 ? 'streaming' : ''}`} />
            )}
          </div>
          <span className="conduit-anchor end" />
        </div>

        {/* 节点 2: Bridge 桥梁核心 */}
        <div className={`track-node ${online ? 'active' : ''}`}>
          <div className="node-icon-row">
            <div className="core-icon-box bridge-core-box">
              <img src={appLogo} alt="Bridge 核心" className="bridge-brand-node-img" />
            </div>
          </div>
          <div className="node-title">Bridge 核心</div>
          <div className="node-detail">
            {online ? '端口 19890 守护中' : '服务未运行'}
          </div>
          <div className="node-status-badge">
            <span className={`badge-dot ${online ? 'green' : 'gray'}`} />
            <span>{online ? '正常运行' : '已停止'}</span>
          </div>
        </div>

        {/* Apple 极简流光微导管 2 */}
        <div className={`conduit-connector ${online && configuredAgentsCount > 0 ? 'active' : ''}`}>
          <span className="conduit-anchor start" />
          <div className="conduit-track">
            <div className="conduit-base" />
            {online && configuredAgentsCount > 0 && (
              <div className="conduit-beam" />
            )}
          </div>
          <span className="conduit-anchor end" />
        </div>

        {/* 节点 3: AI 助手客户端 */}
        <div className={`track-node ${configuredAgentsCount > 0 ? 'active' : ''}`}>
          <div className="node-icon-row">
            <div className="core-icon-box ai-agent-box">
              <Bot size={18} className="ai-agent-icon" />
            </div>
          </div>
          <div className="node-title">AI 智能体</div>
          <div className="node-detail">
            {configuredAgentsCount > 0 
              ? `已配置 ${configuredAgentsCount} 个客户端` 
              : '尚未授权客户端'}
          </div>
          <div className="node-status-badge">
            <span className={`badge-dot ${configuredAgentsCount > 0 ? 'green' : 'gray'}`} />
            <span>{configuredAgentsCount > 0 ? '已授权' : '待配置'}</span>
          </div>
        </div>
      </div>

      {/* 新手 3 步走引导条：借助 motion 实现平滑补间与完成反馈，绝不瞬间闪退 */}
      <AnimatePresence>
        {!guideDismissed && (
          <motion.div
            key="onboarding-guide"
            initial={{ opacity: 0, height: 0 }}
            animate={{ opacity: 1, height: 'auto' }}
            exit={{ opacity: 0, height: 0, marginTop: 0 }}
            transition={{ duration: 0.32, ease: [0.16, 1, 0.3, 1] }}
            style={{ overflow: 'hidden' }}
          >
            <div className="clean-onboarding-box">
              {step === 4 ? (
                /* 完成庆祝态：温和反馈，不突兀消失 */
                <motion.div 
                  className="guide-completed-pill"
                  initial={{ opacity: 0, scale: 0.96 }}
                  animate={{ opacity: 1, scale: 1 }}
                  transition={{ duration: 0.2 }}
                >
                  <CheckCircle2 size={15} className="text-emerald-500" />
                  <span>恭喜！全链路已连通就绪，现在打开办公文档即可与 AI 开始对话协同。</span>
                  <button 
                    className="guide-dismiss-btn"
                    onClick={() => setGuideDismissed(true)}
                  >
                    收起
                  </button>
                </motion.div>
              ) : (
                /* 步骤进行态 */
                <>
                  <div className="onboarding-label">
                    <img src={appLogo} alt="Logo" className="brand-badge-icon" />
                    <span>就绪向导</span>
                  </div>

                  <div className="onboarding-steps-row">
                    {/* 步骤 1 */}
                    <motion.div layout className={`ob-step ${step === 1 ? 'current' : step > 1 ? 'done' : ''}`}>
                      <div className="step-badge">{step > 1 ? <Check size={11} /> : '1'}</div>
                      <div className="step-copy">
                        <strong>启动服务</strong>
                        <span>开启本地桥梁</span>
                      </div>
                      {step === 1 && (
                        <button className="ob-btn" onClick={onStartService}>
                          启动
                        </button>
                      )}
                    </motion.div>

                    <ChevronRight size={12} className="ob-sep" />

                    {/* 步骤 2 */}
                    <motion.div layout className={`ob-step ${step === 2 ? 'current' : step > 2 ? 'done' : ''}`}>
                      <div className="step-badge">{step > 2 ? <Check size={11} /> : '2'}</div>
                      <div className="step-copy">
                        <strong>加载项就绪</strong>
                        <span>一键部署插件</span>
                      </div>
                      {step === 2 && (
                        <button className="ob-btn" onClick={onRepairWps}>
                          部署
                        </button>
                      )}
                    </motion.div>

                    <ChevronRight size={12} className="ob-sep" />

                    {/* 步骤 3 */}
                    <motion.div layout className={`ob-step ${step === 3 ? 'current' : ''}`}>
                      <div className="step-badge">3</div>
                      <div className="step-copy">
                        <strong>授权 AI</strong>
                        <span>绑定客户端</span>
                      </div>
                      {step === 3 && (
                        <button className="ob-btn" onClick={onGoToAgents}>
                          前往
                        </button>
                      )}
                    </motion.div>
                  </div>
                </>
              )}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </motion.section>
  );
};
