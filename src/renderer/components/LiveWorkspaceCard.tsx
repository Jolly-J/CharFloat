import React from 'react';
import { 
  FileSpreadsheet, 
  FileText, 
  Presentation, 
  CheckCircle2, 
  ArrowUpRight, 
  LoaderCircle,
  Wrench,
  ShieldAlert,
  MousePointerClick
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import appLogo from '../assets/app-logo.png';
import wpsLogo from '../assets/wps-logo.png';
import officeLogo from '../assets/office-logo.png';

interface LiveWorkspaceCardProps {
  status: any;
  env: any;
  officeAddon: any;
  busy: string;
  onInstallWpsAddon: () => void;
  onInstallOfficeAddon: () => void;
  onCheckOfficeStatus?: () => void;
}

export const LiveWorkspaceCard: React.FC<LiveWorkspaceCardProps> = ({
  status,
  env,
  officeAddon,
  busy,
  onInstallWpsAddon,
  onInstallOfficeAddon,
}) => {
  const isWpsBlocked = Boolean(env?.addon?.hasBlocked);
  const wpsNeedsUpgrade = Boolean(env?.addon?.needsUpgrade);
  const wpsInstalled = Boolean(env?.addon?.installed);
  const officeInstalled = Boolean(officeAddon?.installed);
  const officeNeedsUpgrade = Boolean(officeAddon?.needsUpgrade);

  // 获取当前正在活跃操作的文档与选区
  const activeDocName = status?.activeWorkbook || 
    status?.components?.excel?.activeDocument || 
    status?.components?.['ms-excel']?.activeDocument || 
    status?.components?.word?.activeDocument || 
    status?.components?.ppt?.activeDocument;

  const activeSheet = status?.activeSheet;
  const currentSelection = status?.currentSelection?.address;
  const hasActiveDoc = Boolean(activeDocName);

  return (
    <section className="live-workspace-container">
      {/* 实时协同文档卡片 (带有 motion layout 物理高度动画) */}
      <motion.div 
        layout 
        className={`spotlight-card ${hasActiveDoc ? 'has-doc' : 'empty-doc'}`}
        transition={{ duration: 0.28, ease: [0.16, 1, 0.3, 1] }}
      >
        <div className="spotlight-header">
          <div className="spotlight-title-group">
            <div className="pulse-beacon">
              <span className={`beacon-dot ${hasActiveDoc ? 'live' : ''}`} />
              <span className={`beacon-ring ${hasActiveDoc ? 'live' : ''}`} />
            </div>
            <h3>正在协同的工作文档</h3>
          </div>
          <AnimatePresence>
            {hasActiveDoc && (
              <motion.span 
                className="live-active-tag"
                initial={{ opacity: 0, scale: 0.9 }}
                animate={{ opacity: 1, scale: 1 }}
                exit={{ opacity: 0, scale: 0.9 }}
                transition={{ duration: 0.2 }}
              >
                <img src={appLogo} alt="Logo" className="brand-badge-icon" />
                <span>实时交互就绪</span>
              </motion.span>
            )}
          </AnimatePresence>
        </div>

        {/* 使用 AnimatePresence mode="wait" 实现空状态与文档之间的平滑物理切换 */}
        <div className="spotlight-view-switch">
          <AnimatePresence mode="wait">
            {hasActiveDoc ? (
              <motion.div
                key="active-doc-stage"
                className="spotlight-body"
                initial={{ opacity: 0, y: 6 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -6 }}
                transition={{ duration: 0.22, ease: [0.16, 1, 0.3, 1] }}
              >
                <div className="doc-icon-badge">
                  <FileSpreadsheet size={20} className="text-emerald-500" />
                </div>
                <div className="doc-meta">
                  <div className="doc-filename" title={activeDocName}>
                    {activeDocName.split(/[\\/]/).pop()}
                  </div>
                  <div className="doc-sub-details">
                    {activeSheet && (
                      <span className="doc-pill sheet-pill">
                        工作表: <strong>{activeSheet}</strong>
                      </span>
                    )}
                    {currentSelection && (
                      <span className="doc-pill range-pill">
                        <MousePointerClick size={11} />
                        <span>选区: <code>{currentSelection}</code></span>
                      </span>
                    )}
                  </div>
                </div>
              </motion.div>
            ) : (
              <motion.div
                key="empty-doc-stage"
                className="spotlight-empty"
                initial={{ opacity: 0, y: 6 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -6 }}
                transition={{ duration: 0.22, ease: [0.16, 1, 0.3, 1] }}
              >
                <p>未检测到正在打开的文档。请在 WPS 或 Office 中打开任意表格，AI 将自动感知识别并建立协同。</p>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </motion.div>

      {/* 宿主办公软件卡片：WPS 与 Office 严格保持完全一致的产品化表达 */}
      <div className="host-status-grid">
        {/* WPS Office 卡片 */}
        <div className="host-card">
          <div className="host-card-top">
            <div className="host-card-brand">
              <span className="brand-avatar">
                <img src={wpsLogo} alt="WPS Office" className="wps-logo-img" />
              </span>
              <div className="brand-info-col">
                <h4>WPS Office</h4>
                <p className="sub-caption">
                  {wpsInstalled 
                    ? `已就绪 · 加载项 v${env?.addon?.installedVersion || '2.0.0'}` 
                    : '尚未部署 WPS 加载项'}
                </p>
              </div>
            </div>

            {/* 状态徽章 */}
            {isWpsBlocked ? (
              <span className="status-badge error">
                <ShieldAlert size={11} /> 策略受限
              </span>
            ) : wpsNeedsUpgrade ? (
              <span className="status-badge warning">
                可升级至 v{env?.addon?.latestVersion}
              </span>
            ) : wpsInstalled ? (
              <span className="status-badge success">
                <CheckCircle2 size={11} /> 最新就绪
              </span>
            ) : (
              <span className="status-badge warning">待安装</span>
            )}
          </div>

          {/* 组件状态胶囊行：表格、文字、演示 */}
          <div className="host-components-strip">
            {[
              { key: 'excel', label: '表格', icon: FileSpreadsheet },
              { key: 'word', label: '文字', icon: FileText },
              { key: 'ppt', label: '演示', icon: Presentation },
            ].map(({ key, label, icon: Icon }) => {
              const comp = status?.components?.[key];
              const isConnected = Boolean(comp?.connected);
              return (
                <div className={`comp-pill ${isConnected ? 'connected' : ''}`} key={key}>
                  <Icon size={12} />
                  <span>{label}</span>
                  <span className={`comp-dot ${isConnected ? 'on' : ''}`} />
                </div>
              );
            })}
          </div>

          {/* 卡片底栏 */}
          <div className="host-card-footer">
            <span className="footer-tip">
              {isWpsBlocked
                ? '因系统安全策略受限需一键修复'
                : wpsNeedsUpgrade
                ? '有更高版本加载项可用'
                : '加载项运行正常，可随时重新部署'}
            </span>
            <button
              className={`repair-btn ${isWpsBlocked || wpsNeedsUpgrade || !wpsInstalled ? 'highlight' : ''}`}
              disabled={Boolean(busy)}
              onClick={onInstallWpsAddon}
            >
              {busy === 'addon' ? (
                <LoaderCircle size={12} className="spin" />
              ) : isWpsBlocked ? (
                <Wrench size={12} />
              ) : (
                <ArrowUpRight size={12} />
              )}
              <span>
                {busy === 'addon'
                  ? '部署中…'
                  : isWpsBlocked
                  ? '一键修复安全策略'
                  : wpsNeedsUpgrade
                  ? '一键升级加载项'
                  : !wpsInstalled
                  ? '一键安装加载项'
                  : '重新部署 / 修复'}
              </span>
            </button>
          </div>
        </div>

        {/* Microsoft Office 卡片 */}
        <div className="host-card">
          <div className="host-card-top">
            <div className="host-card-brand">
              <span className="brand-avatar">
                <img src={officeLogo} alt="Microsoft Office" className="office-logo-img" />
              </span>
              <div className="brand-info-col">
                <h4>Microsoft Office</h4>
                <p className="sub-caption">Microsoft Office 办公套件协同</p>
              </div>
            </div>

            {/* 状态徽章 */}
            {officeNeedsUpgrade ? (
              <span className="status-badge warning">需更新加载项</span>
            ) : officeInstalled ? (
              <span className="status-badge success">
                <CheckCircle2 size={11} /> 最新就绪
              </span>
            ) : (
              <span className="status-badge warning">待安装</span>
            )}
          </div>

          {/* 组件状态胶囊行：与 WPS 完全对称展示表格、文字、演示 */}
          <div className="host-components-strip">
            {[
              { key: 'ms-excel', label: '表格', icon: FileSpreadsheet },
              { key: 'word', label: '文字', icon: FileText },
              { key: 'ppt', label: '演示', icon: Presentation },
            ].map(({ key, label, icon: Icon }) => {
              const comp = status?.components?.[key];
              const isConnected = Boolean(comp?.connected);
              return (
                <div className={`comp-pill ${isConnected ? 'connected' : ''}`} key={key}>
                  <Icon size={12} />
                  <span>{label}</span>
                  <span className={`comp-dot ${isConnected ? 'on' : ''}`} />
                </div>
              );
            })}
          </div>

          {/* 卡片底栏 */}
          <div className="host-card-footer">
            <span className="footer-tip">
              {officeInstalled
                ? '加载项运行正常，可在 Office 中随时启用'
                : '尚未配置加载项，点击右侧一键安装'}
            </span>
            <button
              className={`repair-btn ${!officeInstalled || officeNeedsUpgrade ? 'highlight' : ''}`}
              disabled={Boolean(busy)}
              onClick={onInstallOfficeAddon}
            >
              {busy === 'office-addon' ? (
                <LoaderCircle size={12} className="spin" />
              ) : (
                <ArrowUpRight size={12} />
              )}
              <span>
                {busy === 'office-addon'
                  ? '部署中…'
                  : officeNeedsUpgrade
                  ? '一键升级加载项'
                  : !officeInstalled
                  ? '一键安装加载项'
                  : '重新部署 / 修复'}
              </span>
            </button>
          </div>
        </div>
      </div>
    </section>
  );
};
