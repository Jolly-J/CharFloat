import React, { useEffect, useRef, useState } from 'react';
import { 
  X, 
  Monitor, 
  Sun, 
  Moon, 
  Stethoscope, 
  FileText, 
  Folder, 
  ExternalLink, 
  CheckCircle2, 
  Info, 
  LoaderCircle,
  ShieldCheck,
  Zap
} from 'lucide-react';
import appLogo from '../assets/app-logo.png';

interface SettingsDrawerProps {
  isOpen: boolean;
  theme: 'system' | 'light' | 'dark';
  login: boolean;
  appInfo: any;
  online: boolean;
  busy: string;
  onClose: () => void;
  onChangeTheme: (theme: 'system' | 'light' | 'dark') => void;
  onToggleLogin: (enabled: boolean) => void;
  onRunDiagnose: () => void;
  onOpenLog: () => void;
}

export const SettingsDrawer: React.FC<SettingsDrawerProps> = ({
  isOpen,
  theme,
  login,
  appInfo = {},
  online,
  busy,
  onClose,
  onChangeTheme,
  onToggleLogin,
  onRunDiagnose,
  onOpenLog,
}) => {
  const [rendered, setRendered] = useState(isOpen);
  const [isClosing, setIsClosing] = useState(false);
  const exitTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  useEffect(() => {
    if (isOpen) {
      clearTimeout(exitTimer.current);
      setIsClosing(false);
      setRendered(true);
    } else if (rendered && !isClosing) {
      setIsClosing(true);
      exitTimer.current = setTimeout(() => {
        setRendered(false);
        setIsClosing(false);
      }, 200);
    }
  }, [isOpen]);

  const handleClose = () => {
    if (isClosing) return;
    setIsClosing(true);
    exitTimer.current = setTimeout(() => {
      onClose();
      setIsClosing(false);
      setRendered(false);
    }, 200);
  };

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && rendered && !isClosing) {
        handleClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [rendered, isClosing]);

  if (!rendered) return null;

  const isDiagnosing = busy === 'doctor';

  return (
    <div 
      className={`drawer-backdrop ${isClosing ? 'is-exiting' : 'is-entering'}`} 
      onMouseDown={(e) => e.target === e.currentTarget && handleClose()}
    >
      <aside 
        className={`settings-drawer ${isClosing ? 'is-exiting' : 'is-entering'}`} 
        role="dialog" 
        aria-label="偏好设置与高级诊断"
      >
        {/* 抽屉头部 */}
        <div className="drawer-header">
          <div className="drawer-title-group">
            <h3>系统设置与高级维护</h3>
            <p className="text-muted">管理外观偏好、开机自启与本地底层诊断</p>
          </div>
          <button className="drawer-close-btn" onClick={handleClose} aria-label="关闭抽屉">
            <X size={17} />
          </button>
        </div>

        {/* 抽屉滚动内容 */}
        <div className="drawer-content">
          {/* 外观与个性化 */}
          <section className="drawer-section">
            <h4 className="section-title">界面外观</h4>
            <div className="theme-selector-grid">
              {[
                { id: 'system', label: '跟随系统', icon: Monitor },
                { id: 'light', label: '浅色明亮', icon: Sun },
                { id: 'dark', label: '深色沉浸', icon: Moon },
              ].map(({ id, label, icon: Icon }) => (
                <button
                  key={id}
                  className={`theme-option-card ${theme === id ? 'selected' : ''}`}
                  onClick={() => onChangeTheme(id as any)}
                >
                  <Icon size={18} />
                  <span>{label}</span>
                </button>
              ))}
            </div>
          </section>

          {/* 系统运行与自启 */}
          <section className="drawer-section">
            <h4 className="section-title">启动与运行</h4>
            <div className="setting-card-item">
              <div className="setting-text">
                <strong>开机自动启动后台</strong>
                <p className="text-muted">开机登录后在后台静默就绪，窗口保持关闭</p>
              </div>
              <button
                className={`modern-switch ${login ? 'on' : ''}`}
                role="switch"
                aria-checked={login}
                onClick={() => onToggleLogin(!login)}
              >
                <span className="switch-knob" />
              </button>
            </div>
          </section>

          {/* 本机体检与诊断 */}
          <section className="drawer-section">
            <h4 className="section-title">自检与排障</h4>
            <div className="setting-card-item">
              <div className="setting-text">
                <strong>一键全面体检 (Bridge Doctor)</strong>
                <p className="text-muted">全方位自检后台进程、端口监听、WPS/Office 通道及能力完整度</p>
              </div>
              <button
                className="drawer-action-btn"
                disabled={Boolean(busy) || !online}
                onClick={onRunDiagnose}
              >
                {isDiagnosing ? <LoaderCircle size={13} className="spin" /> : <Stethoscope size={13} />}
                <span>{isDiagnosing ? '体检中…' : '开始体检'}</span>
              </button>
            </div>

            <div className="setting-card-item">
              <div className="setting-text">
                <strong>查看本地运行日志</strong>
                <p className="text-muted text-ellipsis" title={appInfo?.home}>
                  存储在: <code>{appInfo?.home}</code>
                </p>
              </div>
              <button className="drawer-action-btn secondary" onClick={onOpenLog}>
                <FileText size={13} />
                <span>打开日志</span>
              </button>
            </div>
          </section>

          {/* 架构与产品信息 */}
          <section className="drawer-section about-section">
            <div className="about-brand-box">
              <img src={appLogo} alt="Logo" className="about-logo" />
              <div>
                <strong>字浮 CharFloat</strong>
                <span className="version-tag">版本 v{appInfo?.version || '2.0.0'}</span>
              </div>
            </div>
            <p className="about-desc">
              本地高保真办公 AI 桥梁系统 · 纯本地进程间加密通信 · 数据安全隔离在您本机的办公软件内。
            </p>
          </section>
        </div>
      </aside>
    </div>
  );
};
