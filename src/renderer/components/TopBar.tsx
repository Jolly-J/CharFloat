import React from 'react';
import { 
  Power, 
  RefreshCw, 
  Settings2, 
  Sun, 
  Moon, 
  Monitor, 
  LoaderCircle
} from 'lucide-react';
import appLogo from '../assets/app-logo.png';

interface TopBarProps {
  online: boolean;
  connectedCount: number;
  busy: string;
  theme: 'system' | 'light' | 'dark';
  onToggleService: () => void;
  onRefresh: () => void;
  onCycleTheme: () => void;
  onOpenSettings: () => void;
}

export const TopBar: React.FC<TopBarProps> = ({
  online,
  connectedCount,
  busy,
  theme,
  onToggleService,
  onRefresh,
  onCycleTheme,
  onOpenSettings
}) => {
  const isStarting = busy === 'start';
  const isRefreshing = busy === 'refresh';

  return (
    <header className="app-topbar">
      {/* 品牌区：单行紧凑 */}
      <div className="topbar-brand">
        <img src={appLogo} alt="Logo" className="brand-logo-img" />
        <span className="brand-name">Office Agent Bridge</span>
        <span className="brand-version-tag">2.1</span>
      </div>

      {/* 右侧操作区：全胶囊规范（统一 28px 高度，包含居右的协同状态胶囊） */}
      <div className="topbar-actions">
        {/* 协同状态胶囊：居右展示 */}
        <div className={`capsule-pill status-capsule ${online ? 'online' : 'offline'}`}>
          <span className="capsule-dot" />
          <span className="capsule-text">
            {!online
              ? '服务未启动'
              : connectedCount > 0
              ? `${connectedCount} 个应用协同中`
              : '后台已就绪 · 等待连接'}
          </span>
        </div>
        {/* 服务启停胶囊开关 */}
        <button
          className={`capsule-pill service-capsule ${online ? 'is-running' : 'is-stopped'}`}
          disabled={Boolean(busy)}
          onClick={onToggleService}
          title={online ? '点击停止后台服务' : '点击启动后台服务'}
        >
          {isStarting ? (
            <LoaderCircle size={13} className="spin" />
          ) : (
            <Power size={13} />
          )}
          <span>{online ? '服务运行中' : '启动服务'}</span>
        </button>

        {/* 工具胶囊组：统一 28px 组合胶囊 */}
        <div className="capsule-pill tools-capsule-group">
          {/* 刷新 */}
          <button
            className="capsule-tool-btn"
            disabled={Boolean(busy)}
            onClick={onRefresh}
            title="刷新状态"
            aria-label="刷新状态"
          >
            <RefreshCw size={13} className={isRefreshing ? 'spin' : ''} />
          </button>

          <span className="capsule-inner-sep" />

          {/* 主题切换 */}
          <button
            className="capsule-tool-btn"
            onClick={onCycleTheme}
            title={`主题: ${theme === 'system' ? '跟随系统' : theme === 'dark' ? '深色' : '浅色'}`}
            aria-label="切换主题"
          >
            {theme === 'system' ? <Monitor size={13} /> : theme === 'dark' ? <Moon size={13} /> : <Sun size={13} />}
          </button>

          <span className="capsule-inner-sep" />

          {/* 设置与诊断 */}
          <button
            className="capsule-tool-btn"
            onClick={onOpenSettings}
            title="系统偏好设置与体检"
            aria-label="系统偏好设置与体检"
          >
            <Settings2 size={13} />
          </button>
        </div>
      </div>
    </header>
  );
};
