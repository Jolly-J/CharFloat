import React, { useEffect, useRef, useState } from 'react';
import { AnimatePresence, motion, useReducedMotion } from 'motion/react';
import { X, Monitor, Sun, Moon, LoaderCircle, Check, CircleAlert, Minus, ChevronRight, RefreshCw, FileText, Power, Settings2, ArrowDownToLine, PlugZap } from 'lucide-react';
import appLogo from '../assets/app-logo.png';
import type { UpdateState } from '../../main/update-service.js';
import type { DesktopDiagnostics, DiagnosticCheck } from '../../main/desktop-diagnostics.js';
import { UpdatePanel } from './UpdatePanel.js';

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
  onExit: () => void;
  onDiagnosticAction: (action: NonNullable<DiagnosticCheck['action']>) => void;
  diagnostics?: DesktopDiagnostics;
  diagnosticError: string;
  update?: UpdateState;
  updateAction: string;
  updateError: string;
  updateFocus: number;
  onCheckUpdate: () => void;
  onDownloadUpdate: () => void;
  onCancelUpdate: () => void;
  onInstallUpdate: () => void;
}
const tabs = [{ id: 'general', label: '常用', Icon: Settings2 }, { id: 'updates', label: '更新', Icon: ArrowDownToLine }, { id: 'connection', label: '连接检查', Icon: PlugZap }] as const;
type Tab = typeof tabs[number]['id'];
export const SettingsDrawer: React.FC<SettingsDrawerProps> = (props) => {
  const { isOpen, theme, login, appInfo, busy, update, updateAction, updateError, updateFocus, diagnostics, diagnosticError } = props;
  const reduced = useReducedMotion();
  const [tab, setTab] = useState<Tab>('general');
  const [expanded, setExpanded] = useState(false);
  const closeButton = useRef<HTMLButtonElement>(null);
  const aside = useRef<HTMLElement>(null);
  const transition = { duration: reduced ? 0 : 0.23, ease: [0.22, 1, 0.36, 1] as const };
  useEffect(() => { if (isOpen && updateFocus) setTab('updates'); }, [updateFocus]);
  useEffect(() => {
    if (!isOpen) return;
    const key = (event: KeyboardEvent) => {
      if (event.key === 'Escape') props.onClose();
      if (event.key === 'Tab') {
        const targets = [...(aside.current?.querySelectorAll<HTMLElement>('button:not(:disabled), [tabindex="0"]') || [])].filter(element => element.getClientRects().length > 0);
        const first = targets[0], last = targets.at(-1);
        if (event.shiftKey && document.activeElement === first && last) { event.preventDefault(); last.focus(); }
        else if (!event.shiftKey && document.activeElement === last && first) { event.preventDefault(); first.focus(); }
      }
    };
    window.addEventListener('keydown', key); closeButton.current?.focus();
    return () => window.removeEventListener('keydown', key);
  }, [isOpen]);
  const checking = busy === 'doctor';
  const actionText = (action: DiagnosticCheck['action']) => action === 'start' ? '启动' : action === 'agents' ? '查看' : '配置';

  return <AnimatePresence>{isOpen && <motion.div className="drawer-backdrop settings-backdrop" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} transition={transition}
    onMouseDown={event => event.target === event.currentTarget && props.onClose()}>
    <motion.aside ref={aside} className="settings-drawer settings-branded" role="dialog" aria-modal="true" aria-label="字浮设置" initial={{ x: reduced ? 0 : 28 }} animate={{ x: 0 }} exit={{ x: reduced ? 0 : 28 }} transition={transition}>
      <header className="settings-brand-header"><div><img src={appLogo} alt="" /><span>设置</span><button className="settings-version" onClick={() => setTab('updates')}>v{appInfo?.version || '2.2.4'}</button></div><button ref={closeButton} className="drawer-close-btn" onClick={props.onClose} aria-label="关闭设置"><X size={18} /></button></header>
      <nav className="settings-tabs" role="tablist" aria-label="设置分类">{tabs.map(({ id, label, Icon }) => <button key={id} role="tab" aria-selected={tab === id} aria-controls={`settings-${id}`} onClick={() => setTab(id)}>
        {tab === id && <motion.span layoutId="settings-active-tab" className="settings-tab-indicator" transition={transition} />}<Icon size={14} /><span>{label}</span>{id === 'updates' && update?.availableVersion && <i className="settings-tab-dot" />}
      </button>)}</nav>
      <div className="drawer-content settings-content"><AnimatePresence mode="wait" initial={false}>
        <motion.div key={tab} id={`settings-${tab}`} role="tabpanel" className="settings-tab-content" initial={{ opacity: 0, y: reduced ? 0 : 7 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: reduced ? 0 : -5 }} transition={transition}>
          {tab === 'general' && <>
            <section className="settings-section"><h4>外观</h4><div className="settings-theme-grid">{[{ id: 'system', label: '跟随系统', Icon: Monitor }, { id: 'light', label: '浅色', Icon: Sun }, { id: 'dark', label: '深色', Icon: Moon }].map(({ id, label, Icon }) =>
              <button key={id} className={`settings-theme ${theme === id ? 'selected' : ''}`} aria-pressed={theme === id} onClick={() => props.onChangeTheme(id as any)}><span className={`theme-preview preview-${id}`}><i /><i /><i /></span><span><Icon size={13} />{label}{theme === id && <Check size={12} />}</span></button>
            )}</div></section>
            <section className="settings-section"><h4>启动与关闭</h4><div className="settings-list"><div className="settings-row"><div><strong>开机启动</strong><span>登录后在后台运行</span></div><button className={`modern-switch ${login ? 'on' : ''}`} role="switch" aria-label="开机启动" aria-checked={login} disabled={Boolean(busy)} onClick={() => props.onToggleLogin(!login)}><span className="switch-knob" /></button></div><div className="settings-row"><div><strong>关闭窗口</strong></div><span className="settings-row-value">留在{appInfo?.platform === 'darwin' ? '菜单栏' : '托盘'}</span></div></div></section>
            <button className="settings-help-link" onClick={() => setTab('connection')}><PlugZap size={15} /><span>连接有问题？</span><ChevronRight size={14} /></button>
          </>}
          {tab === 'updates' && <section id="version-update" className="settings-section"><UpdatePanel update={update} currentVersion={appInfo.version} action={updateAction} error={updateError} onCheck={props.onCheckUpdate} onDownload={props.onDownloadUpdate} onCancel={props.onCancelUpdate} onInstall={props.onInstallUpdate} /></section>}
          {tab === 'connection' && <>
            <div className="connection-check-heading"><div><h4>{checking ? '正在检查连接' : diagnostics ? diagnostics.attention ? `${diagnostics.attention} 项需要处理` : '检查完成' : '检查本机连接'}</h4><span>{diagnostics ? new Date(diagnostics.checkedAt).toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' }) : '后台、办公软件与 AI 配置'}</span></div><button className="drawer-action-btn update-primary" disabled={Boolean(busy)} onClick={props.onRunDiagnose}>{checking ? <LoaderCircle size={13} className="spin" /> : <RefreshCw size={13} />}{checking ? '检查中' : diagnostics ? '重新检查' : '开始检查'}</button></div>
            {diagnosticError && <p role="alert" className="update-error">{diagnosticError}</p>}
            <AnimatePresence initial={false}>{diagnostics && <motion.div key="results" className="diagnostic-results" initial={{ opacity: 0, height: 0 }} animate={{ opacity: 1, height: 'auto' }} exit={{ opacity: 0, height: 0 }} transition={transition}>
              {diagnostics.checks.map(check => <div className="diagnostic-row" key={check.id}><span className={`diagnostic-status ${check.status}`}>{check.status === 'ok' ? <Check size={13} /> : check.status === 'attention' ? <CircleAlert size={13} /> : <Minus size={13} />}</span><div><strong>{check.title}</strong><span title={check.detail}>{check.value}</span></div>{check.action && <button onClick={() => props.onDiagnosticAction(check.action!)} disabled={Boolean(busy)}>{actionText(check.action)}<ChevronRight size={12} /></button>}</div>)}
              {diagnostics.clients.length > 0 && <><button className="diagnostic-details-toggle" aria-expanded={expanded} onClick={() => setExpanded(value => !value)}>AI 配置详情<ChevronRight size={13} style={{ transform: expanded ? 'rotate(90deg)' : undefined }} /></button><AnimatePresence initial={false}>{expanded && <motion.div className="diagnostic-client-list" initial={{ height: 0, opacity: 0 }} animate={{ height: 'auto', opacity: 1 }} exit={{ height: 0, opacity: 0 }} transition={transition}>{diagnostics.clients.map(client => <div key={client.id}><strong>{client.name}</strong><span>{client.mcp}</span><small>{client.skills}</small></div>)}</motion.div>}</AnimatePresence></>}
            </motion.div>}</AnimatePresence>
            <button className="settings-help-link" onClick={props.onOpenLog}><FileText size={15} /><span>查看日志</span><ChevronRight size={14} /></button>
          </>}
        </motion.div>
      </AnimatePresence></div>
      <footer className="settings-footer"><span><img src={appLogo} alt="" />字浮 CharFloat</span><button onClick={props.onExit} disabled={updateAction === 'install'}><Power size={13} />退出字浮</button></footer>
    </motion.aside>
  </motion.div>}</AnimatePresence>;
};
