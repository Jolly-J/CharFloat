import React from 'react';
import { AnimatePresence, motion, useReducedMotion } from 'motion/react';
import { ArrowDownToLine, Check, CircleAlert, LoaderCircle, RefreshCw, RotateCw, ChevronRight } from 'lucide-react';
import type { UpdateState } from '../../main/update-service.js';

export interface UpdatePanelProps {
  update?: UpdateState;
  currentVersion: string;
  action: string;
  error: string;
  onCheck: () => void;
  onDownload: () => void;
  onCancel: () => void;
  onInstall: () => void;
}

export function UpdatePanel({ update, currentVersion, action, error, onCheck, onDownload, onCancel, onInstall }: UpdatePanelProps) {
  const reduced = useReducedMotion();
  const [notesOpen, setNotesOpen] = React.useState(false);
  const [historyOpen, setHistoryOpen] = React.useState(update?.installation?.phase === 'failed');
  const phase = update?.phase || 'idle';
  const checking = phase === 'checking', downloading = phase === 'downloading', installing = phase === 'installing';
  const ready = phase === 'ready', failed = phase === 'failed' || Boolean(error);
  const known = Boolean(update?.availableVersion), disabled = Boolean(action) || installing;
  const percent = Math.min(100, Math.max(0, (update?.bytes || 0) * 100 / (update?.totalBytes || 1)));
  const transition = { duration: reduced ? 0 : 0.24, ease: [0.22, 1, 0.36, 1] as const };
  const title = installing ? '正在准备重启' : downloading ? '正在下载更新' : ready ? '更新已准备好' : checking ? '正在检查更新' : known ? `有新版本 ${update?.availableVersion}` : phase === 'up-to-date' ? '已是最新版本' : '版本更新';
  const Icon = failed ? CircleAlert : checking || installing ? LoaderCircle : ready || phase === 'up-to-date' ? Check : known ? ArrowDownToLine : RefreshCw;
  const buttonKey = checking || downloading ? 'cancel' : installing ? 'installing' : ready ? 'ready' : known ? 'download' : 'check';

  return <motion.div layout={!reduced} layoutDependency={`${phase}-${Boolean(error)}`} transition={transition} className={`update-panel ${failed ? 'has-error' : ''}`}>
    <div className="update-panel-heading">
      <span className="update-panel-icon"><Icon size={19} className={checking || installing ? 'spin' : ''} /></span>
      <div className="update-panel-title">
        <AnimatePresence mode="wait" initial={false}>
          <motion.strong key={title} initial={{ opacity: 0, y: reduced ? 0 : 3 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0 }} transition={transition}>{title}</motion.strong>
        </AnimatePresence>
        <span>当前版本 {currentVersion}{known && (ready || downloading) ? ` → ${update?.availableVersion}` : ''}</span>
      </div>
      {downloading && <span className="update-percent">{Math.floor(percent)}<small>%</small></span>}
      {ready && <span className="update-ready-tag">已校验</span>}
    </div>

    <AnimatePresence initial={false}>
      {downloading && <motion.div key="progress" className="update-progress-area" initial={{ height: 0, opacity: 0 }} animate={{ height: 'auto', opacity: 1 }} exit={{ height: 0, opacity: 0 }} transition={transition}>
        <div className="update-progress-track" role="progressbar" aria-label="更新下载进度" aria-valuemin={0} aria-valuemax={update?.totalBytes || 1} aria-valuenow={update?.bytes || 0}>
          <motion.div className="update-progress-fill" initial={false} animate={{ width: `${percent}%` }} transition={{ duration: reduced ? 0 : 0.3, ease: 'easeOut' }} />
        </div>
        <div className="update-progress-detail"><span>{((update?.bytes || 0) / 1024 ** 2).toFixed(1)} / {((update?.totalBytes || 0) / 1024 ** 2).toFixed(1)} MB</span><span>后台下载，可继续使用</span></div>
      </motion.div>}
    </AnimatePresence>

    <AnimatePresence mode="wait" initial={false}>
      <motion.div key={phase} className="update-panel-body" initial={{ height: 0, opacity: 0 }} animate={{ height: 'auto', opacity: 1 }} exit={{ height: 0, opacity: 0 }} transition={transition}>
        {['failed', 'cancelled', 'unconfigured'].includes(phase) && <p className="update-message" aria-live="polite">{update?.message}</p>}
        {update?.notes && !downloading && !installing && <><button className="update-disclosure" onClick={() => setNotesOpen(value => !value)} aria-expanded={notesOpen}>更新内容<ChevronRight size={12} style={{ transform: notesOpen ? 'rotate(90deg)' : undefined }} /></button><AnimatePresence initial={false}>{notesOpen && <motion.p className="update-release-notes" initial={{ height: 0, opacity: 0 }} animate={{ height: 'auto', opacity: 1 }} exit={{ height: 0, opacity: 0 }} transition={transition}>{update.notes}</motion.p>}</AnimatePresence></>}
      </motion.div>
    </AnimatePresence>
    <AnimatePresence initial={false}>
      {error && <motion.p className="update-error" role="alert" initial={{ height: 0, opacity: 0 }} animate={{ height: 'auto', opacity: 1 }} exit={{ height: 0, opacity: 0 }} transition={transition}>{error}</motion.p>}
    </AnimatePresence>

    <motion.div layout={!reduced} transition={transition} className="update-panel-actions">
      <AnimatePresence mode="wait" initial={false}>
        <motion.div key={buttonKey} className="update-actions" initial={{ opacity: 0, y: reduced ? 0 : 4 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: reduced ? 0 : -3 }} transition={transition}>
          {checking || downloading ? <button className="drawer-action-btn secondary" onClick={onCancel}>取消</button>
            : installing ? <button className="drawer-action-btn update-primary" disabled><LoaderCircle size={13} className="spin" />正在准备…</button>
            : ready ? <><button className="drawer-action-btn update-primary" disabled={disabled} onClick={onInstall}><RotateCw size={13} />重启并更新</button><button className="drawer-action-btn secondary" disabled={disabled} onClick={onDownload}>重新下载</button></>
            : <>{known && <button className="drawer-action-btn update-primary" disabled={disabled} onClick={onDownload}><ArrowDownToLine size={13} />下载更新</button>}<button className={`drawer-action-btn ${known ? 'secondary' : 'update-primary'}`} disabled={disabled || !update || phase === 'unconfigured'} onClick={onCheck}><RefreshCw size={12} />检查更新</button></>}
        </motion.div>
      </AnimatePresence>
    </motion.div>
    <p className="update-panel-hint">{ready ? '请先保存办公文档。点击后重启字浮，配置会保留。' : installing ? '正在安全关闭后台，完成后自动打开新版本。' : '下载完成后由你决定何时重启，当前工作不会中断。'}</p>
    {update?.installation && <div className={`update-installation-note ${update.installation.phase === 'failed' ? 'is-error' : ''}`}><button className="update-disclosure" onClick={() => setHistoryOpen(value => !value)} aria-expanded={historyOpen}>上次更新{update.installation.phase === 'failed' ? '未完成' : ''}<ChevronRight size={12} style={{ transform: historyOpen ? 'rotate(90deg)' : undefined }} /></button><AnimatePresence initial={false}>{historyOpen && <motion.p initial={{ height: 0, opacity: 0 }} animate={{ height: 'auto', opacity: 1 }} exit={{ height: 0, opacity: 0 }} transition={transition}>{update.installation.message}</motion.p>}</AnimatePresence></div>}
  </motion.div>;
}
