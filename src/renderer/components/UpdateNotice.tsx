import React, { useEffect, useState } from 'react';
import { motion, useReducedMotion } from 'motion/react';
import { ArrowRight, X } from 'lucide-react';
import idea from '../assets/update-idea.png';
import appLogo from '../assets/app-logo.png';

const api = (window as any).api;
export function UpdateNotice() {
  const [info, setInfo] = useState<{ version: string; currentVersion: string; notes: string }>();
  const [error, setError] = useState('');
  const reduced = useReducedMotion();
  useEffect(() => {
    if (!info && !error) return;
    const frame = requestAnimationFrame(() => { void api.updateNoticeReady().catch(() => {}); });
    return () => cancelAnimationFrame(frame);
  }, [info, error]);
  useEffect(() => {
    void api.updateNoticeInfo().then((value: any) => {
      document.documentElement.setAttribute('data-theme', value.theme); setInfo(value);
    }).catch(() => setError('提醒内容未能加载，请打开字浮查看更新。'));
    const keydown = (event: KeyboardEvent) => { if (event.key === 'Escape') void api.dismissUpdateNotice(); };
    window.addEventListener('keydown', keydown); return () => window.removeEventListener('keydown', keydown);
  }, []);
  return <main className="update-reminder">
    <header className="update-reminder-header"><span><img src={appLogo} alt="" />字浮 CharFloat</span><button aria-label="关闭更新提醒" onClick={() => api.dismissUpdateNotice()}><X size={17} /></button></header>
    <motion.div className="update-reminder-content" initial={{ opacity: 0, y: reduced ? 0 : 8 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: reduced ? 0 : 0.35 }}>
      <div className="update-reminder-hero"><span className="update-reminder-halo" /><img src={idea} alt="字浮正在想新点子" /></div>
      <span className="update-reminder-eyebrow">有新版本</span>
      <h1>字浮，又进步了一点</h1>
      <div className="update-reminder-versions"><span>{info?.currentVersion || '当前版本'}</span><ArrowRight size={12} /><strong>{info?.version || '新版本'}</strong></div>
      {info?.notes && <p className="update-reminder-notes" title={info.notes}>{info.notes}</p>}
      {error && <p className="update-reminder-error" role="alert">{error}</p>}
    </motion.div>
    <footer className="update-reminder-footer"><div><button className="drawer-action-btn secondary" onClick={() => api.dismissUpdateNotice()}>稍后</button><button className="drawer-action-btn update-primary" onClick={() => api.openUpdateFromNotice().catch(() => setError('未能打开客户端，请从托盘打开字浮。'))}>查看新版<ArrowRight size={14} /></button></div></footer>
  </main>;
}
