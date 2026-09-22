/**
 * 完全磁盘访问权限引导（置顶浮窗版，方案 B）。
 *
 * 与旧弹窗的区别：
 * - **置顶**：能停在系统设置面板上方，用户可以直接把下面的图标拖进去；
 * - **可拖拽的应用图标**：拖拽即把当前该授权的 App 包交给设置面板，不需要用户自己去找路径；
 * - **面向用户的语言**：不再贴原始路径当主角，路径退到"高级信息"里。
 */
import React, { useEffect, useState } from 'react';
import { ShieldAlert, ExternalLink, RefreshCw, GripVertical } from 'lucide-react';

// 注意：这里**不用** api?. 可选链——预加载失败时可选链会把"整个桥不可用"变成"点了没反应"，
// 曾经因此让用户以为按钮坏了。桥缺失必须显式暴露。
const bridge = (window as any).api;
const api = bridge as Record<string, (...args: any[]) => any>;

export function PermissionFloat() {
  const [issue, setIssue] = useState<any>(null);
  const [info, setInfo] = useState<{ appToAuthorize: string; isDev: boolean; platform: string } | null>(null);
  const [busy, setBusy] = useState(false);
  const [showPath, setShowPath] = useState(false);
  const [actionError, setActionError] = useState('');
  const [iconUrl, setIconUrl] = useState('');

  useEffect(() => {
    if (!bridge?.permissionInfo) return;
    bridge.permissionInfo().then((v: any) => setInfo(v)).catch(() => {});
    Promise.resolve(bridge.appIcon?.()).then((v: string) => setIconUrl(v || '')).catch(() => {});
    const off = bridge.onPermissionIssue?.((v: any) => setIssue(v));
    return () => { off?.(); };
  }, []);

  const isDev = info?.isDev ?? Boolean(issue?.isDev);
  const target = info?.appToAuthorize || issue?.appToAuthorize || '';
  const appName = target ? (target.split('/').pop() || '').replace(/\.app$/, '') : '本应用';

  const openSettings = async () => {
    setActionError('');
    try { await api.openFullDiskAccess(); }
    catch (e: any) { setActionError(`打开设置失败：${e?.message || e}`); }
  };
  /**
   * 拖拽必须挂在 HTML5 的 dragstart 上（不是 mousedown），并走 send 单向消息。
   * 用 invoke 的异步往返会在主进程 startDrag 之前就结束拖拽会话——表现就是"拖不动"。
   */
  const onDragStart = (e: React.DragEvent) => {
    e.preventDefault();
    if (!bridge?.startAppDrag) { setActionError('拖拽不可用：窗口桥未就绪，请点下方位置手动添加。'); setShowPath(true); return; }
    api.startAppDrag();
  };
  const retry = async () => {
    setBusy(true); setActionError('');
    try { await api.retryPermissionInstall(); }
    catch (e: any) { setActionError(`重试失败：${e?.message || e}`); }
    finally { setBusy(false); }
  };

  if (!bridge) {
    return (
      <div className="pg-root">
        <div className="pg-body" style={{ justifyContent: 'center' }}>
          <div className="pg-head">
            <span className="pg-head-icon"><ShieldAlert size={18} /></span>
            <div>
              <h1>引导窗口未就绪</h1>
              <p>窗口桥加载失败。请到「系统设置 → 隐私与安全性 → 完全磁盘访问权限」手动添加本应用后重试安装。</p>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="pg-root">
      <div className="pg-body">
        <div className="pg-head">
          <span className="pg-head-icon"><ShieldAlert size={18} /></span>
          <div>
            <h1>还差一步就能装好</h1>
            <p>macOS 要求你亲自授权，系统不会替应用弹窗。</p>
          </div>
        </div>

        <section className="pg-step">
          <span className="pg-step-no">1</span>
          <div className="pg-step-body">
            <strong>打开权限设置面板</strong>
            <p>会直接跳到「完全磁盘访问权限」。</p>
            <button className="pg-btn pg-btn-primary" onClick={openSettings}>
              <ExternalLink size={14} /> 打开设置面板
            </button>
          </div>
        </section>

        <section className="pg-step">
          <span className="pg-step-no">2</span>
          <div className="pg-step-body">
            <strong>把下面这个图标拖进列表并勾选</strong>
            <p>{isDev ? '当前是开发模式，请拖 Electron（与打包版权限互不相通）。' : `请拖 ${appName}。`}</p>
            <div className="pg-drag-card" draggable onDragStart={onDragStart} title="按住我，拖到设置面板的列表里">
              <span className="pg-drag-handle"><GripVertical size={14} /></span>
              {iconUrl
                ? <img className="pg-app-icon" src={iconUrl} alt="" draggable={false} />
                : <span className="pg-app-icon pg-app-icon-fallback" aria-hidden />}
              <span className="pg-drag-name">{appName}</span>
              <span className="pg-drag-hint">拖我</span>
            </div>
            <button className="pg-link" onClick={() => setShowPath(v => !v)}>
              {showPath ? '隐藏位置' : '图标拖不动？查看位置'}
            </button>
            {showPath && (
              <code className="pg-path" onClick={() => api?.copyText?.(target)} title="点击复制">{target}</code>
            )}
          </div>
        </section>

        <section className="pg-step">
          <span className="pg-step-no">3</span>
          <div className="pg-step-body">
            <strong>回到这里点「重试」</strong>
            <p>刚授权完通常需要重启本应用才生效。</p>
          </div>
        </section>

        {actionError && <div className="pg-error">{actionError}</div>}

        {issue?.detail && (
          <details className="pg-detail">
            <summary>技术详情</summary>
            <pre>{issue.detail}</pre>
          </details>
        )}
      </div>

      {/* 固定页脚：整行主操作，滚动内容时不跟着滚走 */}
      <footer className="pg-actions">
        <button className="pg-btn pg-btn-primary" onClick={retry} disabled={busy}>
          <RefreshCw size={15} className={busy ? 'pg-spin' : ''} /> {busy ? '正在重试…' : '我已授权，重试'}
        </button>
      </footer>
    </div>
  );
}
