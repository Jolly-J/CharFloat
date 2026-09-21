import React, { useEffect, useRef, useState } from 'react';
import { Activity, Cable, History, Settings2, ArrowUpRight, ChevronDown, ChevronRight, Check, Copy, RefreshCw, Power, FileSpreadsheet, FileText, Presentation, Info, X, LoaderCircle, ArrowRight, CheckCircle2, CircleAlert, Monitor, Sun, Moon, ExternalLink, RotateCcw, Trash2 } from 'lucide-react';
import appLogo from './assets/app-logo.png';
import wpsLogo from './assets/wps-logo.png';
import officeLogo from './assets/office-logo.png';
import doubaoLogo from './assets/doubao-color.svg';
import workbuddyLogo from './assets/workbuddy.svg';
import qwenLogo from './assets/qwen-color.svg';
import kimiLogo from './assets/kimi.webp';
import claudeLogo from './assets/claude-color.svg';
import openaiLogo from './assets/openai.svg';

const agentLogos: Record<string, string> = {
  doubao: doubaoLogo,
  workbuddy: workbuddyLogo,
  qwen: qwenLogo,
  kimi: kimiLogo,
  'claude-code': claudeLogo,
  codex: openaiLogo
};

type Tab = 'overview' | 'connections' | 'history' | 'settings';
const tabs = [{ id: 'overview', label: '概览', icon: Activity }, { id: 'connections', label: 'AI 接入', icon: Cable }, { id: 'history', label: '修改记录', icon: History }, { id: 'settings', label: '设置', icon: Settings2 }] as const;
const api = (window as any).api;
const cn = (...parts: (string | false | undefined)[]) => parts.filter(Boolean).join(' ');
function Badge({ children, tone = '' }: { children: React.ReactNode; tone?: string }) { return <span className={cn('badge', tone)}>{children}</span>; }
function Modal({ title, children, close }: { title: string; children: React.ReactNode; close: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const previous = document.activeElement as HTMLElement;
    const target = ref.current!;
    const first = target.querySelector('button') as HTMLElement; first?.focus();
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') close();
      if (e.key === 'Tab') {
        const nodes = Array.from(target.querySelectorAll<HTMLElement>('button:not(:disabled), input, a, [tabindex="0"]'));
        const index = nodes.indexOf(document.activeElement as HTMLElement);
        if (e.shiftKey && index === 0) { e.preventDefault(); nodes.at(-1)?.focus(); }
        if (!e.shiftKey && index === nodes.length - 1) { e.preventDefault(); nodes[0]?.focus(); }
      }
    };
    document.addEventListener('keydown', key);
    return () => { document.removeEventListener('keydown', key); previous?.focus(); };
  }, []);
  return <div className="modal-overlay" onMouseDown={e => { if (e.target === e.currentTarget) close(); }}><div className="modal" ref={ref} role="dialog" aria-modal="true" aria-label={title}><div className="modal-title"><h2>{title}</h2><button className="icon-button" aria-label="关闭对话框" onClick={close}><X size={17}/></button></div>{children}</div></div>;
}
export default function App() {
  const [tab, setTab] = useState<Tab>('overview');
  const [status, setStatus] = useState<any>({ online: false, components: {} });
  const [info, setInfo] = useState<any>({ version: '2.0.0', port: 19890, theme: 'system' });
  const [env, setEnv] = useState<any>(null);
  const [busy, setBusy] = useState('');
  const [toast, setToast] = useState<{ text: string; error: boolean } | null>(null);
  const [selected, setSelected] = useState<string[]>([]);
  const [syncSkills, setSyncSkills] = useState(true);
  const [logs, setLogs] = useState<any[]>([]);
  const [advanced, setAdvanced] = useState(false);
  const [office, setOffice] = useState<any>(null);
  const [officeAddon, setOfficeAddon] = useState<any>(null);
  const [records, setRecords] = useState<any[]>([]);
  const [totalRecords, setTotalRecords] = useState<number>(0);
  const [record, setRecord] = useState<any>(null);
  const [offset, setOffset] = useState(0);
  const [confirm, setConfirm] = useState<'stop' | 'rollback' | 'clear-audit' | null>(null);
  const toastTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const notify = (text: string, error = false) => { clearTimeout(toastTimer.current); setToast({ text, error }); toastTimer.current = setTimeout(() => setToast(null), error ? 9000 : 4500); };
  const run = async (id: string, action: () => Promise<any>, message?: string) => {
    if (busy) return;
    setBusy(id);
    try { const result = await action(); if (result?.success === false) throw new Error(result.message || result.error); if (message) notify(message); return result; }
    catch (e: any) { notify(e.message.replace(/^Error invoking remote method '[^']+': Error: /, ''), true); }
    finally { setBusy(''); }
  };
  const refresh = async () => {
    setStatus(await api.getStatus());
    setEnv(await api.detectEnvironment());
    try { setOfficeAddon(await api.checkOfficeAddonStatus()); } catch {}
  };
  useEffect(() => {
    if (!api) { notify('请从桌面客户端打开；浏览器预览未连接本机服务。', true); return; }
    void api.getStatus().then(setStatus).catch((e: any) => notify(e.message, true));
    void api.getAppInfo().then(setInfo).catch((e: any) => notify(e.message, true));
    void api.detectEnvironment().then(setEnv).catch((e: any) => notify(e.message, true));
    void api.checkOfficeAddonStatus().then(setOfficeAddon).catch(() => {});
    const unsubscribe = api.onStatusChange(setStatus);
    return () => { unsubscribe(); clearTimeout(toastTimer.current); };
  }, []);
  useEffect(() => { document.documentElement.dataset.theme = info.theme; }, [info.theme]);
  const loadAuditRecords = async (currentOffset = offset) => {
    try {
      const res = await api.getAuditRecords({ offset: currentOffset });
      if (Array.isArray(res)) {
        setRecords(res);
        setTotalRecords(res.length);
      } else {
        setRecords(res?.records || []);
        setTotalRecords(res?.total ?? (res?.records ? res.records.length : 0));
      }
    } catch (e: any) {
      notify(e.message, true);
    }
  };
  useEffect(() => {
    if (tab !== 'history' || !api || !status.online) return;
    void loadAuditRecords(offset);
  }, [tab, offset, status.online]);
  const start = () => run('start', async () => { await api.startService().then((r: any) => { if (!r.success) throw new Error(r.message); }); await refresh(); }, '后台服务已启动');
  const title = tabs.find(t => t.id === tab)!.label;
  const activeWorkbook = status.activeWorkbook;
  const connected = Object.values(status.components || {}).filter((s: any) => s.connected).length;
  const isWpsBlocked = !!env?.addon?.hasBlocked && !connected;
  return <div className={cn('desktop', api?.platform === 'darwin' && 'mac')}>
    <header className="titlebar">
      <div className="titlebar-center">
        <img src={appLogo} alt="Logo" className="app-logo-img"/>
        <span className="titlebar-title">Office Agent Bridge</span>
        <span className="titlebar-caption">跨平台办公 AI 桥梁</span>
      </div>
      <span className="titlebar-status">
        <i className={cn('status-dot', status.online && 'online')}/>
        {status.online ? '后台运行中' : '服务未就绪'}
      </span>
    </header>
    <div className="workspace">
      <aside className="sidebar"><nav aria-label="主要导航">{tabs.map(({ id, label, icon: Icon }) => <button key={id} className={cn('nav-item', tab === id && 'active')} aria-current={tab === id ? 'page' : undefined} onClick={() => setTab(id)}><Icon size={17}/><span>{label}</span>{id === 'connections' && <span className="nav-count">MCP</span>}</button>)}</nav><div className="sidebar-foot"><span className="tiny-label">OFFICE AGENT BRIDGE</span><span>v{info.version}</span><span className="muted">窗口可关闭<br/>{status.online ? '服务仍在后台运行' : '启动后可后台运行'}</span></div></aside>
      <div className="main-shell"><div className="page-heading"><h1>{title}</h1><button className="icon-button" title="刷新状态" aria-label="刷新状态" disabled={!!busy} onClick={() => run('refresh', refresh)}><RefreshCw size={16} className={busy === 'refresh' ? 'spin' : ''}/></button></div>
      <main className="page-scroll" key={tab}>
      {tab === 'overview' && <div className="page-content">
        <section className="service-card">
          <img src={appLogo} alt="Office Agent Bridge" className="service-logo-hero"/>
          <div className="service-copy">
            <div className="inline">
              <h2>{status.online ? 'Bridge 已就绪' : '启动你的办公连接'}</h2>
              {status.online && <Badge tone="green">运行中</Badge>}
            </div>
            <p>{status.online ? `${connected} 个办公组件已连接 · MCP 与窗口独立运行` : '启动本机后台，让 AI 连接正在打开的 Office 与 WPS 办公文档。'}</p>
          </div>
          <button className={status.online ? 'button secondary' : 'button primary'} disabled={!!busy} onClick={status.online ? () => setConfirm('stop') : start}>
            {busy === 'start' ? <LoaderCircle className="spin" size={15}/> : <Power size={15}/>}
            <span>{status.online ? '停止' : '启动服务'}</span>
          </button>
        </section>
        {!status.online && status.notice && <div className="notice warning"><CircleAlert size={16}/><span>{status.notice}</span></div>}
        <div className="section-label"><h2>办公软件</h2><span>本机连接</span></div>
        <section className="panel host-panel">
          <div className="host-header">
            <span className="product-mark wps">
              <img src={wpsLogo} alt="WPS Office" className="host-logo-img"/>
            </span>
            <div>
              <div className="inline" style={{ gap: 8 }}>
                <h3>WPS Office</h3>
                {env?.addon?.installed ? (
                  <Badge tone={isWpsBlocked ? 'red' : env?.addon?.needsUpgrade ? 'yellow' : 'green'}>
                    {isWpsBlocked ? '被 WPS 阻断' : env?.addon?.needsUpgrade ? `需升级至 v${env?.addon?.latestVersion}` : `v${env?.addon?.installedVersion} 最新`}
                  </Badge>
                ) : (
                  <Badge tone="yellow">未安装加载项</Badge>
                )}
              </div>
              <p>{env?.addon?.installed ? `已安装加载项 v${env?.addon?.installedVersion} · 最新可用 v${env?.addon?.latestVersion || info.version}` : '尚未安装加载项 · 点击下方一键安装'}</p>
            </div>
            <Badge tone={connected ? 'green' : ''}>{connected ? '已连接' : '未连接'}</Badge>
          </div>
          {[{ key: 'excel', label: '表格', icon: FileSpreadsheet, hint: '推荐使用' }, { key: 'word', label: '文字', icon: FileText, hint: '按工具清单使用' }, { key: 'ppt', label: '演示', icon: Presentation, hint: '有限支持' }].map(({ key, label, icon: Icon, hint }) => { const s = status.components?.[key]; return <div className="host-row" key={key}><Icon size={17}/><span className="component-label">{label}</span><span className="document-name" title={s?.activeDocument}>{s?.connected ? s.activeDocument?.split(/[\\/]/).pop() || '已连接，等待文档' : '等待连接'}</span><span className="component-hint">{hint}</span><i className={cn('status-dot', s?.connected && 'online')}/></div>; })}
          <div className="host-footer">
            <span>
              {isWpsBlocked
                ? '检测到加载项安全策略受限；请点击右侧一键修复'
                : env?.addon?.needsUpgrade
                  ? `加载项版本 (v${env?.addon?.installedVersion}) 需要升级至 v${env?.addon?.latestVersion}`
                  : `加载项已就绪 (v${env?.addon?.installedVersion})；连接状态以各组件为准`}
            </span>
            <button
              className={env?.addon?.needsUpgrade || isWpsBlocked ? 'button small primary' : 'text-button'}
              disabled={!!busy}
              onClick={() => run('addon', async () => {
                const r = await api.installAddon();
                if (!r.success) throw new Error(r.message);
                await refresh();
                notify(r.message);
              })}
            >
              {busy === 'addon' ? <LoaderCircle size={13} className="spin"/> : (env?.addon?.needsUpgrade || isWpsBlocked) ? <ArrowUpRight size={13}/> : null}
              <span>{busy === 'addon' ? '正在部署…' : isWpsBlocked ? '一键修复安全策略' : env?.addon?.needsUpgrade ? `一键升级插件至 v${env?.addon?.latestVersion || '2.0.0'}` : '重新部署 / 修复加载项'}</span>
              {!(env?.addon?.needsUpgrade || isWpsBlocked) && <ArrowUpRight size={13}/>}
            </button>
          </div>
        </section>
        <section className="panel office-panel">
          <div className="host-header">
            <span className="product-mark microsoft">
              <img src={officeLogo} alt="Microsoft Office" className="host-logo-img"/>
            </span>
            <div>
              <div className="inline" style={{ gap: 8 }}>
                <h3>Microsoft Office</h3>
                {officeAddon?.installed ? (
                  <Badge tone={officeAddon?.needsUpgrade ? 'yellow' : 'green'}>
                    {officeAddon?.needsUpgrade ? `需更新至 v${officeAddon?.latestVersion}` : `v${officeAddon?.installedVersion} 加载项就绪`}
                  </Badge>
                ) : (
                  <Badge tone="yellow">未部署加载项</Badge>
                )}
              </div>
              <p>Office.js 官方架构 · 42 项全量 Excel 原生结构化操作</p>
            </div>
            <button className="button small secondary" disabled={!!busy || !status.online} onClick={() => run('office', async () => setOffice(await api.officeStatus()))}>
              {busy === 'office' ? '检查中…' : '检查连接'}
            </button>
          </div>
          {(() => {
            const s = status.components?.['ms-excel'];
            return (
              <div className="host-row">
                <FileSpreadsheet size={17} />
                <span className="component-label">表格 (Excel)</span>
                <span className="document-name" title={s?.activeDocument}>
                  {s?.connected ? s.activeDocument?.split(/[\\/]/).pop() || '已连接，等待文档' : officeAddon?.installed ? '已部署，待在 Excel 中启用' : '等待部署加载项'}
                </span>
                <span className="component-hint">Office.js 官方架构 · 42 项能力</span>
                <i className={cn('status-dot', s?.connected && 'online')} />
              </div>
            );
          })()}
          {office && (
            <div className="office-result">
              {['excel', 'word', 'ppt'].map(key => (
                <span key={key}>
                  <i className={cn('status-dot', office.runningComponents?.[key] && 'online')} />
                  {key === 'ppt' ? 'PowerPoint' : key === 'word' ? 'Word' : 'Excel (系统通道)'} · {office.runningComponents?.[key] ? '可访问' : '未连接'}
                </span>
              ))}
            </div>
          )}
          <div className="host-footer">
            <span>
              {officeAddon?.installed
                ? `加载项已就绪；在 Excel【插入 -> 我的加载项】启用`
                : '尚未部署 Office 官方加载项清单'}
            </span>
            <button
              className={officeAddon?.needsUpgrade ? 'button small primary' : 'text-button'}
              disabled={!!busy}
              onClick={() => run('office-addon', async () => {
                const r = await api.installOfficeAddon();
                if (!r.success) throw new Error(r.message);
                try { setOfficeAddon(await api.checkOfficeAddonStatus()); } catch {}
                notify(r.message);
              })}
            >
              {busy === 'office-addon' ? <LoaderCircle size={13} className="spin" /> : officeAddon?.needsUpgrade ? <ArrowUpRight size={13} /> : null}
              <span>{busy === 'office-addon' ? '正在部署…' : officeAddon?.needsUpgrade ? '一键部署 Office 加载项' : '重新部署 / 修复加载项'}</span>
              {!officeAddon?.needsUpgrade && <ArrowUpRight size={13} />}
            </button>
          </div>
        </section>
        {activeWorkbook && <div className="selection-strip"><MouseSelection/><span>{status.activeSheet || '当前工作表'}</span><code>{status.currentSelection?.address || '—'}</code></div>}
        <div className="subtle-note"><Info size={15}/><span>表格是主要使用场景。演示支持基础编辑，复杂图表和母版请先确认能力。</span></div>
      </div>}
      {tab === 'connections' && <div className="page-content">
        <div className="intro"><h2>让 AI 接入你的办公软件</h2><p>选择需要配置的客户端。自动同步 Office Agent Bridge 条目，保留其他设置。</p></div>
        <section className="panel agent-list">{env?.agents?.map((agent: any) => <label className={cn('agent-row', selected.includes(agent.id) && 'selected')} key={agent.id}><input type="checkbox" checked={selected.includes(agent.id)} onChange={e => setSelected(v => e.target.checked ? [...v, agent.id] : v.filter(x => x !== agent.id))}/><span className="agent-icon-wrap">{agentLogos[agent.id] ? <img src={agentLogos[agent.id]} alt={agent.name} className="agent-icon-img"/> : <span className="agent-avatar">{agent.name[0]}</span>}</span><div className="agent-copy"><strong>{agent.name}</strong><span>{agent.detail || (agent.detected ? '已检测到本机配置目录' : '未检测到 · 选择后创建配置')}</span></div><Badge tone={agent.status === 'configured' ? 'green' : ''}>{agent.status === 'configured' ? '配置一致' : agent.detected ? '待配置' : '未检测到'}</Badge></label>)}</section>
        <label className="check-option"><input type="checkbox" checked={syncSkills} onChange={e => setSyncSkills(e.target.checked)}/><span>同步 Agent 使用技能<span className="muted"> · 支持技能目录的客户端</span></span></label>
        <div className="action-line"><span className="muted">已选择 {selected.length} 个客户端</span><button className="button primary" disabled={!!busy || !selected.length} onClick={() => run('install', async () => { const r = await api.executeInstall({ agents: selected, skills: syncSkills }); setLogs(r.logs); await refresh(); return r; }, '所选客户端配置已更新')} >{busy === 'install' ? <LoaderCircle size={15} className="spin"/> : <Cable size={15}/>}配置所选客户端</button></div>
        {logs.length > 0 && <div className="install-results" role="status">{logs.map((l, i) => <div key={i}>{l.status === 'success' ? <CheckCircle2 size={15}/> : <CircleAlert size={15}/>}<span><strong>{l.step}</strong> · {l.detail}</span></div>)}</div>}
        <section className="panel disclosure"><button className="disclosure-toggle" aria-expanded={advanced} onClick={() => setAdvanced(v => !v)}><span>其他客户端 / 手动配置</span><ChevronDown size={16} className={advanced ? 'rotate' : ''}/></button><div className={cn('disclosure-body', advanced && 'expanded')} inert={!advanced} aria-hidden={!advanced}><div><p>添加一个 stdio MCP 服务。下方使用本机安装路径，无需另外安装 Node.js。</p><div className="code-header"><span>MCP 配置</span><button className="text-button" onClick={() => run('copy', () => api.copyText(JSON.stringify(info.config, null, 2)), '已复制配置')}><Copy size={13}/>复制</button></div><pre>{JSON.stringify(info.config, null, 2)}</pre><p className="muted">HTTP 地址：127.0.0.1:{info.port}/mcp。需要先启动后台并配置本机 Bearer 凭据；请按 Skill 指引接入。</p></div></div></section>
      </div>}
      {tab === 'history' && <div className="page-content">
        <div className="intro history-intro">
          <div>
            <h2>每一次修改，有据可查</h2>
            <p>当前记录单元格值与公式修改。样式、图表和结构操作不在回滚范围内。</p>
          </div>
          {status.online && (
            <div className="history-actions">
              <button className="button small secondary" title="刷新修改记录" disabled={!!busy} onClick={() => run('audit-refresh', () => loadAuditRecords(offset), '修改记录已刷新')}>
                <RefreshCw size={13} className={busy === 'audit-refresh' ? 'spin' : ''}/>
                <span>刷新</span>
              </button>
              {totalRecords > 0 && (
                <button className="button small secondary danger-btn" title="清空全部历史记录" disabled={!!busy} onClick={() => setConfirm('clear-audit')}>
                  <Trash2 size={13}/>
                  <span>清空记录</span>
                </button>
              )}
            </div>
          )}
        </div>
        {!status.online ? <div className="empty"><Power size={27}/><h3>先启动 Bridge</h3><p>连接后台后查看本地修改记录。</p><button className="button secondary" onClick={start} disabled={!!busy}>启动服务</button></div> : !records.length ? <div className="empty"><History size={29}/><h3>{offset ? '没有更早的记录' : '还没有修改记录'}</h3><p>AI 修改单元格值或公式后，记录会显示在这里。</p></div> : <section className="panel history-list">{records.map(rec => <button className="history-row" key={rec.id} onClick={() => run('record', async () => setRecord(await api.getAuditRecord(rec.id)))} disabled={!!busy}><span className={cn('history-icon', rec.status === 'rolled_back' && 'undone')}><History size={17}/></span><span className="history-copy"><strong>{rec.description}</strong><span>{rec.workbookName} / {rec.sheetName} <code>{rec.address}</code></span><small>{new Date(rec.timestamp).toLocaleString('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })} · {rec.clientName} · {rec.modifiedCount} 个单元格</small></span><Badge>{rec.status === 'rolled_back' ? '已撤销' : '已修改'}</Badge><ChevronRight size={15}/></button>)}</section>}
        {totalRecords > 0 && (
          <div className="pagination">
            <button className="button small secondary" disabled={!offset || !!busy} onClick={() => setOffset(Math.max(0, offset - 20))}>上一页</button>
            <span className="pagination-text">第 {Math.floor(offset / 20) + 1} 页 / 共 {Math.max(1, Math.ceil(totalRecords / 20))} 页 <span className="muted">（共 {totalRecords} 条记录）</span></span>
            <button className="button small secondary" disabled={offset + 20 >= totalRecords || !!busy} onClick={() => setOffset(offset + 20)}>下一页</button>
          </div>
        )}
        <div className="subtle-note"><Info size={15}/><span>回滚前会检查区域是否有后续修改，避免覆盖你或其他 AI 新写入的内容。</span></div>
      </div>}
      {tab === 'settings' && <div className="page-content">
        <div className="section-label first"><h2>外观与启动</h2></div><section className="panel settings-list"><div className="setting-row"><div><strong>主题</strong><p>选择适合当前环境的外观</p></div><div className="segmented" aria-label="主题">{[{key:'system',icon:Monitor,label:'系统'},{key:'light',icon:Sun,label:'浅色'},{key:'dark',icon:Moon,label:'深色'}].map(({key,icon:Icon,label})=><button key={key} aria-pressed={info.theme===key} className={info.theme===key?'selected':''} onClick={() => run('theme', async () => { await api.setTheme(key); setInfo({ ...info, theme: key }); })}><Icon size={14}/>{label}</button>)}</div></div><div className="setting-row"><div><strong>登录时启动</strong><p>在后台启动服务，窗口保持关闭</p></div><button role="switch" aria-checked={!!info.login} aria-label="登录时启动" disabled={!!busy} className={cn('switch', info.login && 'on')} onClick={() => run('login', async () => setInfo({ ...info, login: await api.setLogin(!info.login) }))}><span/></button></div></section>
        <div className="section-label"><h2>运行与诊断</h2></div><section className="panel settings-list"><div className="setting-row"><div><strong>连接诊断</strong><p>检查后台、组件连接和能力范围</p></div><button className="button small secondary" disabled={!!busy || !status.online} onClick={() => run('doctor', async () => { const r=await api.diagnose(); notify(`后台正常；${r.checks.filter((c:any)=>c.status==='connected').length} 个办公组件已连接。未连接组件请检查加载项。`); })}>{busy==='doctor'?'检查中…':'运行检查'}</button></div><div className="setting-row"><div><strong>运行日志</strong><p className="path-text" title={info.home}>{info.home}</p></div><button className="text-button" onClick={() => run('log', api.openLog)}>打开日志<ArrowUpRight size={14}/></button></div><div className="setting-row"><div><strong>关闭窗口后的行为</strong><p>服务继续运行；需要断开所有 AI 时，主动停止服务。</p></div></div></section>
        <div className="about-row"><span className="about-brand"><img src={appLogo} alt="Office Agent Bridge" className="about-logo-img"/><strong>Office Agent Bridge</strong><Badge>v{info.version}</Badge></span><p>本地运行 · 数据在你的 Office / WPS 办公软件中处理</p></div>
      </div>}
      </main><footer className="statusbar"><span><i className={cn('status-dot', status.online && 'online')}/>{status.online ? `127.0.0.1:${info.port}` : '后台未连接'}</span><span>{info.platform === 'darwin' ? 'macOS' : info.platform === 'win32' ? 'Windows' : '桌面客户端'}</span></footer></div>
    </div>
    {toast && <div className={cn('toast', toast.error && 'error')} role={toast.error ? 'alert' : 'status'}>{toast.error ? <CircleAlert size={17}/> : <CheckCircle2 size={17}/>}<span>{toast.text}</span><button className="icon-button" aria-label="关闭提示" onClick={() => setToast(null)}><X size={15}/></button></div>}
    {record && <Modal title="修改详情" close={() => { setRecord(null); setConfirm(null); }}><p className="modal-description">{record.workbookName} / {record.sheetName} · {record.address}</p><p>{record.description}</p><div className="diff-table"><table><thead><tr><th>单元格</th><th>修改前</th><th>修改后</th></tr></thead><tbody>{record.diff?.slice(0,100).map((d:any,i:number)=><tr key={i}><td><code>{d.cell}</code></td><td>{String(d.oldFormula?.startsWith?.('=') ? d.oldFormula : d.oldValue ?? '空')}</td><td>{String(d.newFormula?.startsWith?.('=') ? d.newFormula : d.newValue ?? '空')}</td></tr>)}</tbody></table></div>{record.diff?.length>100&&<p className="muted">仅展示前 100 个差异。</p>}<div className="modal-actions"><button className="button secondary" onClick={() => setRecord(null)}>关闭</button>{record.status !== 'rolled_back' && <button className="button danger" disabled={!!busy} onClick={() => setConfirm('rollback')}><RotateCcw size={14}/>撤销这次修改</button>}</div></Modal>}
    {confirm && <Modal title={confirm==='stop'?'停止 Bridge 服务？':confirm==='clear-audit'?'清空全部修改记录？':'撤销这次单元格修改？'} close={() => setConfirm(null)}><p className="modal-description">{confirm==='stop'?'所有 AI 客户端会断开连接。办公软件保持打开，之后可重新启动 Bridge。':confirm==='clear-audit'?'此操作将永久清空本地所有历史修改记录与回滚快照，不可恢复。':'将恢复记录中的值和公式。如果区域已有后续修改，系统会拒绝覆盖。'}</p><div className="modal-actions"><button className="button secondary" onClick={() => setConfirm(null)}>取消</button><button className="button danger" disabled={!!busy} onClick={() => { const kind=confirm; void run(kind, async () => { if(kind==='stop'){await api.stopService();setStatus({online:false,components:{}});}else if(kind==='clear-audit'){await api.clearAuditRecords();setOffset(0);await loadAuditRecords(0);}else{await api.rollbackRecord(record.id);setRecord(null);await loadAuditRecords(offset);}setConfirm(null); },kind==='stop'?'服务已停止':kind==='clear-audit'?'修改记录已全部清空':'已撤销修改'); }}>{busy===confirm?<LoaderCircle size={15} className="spin"/>:null}{confirm==='stop'?'停止服务':confirm==='clear-audit'?'确认清空':'确认撤销'}</button></div></Modal>}
  </div>;
}
function MouseSelection(){return <span className="selection-icon"><FileSpreadsheet size={14}/></span>}
