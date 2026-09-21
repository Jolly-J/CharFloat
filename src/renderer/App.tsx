import React, { useEffect, useRef, useState } from 'react';
import { 
  CheckCircle2, 
  CircleAlert, 
  X 
} from 'lucide-react';
import { TopBar } from './components/TopBar.js';
import { TopologyView } from './components/TopologyView.js';
import { LiveWorkspaceCard } from './components/LiveWorkspaceCard.js';
import { AgentHub } from './components/AgentHub.js';
import { SafetyTimeMachine } from './components/SafetyTimeMachine.js';
import { DiffModal } from './components/DiffModal.js';
import { SettingsDrawer } from './components/SettingsDrawer.js';

const api = (window as any).api;

export default function App() {
  // 核心状态
  const [status, setStatus] = useState<any>({ online: false, components: {} });
  const [info, setInfo] = useState<any>({ version: '2.0.0', port: 19890, theme: 'system' });
  const [env, setEnv] = useState<any>(null);
  const [officeAddon, setOfficeAddon] = useState<any>(null);
  const [busy, setBusy] = useState('');
  const [toast, setToast] = useState<{ text: string; error: boolean } | null>(null);

  // Agent 接入状态
  const [selectedAgents, setSelectedAgents] = useState<string[]>([]);
  const [syncSkills, setSyncSkills] = useState(true);
  const [installLogs, setInstallLogs] = useState<any[]>([]);

  // 安全时光机审计状态
  const [records, setRecords] = useState<any[]>([]);
  const [totalRecords, setTotalRecords] = useState<number>(0);
  const [offset, setOffset] = useState(0);
  const [activeRecord, setActiveRecord] = useState<any>(null);

  // 弹窗与抽屉控制
  const [confirm, setConfirm] = useState<'stop' | 'rollback' | 'clear-audit' | null>(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);

  const toastTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  // 统一通知 Toast
  const notify = (text: string, error = false) => {
    clearTimeout(toastTimer.current);
    setToast({ text, error });
    toastTimer.current = setTimeout(() => setToast(null), error ? 7500 : 4000);
  };

  // 统一异步调用封装
  const run = async (id: string, action: () => Promise<any>, successMsg?: string) => {
    if (busy) return;
    setBusy(id);
    try {
      const result = await action();
      if (result?.success === false) {
        throw new Error(result.message || result.error || '操作未成功');
      }
      if (successMsg) notify(successMsg);
      return result;
    } catch (e: any) {
      notify(e.message?.replace(/^Error invoking remote method '[^']+': Error: /, '') || '未知异常', true);
    } finally {
      setBusy('');
    }
  };

  // 刷新所有状态
  const refreshAll = async () => {
    if (!api) return;
    try {
      const s = await api.getStatus();
      setStatus(s);
    } catch {}
    try {
      const e = await api.detectEnvironment();
      setEnv(e);
      // 默认选中已检测到或未配置的 agent
      if (e?.agents && selectedAgents.length === 0) {
        const autoSelected = e.agents
          .filter((a: any) => a.status === 'configured' || a.detected)
          .map((a: any) => a.id);
        setSelectedAgents(autoSelected);
      }
    } catch {}
    try {
      const oa = await api.checkOfficeAddonStatus();
      setOfficeAddon(oa);
    } catch {}
    if (status?.online) {
      await loadAuditRecords(offset);
    }
  };

  // 加载审计历史快照
  const loadAuditRecords = async (currentOffset = offset) => {
    if (!api) return;
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
      // 忽略历史拉取失败静默
    }
  };

  // 初始化加载
  useEffect(() => {
    if (!api) {
      notify('提示：请从桌面客户端打开以连接本地完整能力。', true);
      return;
    }

    void api.getStatus().then(setStatus).catch((e: any) => notify(e.message, true));
    void api.getAppInfo().then(setInfo).catch(() => {});
    void api.detectEnvironment().then((e: any) => {
      setEnv(e);
      if (e?.agents) {
        const auto = e.agents.filter((a: any) => a.status === 'configured' || a.detected).map((a: any) => a.id);
        setSelectedAgents(auto);
      }
    }).catch(() => {});
    void api.checkOfficeAddonStatus().then(setOfficeAddon).catch(() => {});

    // 订阅后端事件
    const unsubscribe = api.onStatusChange((newStatus: any) => {
      setStatus(newStatus);
    });

    return () => {
      unsubscribe?.();
      clearTimeout(toastTimer.current);
    };
  }, []);

  // 监听主题变化并应用到 html dataset
  useEffect(() => {
    if (info?.theme) {
      document.documentElement.dataset.theme = info.theme;
    }
  }, [info?.theme]);

  // 当处于在线状态时定期同步审计记录
  useEffect(() => {
    if (status?.online) {
      void loadAuditRecords(offset);
    }
  }, [status?.online, offset]);

  // 启停服务
  const handleToggleService = () => {
    if (status?.online) {
      setConfirm('stop');
    } else {
      void run('start', async () => {
        const r = await api.startService();
        if (!r.success) throw new Error(r.message);
        await refreshAll();
      }, 'Bridge 办公桥梁服务已就绪');
    }
  };

  // 循环切换主题
  const handleCycleTheme = () => {
    const nextTheme = info.theme === 'system' ? 'light' : info.theme === 'light' ? 'dark' : 'system';
    void run('theme', async () => {
      await api.setTheme(nextTheme);
      setInfo({ ...info, theme: nextTheme });
    });
  };

  // WPS 加载项安装/修复
  const handleInstallWps = () => {
    void run('addon', async () => {
      const r = await api.installAddon();
      if (!r.success) throw new Error(r.message);
      await refreshAll();
      return r;
    }, 'WPS 加载项配置已更新');
  };

  // Office 加载项安装/修复
  const handleInstallOffice = () => {
    void run('office-addon', async () => {
      const r = await api.installOfficeAddon();
      if (!r.success) throw new Error(r.message);
      await refreshAll();
      return r;
    }, 'Microsoft Office 加载项已部署');
  };

  // Office 运行通道检查
  const handleCheckOfficeStatus = () => {
    void run('office', async () => {
      const res = await api.officeStatus();
      notify(`Office 通道自检完成：${res?.runningComponents ? Object.keys(res.runningComponents).join(', ') : '已连通'}`);
    });
  };

  // Agent 选择切换
  const handleToggleAgent = (id: string) => {
    setSelectedAgents((prev) =>
      prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]
    );
  };

  const handleToggleSelectAll = () => {
    const allIds = (env?.agents || []).map((a: any) => a.id);
    setSelectedAgents(selectedAgents.length === allIds.length ? [] : allIds);
  };

  // 执行 Agent MCP 配置安装
  const handleExecuteInstall = () => {
    void run('install', async () => {
      const r = await api.executeInstall({
        agents: selectedAgents,
        skills: syncSkills,
      });
      setInstallLogs(r.logs || []);
      await refreshAll();
      return r;
    }, `所选 ${selectedAgents.length} 个 AI 客户端配置已更新`);
  };

  // 复制文本剪贴板
  const handleCopyText = (text: string) => {
    void run('copy', async () => {
      await api.copyText(text);
    }, '已成功复制到剪贴板');
  };

  // 查看快照 Diff 详情
  const handleSelectRecord = (id: string) => {
    void run('record', async () => {
      const rec = await api.getAuditRecord(id);
      setActiveRecord(rec);
    });
  };

  // 执行二次确认的动作
  const handleConfirmAction = () => {
    const kind = confirm;
    if (!kind) return;
    void run(kind, async () => {
      if (kind === 'stop') {
        await api.stopService();
        setStatus({ online: false, components: {} });
      } else if (kind === 'clear-audit') {
        await api.clearAuditRecords();
        setOffset(0);
        await loadAuditRecords(0);
      } else if (kind === 'rollback' && activeRecord) {
        await api.rollbackRecord(activeRecord.id);
        setActiveRecord(null);
        await loadAuditRecords(offset);
      }
      setConfirm(null);
    }, kind === 'stop' ? '服务已停止' : kind === 'clear-audit' ? '快照已清空' : '已成功撤销修改并还原单元格');
  };

  // 统计连接情况
  const connectedCount = Object.values(status?.components || {}).filter((s: any) => s?.connected).length;
  const isWpsReady = Boolean(env?.addon?.installed && !env?.addon?.hasBlocked);
  const isOfficeReady = Boolean(officeAddon?.installed);
  const configuredAgentsCount = (env?.agents || []).filter((a: any) => a?.status === 'configured').length;

  return (
    <div className={`app-shell ${api?.platform === 'darwin' ? 'mac-os' : ''}`}>
      {/* 1. 顶部控制栏 */}
      <TopBar
        online={status.online}
        connectedCount={connectedCount}
        busy={busy}
        theme={info.theme || 'system'}
        onToggleService={handleToggleService}
        onRefresh={() => run('refresh', refreshAll, '全部状态已刷新')}
        onCycleTheme={handleCycleTheme}
        onOpenSettings={() => setIsSettingsOpen(true)}
      />

      {/* 2. 主体工作区 (一体化协同大盘) */}
      <main className="app-main-viewport">
        <div className="dashboard-scroll-content">
          {/* 异常警示条 */}
          {!status.online && status.notice && (
            <div className="global-notice-bar warning">
              <CircleAlert size={16} />
              <span>{status.notice}</span>
            </div>
          )}

          {/* 模块 A: 三段式连通拓扑与新手 3 步开箱引导 */}
          <TopologyView
            online={status.online}
            connectedCount={connectedCount}
            activeWorkbook={status.activeWorkbook}
            configuredAgentsCount={configuredAgentsCount}
            isWpsReady={isWpsReady}
            isOfficeReady={isOfficeReady}
            onStartService={handleToggleService}
            onGoToAgents={() => {
              const el = document.getElementById('agent-hub-section');
              el?.scrollIntoView({ behavior: 'smooth' });
            }}
            onRepairWps={handleInstallWps}
          />

          {/* 模块 B: 正在协同的工作区感知与办公软件状态看板 */}
          <LiveWorkspaceCard
            status={status}
            env={env}
            officeAddon={officeAddon}
            busy={busy}
            onInstallWpsAddon={handleInstallWps}
            onInstallOfficeAddon={handleInstallOffice}
            onCheckOfficeStatus={handleCheckOfficeStatus}
          />

          {/* 模块 C: AI 助手授权中心 (应用卡片矩阵 + 新手常用 Prompt 展台) */}
          <div id="agent-hub-section">
            <AgentHub
              agents={env?.agents || []}
              selectedAgents={selectedAgents}
              syncSkills={syncSkills}
              busy={busy}
              installLogs={installLogs}
              mcpConfig={info.config}
              onToggleAgent={handleToggleAgent}
              onToggleSelectAll={handleToggleSelectAll}
              onToggleSyncSkills={setSyncSkills}
              onExecuteInstall={handleExecuteInstall}
              onCopyText={handleCopyText}
            />
          </div>

          {/* 模块 D: 安全时光机 (修改快照流水 + 一键还原) */}
          <SafetyTimeMachine
            online={status.online}
            records={records}
            totalRecords={totalRecords}
            offset={offset}
            busy={busy}
            onRefresh={() => run('audit-refresh', () => loadAuditRecords(offset), '历史快照已刷新')}
            onClearHistory={() => setConfirm('clear-audit')}
            onSelectRecord={handleSelectRecord}
            onPageChange={(newOffset: number) => setOffset(newOffset)}
            onStartService={handleToggleService}
          />
        </div>
      </main>

      {/* 3. 底部轻量状态条 */}
      <footer className="app-footer-bar">
        <div className="footer-left">
          <span className={`status-pill-dot ${status.online ? 'active' : ''}`} />
          <span>{status.online ? `本地加密通道已监听端口 ${info.port}` : '后台服务未启动'}</span>
        </div>
        <div className="footer-right">
          <span>Office Agent Bridge · 本地高保真办公引擎</span>
          <span className="footer-divider">/</span>
          <span>{info.platform === 'darwin' ? 'macOS 客户端' : 'Windows 客户端'}</span>
        </div>
      </footer>

      {/* 4. 全局 Toast 提示 */}
      {toast && (
        <div className={`floating-toast ${toast.error ? 'error' : 'success'}`} role="status">
          {toast.error ? <CircleAlert size={16} /> : <CheckCircle2 size={16} />}
          <span className="toast-message">{toast.text}</span>
          <button className="toast-close" onClick={() => setToast(null)} aria-label="关闭">
            <X size={14} />
          </button>
        </div>
      )}

      {/* 5. 变更对比弹窗 & 二次危险确认弹层 */}
      <DiffModal
        record={activeRecord}
        confirm={confirm}
        busy={busy}
        onCloseRecord={() => setActiveRecord(null)}
        onCloseConfirm={() => setConfirm(null)}
        onRequestRollback={() => setConfirm('rollback')}
        onConfirmAction={handleConfirmAction}
      />

      {/* 6. 高级设置与系统体检抽屉 */}
      <SettingsDrawer
        isOpen={isSettingsOpen}
        theme={info.theme || 'system'}
        login={Boolean(info.login)}
        appInfo={info}
        online={status.online}
        busy={busy}
        onClose={() => setIsSettingsOpen(false)}
        onChangeTheme={(th: 'system' | 'light' | 'dark') => {
          void run('theme', async () => {
            await api.setTheme(th);
            setInfo({ ...info, theme: th });
          });
        }}
        onToggleLogin={(en: boolean) => {
          void run('login', async () => {
            const nextLogin = await api.setLogin(en);
            setInfo({ ...info, login: nextLogin });
          }, en ? '已开启开机自启' : '已关闭开机自启');
        }}
        onRunDiagnose={() => {
          void run('doctor', async () => {
            const r = await api.diagnose();
            const conn = r?.checks?.filter((c: any) => c.status === 'connected')?.length || 0;
            notify(`体检完毕：后台正常，${conn} 个办公通道就绪`);
          });
        }}
        onOpenLog={() => {
          void run('log', api.openLog);
        }}
      />
    </div>
  );
}
