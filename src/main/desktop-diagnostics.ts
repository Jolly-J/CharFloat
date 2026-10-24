import { probeService, serviceRequest } from '../bridge/service-client.js';
import { AddonInstaller, OfficeAddonInstaller } from './addon-installer.js';
import { InstallerEngine, type AgentRuntimeInfo } from './installer-engine.js';
import { VERSION } from '../bridge/runtime.js';

export type DiagnosticCheck = { id: string; title: string; status: 'ok' | 'attention' | 'idle' | 'unknown'; value: string; detail?: string; action?: 'start' | 'wps' | 'office' | 'agents' };
export type DesktopDiagnostics = { checkedAt: number; checks: DiagnosticCheck[]; attention: number; clients: { id: string; name: string; mcp: string; skills: string }[] };
type Dependencies = {
  probe: typeof probeService;
  status: () => Promise<any>;
  wps: () => any;
  office: () => Promise<any>;
  clients: () => AgentRuntimeInfo[];
};
/** All checks are observations. No process, Office application, installer or arbitrary MCP command is started. */
export async function runDesktopDiagnostics(overrides: Partial<Dependencies> = {}): Promise<DesktopDiagnostics> {
  const deps: Dependencies = { probe: probeService, status: () => serviceRequest('/api/v1/status'), wps: () => AddonInstaller.checkStatus(), office: () => OfficeAddonInstaller.checkStatus(), clients: () => InstallerEngine.detectEnvironment().agents, ...overrides };
  const checks: DiagnosticCheck[] = [];
  const probe = await deps.probe(); let status: any;
  if (probe.ready) {
    checks.push({ id: 'bridge', title: '字浮后台', status: probe.info?.version === VERSION ? 'ok' : 'attention', value: probe.info?.version === VERSION ? '运行正常' : '版本不一致', detail: `后台 ${probe.info?.version || '未知'} · 客户端 ${VERSION}` });
    try { status = await deps.status(); } catch { checks.push({ id: 'bridge-status', title: '后台状态', status: 'attention', value: '读取失败', detail: '后台可达，但状态请求未成功。' }); }
  } else checks.push({ id: 'bridge', title: '字浮后台', status: 'attention', value: probe.occupied ? '端口被占用' : '尚未启动', action: probe.occupied ? undefined : 'start' });

  let wps: any, office: any;
  try { wps = deps.wps(); } catch {}
  try { office = await deps.office(); } catch {}
  for (const [id, title, installed, connected, current] of [
    ['wps', 'WPS', wps?.installed, Boolean(status?.isWpsConnected), wps?.needsUpgrade === true ? false : status?.build?.wpsAddon?.deployedMatchesResource ?? wps?.current],
    ['office', 'Microsoft Excel', office?.installed, Boolean(status?.isMsOfficeConnected || status?.components?.msExcel?.connected), office?.current],
  ] as const) {
    const connectedKnown = Boolean(status);
    checks.push({ id, title, status: installed === false || current === false ? 'attention' : connected ? 'ok' : installed === undefined ? 'unknown' : connectedKnown ? 'idle' : 'unknown',
      value: installed === false ? '未安装加载项' : current === false ? '加载项需更新' : connected ? '已连接' : installed === undefined ? '加载项状态未确认' : connectedKnown ? '等待文档连接' : '连接状态未确认',
      detail: installed === false || current === false ? '可重新配置加载项。' : connected ? '由当前后台连接状态确认。' : '打开文档后，再检查连接。',
      ...((installed === false || current === false) ? { action: id } : {}),
    });
  }
  let clients: AgentRuntimeInfo[] = [];
  try { clients = deps.clients(); }
  catch { checks.push({ id: 'clients-read', title: 'AI 接入配置', status: 'unknown', value: '读取失败' }); }
  const available = clients.filter(client => client.detected && !client.manual);
  const invalid = available.filter(client => client.mcp.status !== 'valid');
  const skills = available.filter(client => client.skills.status !== 'unsupported');
  const unlinked = skills.filter(client => client.skills.status !== 'linked');
  if (clients.length) {
    checks.push({ id: 'clients', title: 'AI 接入配置', status: invalid.length ? 'attention' : available.length ? 'ok' : 'idle', value: available.length ? `${available.length - invalid.length} / ${available.length} 项一致` : '尚未检测到客户端', action: 'agents', detail: '检查的是本机配置；是否已加载，以各客户端的连接状态为准。' });
    checks.push({ id: 'skills', title: '技能链接', status: unlinked.length ? 'attention' : skills.length ? 'ok' : 'idle', value: skills.length ? `${skills.length - unlinked.length} / ${skills.length} 项已链接` : '尚未配置', action: 'agents' });
  }
  return { checkedAt: Date.now(), checks, attention: checks.filter(check => check.status === 'attention').length,
    clients: clients.filter(client => client.detected).map(client => ({ id: client.id, name: client.name, mcp: client.mcp.detail, skills: client.skills.detail })),
  };
}
