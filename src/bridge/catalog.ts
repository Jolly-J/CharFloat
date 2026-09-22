import { UniversalGateway } from './gateway.js';
import { assembleTools } from './tools/index.js';
import { currentHost } from './context.js';
import { bridgeServer } from './ws-server.js';
import { VERSION } from './runtime.js';
import { BridgeError } from './errors.js';
import { collectBuildFingerprints } from './build-fingerprint.js';
import { EXCEL_METHODS, HOST_IMPLEMENTATION_GAPS, isReadOnlyTool, isUnimplementedOn, MCP_READ_ONLY_TOOLS } from './contracts/host-methods.js';

/**
 * 兼容导出：宿主方法标识与只读判定已上移到契约层 `contracts/host-methods.ts`（P2.1），
 * 这里保持原有导出名不变，既有调用方与脚本无需改动。契约层不反向依赖本模块，
 * 因此 `office/adapter.ts` 可以直接引用它而不形成回环。
 */
export { EXCEL_METHODS, HOST_IMPLEMENTATION_GAPS, isReadOnlyTool };
export const READ_TOOLS = MCP_READ_ONLY_TOOLS;

export function capabilities() {
  const registered = new Set(getTools().map(t => t.name));
  const callable = EXCEL_METHODS.filter(m => registered.has(`wps_${m}`) && registered.has(`excel_${m}`));
  const declaredNotCallable = EXCEL_METHODS.filter(m => !callable.includes(m));
  const implementedOn = (host: 'wps' | 'microsoft') => callable.filter(m => !HOST_IMPLEMENTATION_GAPS[host].includes(m));
  return {
    version: VERSION, platform: process.platform, preferredWorkflow: '在已打开的 Excel / WPS 表格中操作，先读后写，明确目标并核验结果',
    connection: bridgeServer.getState(),
    hosts: {
      wps: {
        transport: 'authenticated-addon',
        excel: {
          implemented: [...implementedOn('wps')],
          declaredNotCallable: [...declaredNotCallable],
          unimplementedOnHost: [...HOST_IMPLEMENTATION_GAPS.wps],
          validation: process.platform === 'darwin' ? '既有功能曾在 macOS 使用；2.0 回归需实机确认' : 'Windows 实机待验收'
        },
        word: '有限支持：读取、编辑、保存；按实际工具清单使用',
        ppt: '有限支持：基础页面、文字、形状、表格；复杂图表与母版需实机验证，不保证高保真复刻'
      },
      microsoft: {
        transport: 'Office.js Web Add-in / WebSocket',
        excel: {
          implemented: [...implementedOn('microsoft')],
          declaredNotCallable: [...declaredNotCallable],
          unimplementedOnHost: [...HOST_IMPLEMENTATION_GAPS.microsoft],
          validation: '代码路径已实现，但本候选版本尚未在真实 Microsoft Excel 上做实机验收；静态与模拟通道不代表实机通过'
        },
        word: '未提供 Microsoft 结构化通道：Word 工具为 WPS 专用；Microsoft Word 需改用 office_execute_script 原生脚本（Windows 分支待实机验收）',
        ppt: '未提供 Microsoft 结构化通道：PPT 工具为 WPS 专用；Microsoft PowerPoint 需改用 office_execute_script 原生脚本（Windows 分支待实机验收）'
      }

    },
    audit: {
      covered: ['patch_cells 的值与公式（可回滚）'],
      // ISS-48：其他写操作现在**会留痕但不可回滚**（rollbackable:false），不再"完全不进审计"。
      loggedNotRollbackable: ['样式/格式', '条件格式', '冻结窗格', '行列结构（含列宽自适应）', '图表增删改', '数据有效性', '筛选与排序', '批注', '工作表结构（新建/删除/克隆/改名等）', '批量查找替换', '数据透视表', '保存工作簿'],
      uncovered: ['Word/PPT 文档操作', '原生脚本内部的细粒度改动（只登记一次脚本执行）'],
      rollback: '回滚前依次校验：① 目标工作表仍存在；② 同区域没有更晚的未回滚记录；③ 记录之后没有发生过未审计的身份类操作（脚本/清表/工作表增删/行列增删/批量替换）；④ 区域内内容与记录的修改后快照一致。任一不满足即拒绝覆盖，并给出用 patch_cells 显式写回旧快照的替代路径。',
      rollbackHelp: 'wps_get_audit_history 的 summary 视图带 rollbackable 字段；wps_get_audit_record 可取回完整前后快照。'
    },
    targetResolution: {
      order: ['本次请求显式传入的 workbookName', '本会话目标锁 wps_lock_target_document（可能来自更早的会话）', '都不传时由宿主使用活动文稿 / 唯一打开的文稿'],
      note: 'wps_get_workspace_summary 的 hasOpenWorkbook 表示"宿主当前有没有打开的工作簿"，lockTarget.missing 才表示"本会话锁目标已不存在"；两者原来混在一个字段里（ISS-02）。wps_get_locked_status 会列出真实打开的文稿，可用于自救。'
    },
    sessionLifecycle: {
      mcp: 'HTTP /mcp 会话默认上限 64、空闲 15 分钟自动回收（WPS_BRIDGE_MCP_MAX_SESSIONS / WPS_BRIDGE_MCP_IDLE_MS 可调）；也可用 DELETE /mcp 立即释放。',
      hostComponents: '加载项连接断开时会记录掉线原因；非正常断开（心跳超时/无关闭帧）标记为"疑似宿主崩溃"并给出恢复指引（见 connection 的 suspectedHostCrash 与 components.*.lastDisconnect）。'
    },
    build: collectBuildFingerprints(),
    lifecycle: '后台进程独立于窗口。stdio 自动启动或复用；HTTP 客户端需先启动 Bridge。停止服务会断开所有客户端。',
    repairs: ['启动后台', '检查版本与端口', '重新部署本项目加载项', '仅合并用户选择的客户端配置'],
    serviceErrors: '后台不可用时的错误统一给出"服务名 + 地址 + 恢复动作"；连接类原始错误单列在 error 字段中（ISS-50）。',
    limitations: ['不承诺未实测的平台可用', '写入超时结果未知，先读回，禁止自动重放', 'MCP 已连接不等于办公软件已连接', '持久保存必须有明确保存结果', '保存是工作簿级操作：会连同其他会话/任务的在途改动一起落盘（ISS-16）']
  };
}
/**
 * 对外工具清单（P2.3）：定义已按类迁移到 `tools/`，本函数只做装配调用，保留原有导出名与顺序。
 * 注意诊断工具的顺序：`bridge_diagnose` 在 `bridge_get_capabilities` 之前，与迁移前一致。
 */
export function getTools() {
  return assembleTools();
}

/**
 * 统一目标参数（ISS-03）。
 *
 * 现状：`excel_*` / `wps_*` 里大多数工具接受 `host` + `workbookName` + `sheetName`，
 * 但 `wps_rollback` 只接受 `auditId`、`excel_save_workbook` 不接受 `sheetName`——
 * 调用方按同一套约定批量调用会直接 422。这里对"目标参数"做统一容忍：
 * 该工具不适用时**忽略但如实回报**（返回体里的 `ignoredParams`），不做静默吞掉。
 * 其他未知参数仍然报错，并在错误里列出该工具允许的完整参数集（ISS-03 / ISS-37）。
 */
const TOLERATED_TARGET_PARAMS = new Set(['host', 'workbookName', 'sheetName']);

/** 从入参里摘掉"本工具 schema 不认、但属统一目标参数"的键，返回被忽略的参数名。 */
function stripToleratedParams(toolName: string, schema: any, args: any): string[] {
  if (!/^(excel|wps)_/.test(toolName) || !args || typeof args !== 'object') return [];
  const ignored: string[] = [];
  for (const key of Object.keys(args)) {
    if (!TOLERATED_TARGET_PARAMS.has(key)) continue;
    if (schema?.properties?.[key]) continue;
    if (toolName === 'wps_lock_target_document') continue;
    delete args[key];
    ignored.push(key);
  }
  return ignored;
}

export function validateArgs(schema: any, value: any, location = 'arguments'): void {
  if (!schema) return;
  if (schema.oneOf) {
    const matches = schema.oneOf.filter((part: any) => { try { validateArgs(part, value, location); return true; } catch { return false; } });
    if (matches.length !== 1) throw new Error(`${location}: 类型不匹配`);
  }
  if (schema.enum && !schema.enum.includes(value)) throw new Error(`${location}: 不在允许值中（允许值：${schema.enum.join(' / ')}）`);
  const types = Array.isArray(schema.type) ? schema.type : [schema.type];
  if (schema.type && !types.some((type: string) => type === 'null' ? value === null : type === 'array' ? Array.isArray(value) : type === 'object' ? value !== null && typeof value === 'object' && !Array.isArray(value) : type === 'integer' ? Number.isInteger(value) : typeof value === type)) throw new Error(`${location}: 类型不正确（期望 ${types.join(' | ')}，实际 ${Array.isArray(value) ? 'array' : value === null ? 'null' : typeof value}）`);
  if (typeof value === 'number' && !Number.isFinite(value)) throw new Error(`${location}: 数字无效`);
  if (Array.isArray(value)) value.forEach((v, i) => validateArgs(schema.items, v, `${location}[${i}]`));
  else if (value && typeof value === 'object') {
    for (const key of schema.required || []) if (value[key] === undefined) throw new Error(`${location}.${key}: 必填`);
    const allowed = schema.properties ? Object.keys(schema.properties) : [];
    for (const key of Object.keys(value)) {
      // ISS-03 / ISS-37：未知参数必须列出该工具真正接受的参数集，否则调用方只能反复试探。
      if (schema.additionalProperties === false && !schema.properties?.[key]) {
        throw new Error(`${location}.${key}: 未知参数；本工具允许的参数：${allowed.length ? allowed.join(', ') : '（无）'}`);
      }
      validateArgs(schema.properties?.[key], value[key], `${location}.${key}`);
    }
  }
}
const executionQueues = new Map<string, Promise<any>>();
export async function executeCatalogTool(name: string, args: any, clientName: string) {
  const key = 'office';
  const job = (executionQueues.get(key) || Promise.resolve()).catch(() => {}).then(() => executeValidatedTool(name, args, clientName));
  executionQueues.set(key, job);
  try { return await job; } finally { if (executionQueues.get(key) === job) executionQueues.delete(key); }
}
async function executeValidatedTool(name: string, args: any, clientName: string) {
  const tool = getTools().find(t => t.name === name);
  if (!tool) throw new BridgeError('rejected', `未知工具 ${name}`, { channel: 'bridge', method: name, executed: 'no' });
  const working = args && typeof args === 'object' && !Array.isArray(args) ? { ...args } : {};
  // 统一目标参数：本工具 schema 不认的 host/workbookName/sheetName 忽略并回报，不做 422（ISS-03）。
  const ignoredParams = stripToleratedParams(name, tool.inputSchema, working);
  try {
    validateArgs(tool.inputSchema, working);
  } catch (error: any) {
    // 入参校验在调用宿主之前完成，可确认宿主未执行。
    throw new BridgeError('rejected', error?.message || String(error), { channel: 'bridge', method: name, executed: 'no', cause: error });
  }
  if (name === 'bridge_get_capabilities') return capabilities();
  if (name === 'bridge_diagnose') {
    const state = bridgeServer.getState();
    const build = state.build as any;
    return {
      ...capabilities(),
      checks: [
        { id: 'service', status: 'ok' },
        ...Object.entries(state.components).map(([id, s]: [string, any]) => ({
          id: `wps-${id}`,
          status: s.connected ? 'connected' : 'not-connected',
          next: s.connected ? '先读取目标文档' : '打开目标办公组件，并检查加载项是否安装、版本是否匹配',
          // ISS-90：非正常掉线时给出"疑似宿主崩溃"标记与恢复指引。
          ...(s.lastDisconnect ? { lastDisconnect: s.lastDisconnect, suspectedCrash: s.lastDisconnect.suspectCrash === true } : {}),
          // ISS-59：连接上报的版本与当前桥接版本的比对结论。
          ...(s.versionStatus ? { versionStatus: s.versionStatus } : {})
        })),
        {
          id: 'build-fingerprint',
          status: build?.wpsAddon?.deployedMatchesResource === false ? 'mismatch' : 'ok',
          // 只报告磁盘/部署副本是否一致；"进程里加载的是哪一份"无法证实（见 note）。
          detail: build ? {
            bridgeEntry: build.bridgeEntry?.sha256 || build.bridgeEntry?.error,
            wpsAddonResource: build.wpsAddon?.resource?.sha256 || build.wpsAddon?.resource?.error,
            wpsAddonDeployed: (build.wpsAddon?.deployed || []).map((d: any) => ({ path: d.path, sha256: d.sha256, mtime: d.mtime })),
            deployedMatchesResource: build.wpsAddon?.deployedMatchesResource,
            officeAddon: build.officeAddon?.resource?.sha256,
            note: build.note
          } : null,
          next: build?.wpsAddon?.deployedMatchesResource === false
            ? '已部署的 WPS 加载项副本与当前构建不一致：重新部署（npm run setup -- --addon）并在宿主里重新加载加载项，否则宿主里跑的可能仍是旧构建。'
            : '部署副本与当前构建一致（仍无法证明宿主进程里加载的就是该副本）。'
        },
        { id: 'mcp-sessions', status: 'ok', detail: state.mcpSessions },
        ...(state.suspectedHostCrash ? [{ id: 'host-crash', status: 'suspect', detail: state.suspectedHostCrash, next: state.suspectedHostCrash.guidance }] : [])
      ]
    };
  }
  const realName = name.replace(/^excel_/, 'wps_');

  // DP4：契约里声明了但目标宿主加载项没有实现的操作，在**调用宿主之前**明确拒绝。
  // 否则请求会落到加载项的 default 分支，报出与事实不符的"未知的 RPC 方法"。
  const host = currentHost();
  const bareMethod = realName.replace(/^wps_/, '');
  if (isUnimplementedOn(host, bareMethod)) {
    const hostLabel = host === 'wps' ? 'WPS' : 'Microsoft';
    throw new BridgeError(
      'rejected',
      `${name} 在 ${hostLabel} 宿主上未实现（加载项缺少该 RPC 分支），已拒绝且未执行。` +
      (host === 'wps' ? `如需修改已有图表，可用 wps_execute_script 调用原生 API；改用 host=microsoft 亦可。` : ''),
      { channel: 'bridge', method: name, executed: 'no' }
    );
  }

  const payload = structuredClone(working);
  if (realName === 'wps_patch_cells') {
    for (const field of ['values', 'formulas']) if (payload[field] != null && (!Array.isArray(payload[field]) || payload[field].length === 0 || payload[field].some((row: any) => !Array.isArray(row) || row.length !== payload[field][0].length))) throw new BridgeError('rejected', `${field} 必须为非空矩形二维数组`, { channel: 'bridge', method: name, executed: 'no' });
  }
  const result = await UniversalGateway.executeTool(realName, payload, clientName, { ignoredParams });
  // 宿主已接收请求并明确报告失败：是否已部分生效未知，不得当作"未执行"重放。
  if (result?.success === false) throw new BridgeError('failed', result.error || result.message || '办公软件未完成请求', { channel: 'bridge', method: name });
  // 被忽略的统一目标参数如实回报，避免"传了但没生效"变成静默行为（ISS-03）。
  if (ignoredParams.length && result && typeof result === 'object' && !Array.isArray(result)) {
    return {
      ...result,
      ignoredParams,
      ignoredParamsNote: `以下参数对 ${name} 不适用，已忽略：${ignoredParams.join(', ')}（本工具允许的参数见 schema）`
    };
  }
  return result;
}
