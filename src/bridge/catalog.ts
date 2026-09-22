import { UniversalGateway } from './gateway.js';
import { assembleTools } from './tools/index.js';
import { currentHost } from './context.js';
import { bridgeServer } from './ws-server.js';
import { VERSION } from './runtime.js';
import { BridgeError } from './errors.js';
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
    audit: { covered: ['patch_cells 值与公式'], uncovered: ['样式', '图表', '结构修改', 'Word/PPT', '原生脚本'], rollback: '检查当前值和公式与记录的修改后快照一致；有后续修改则拒绝覆盖' },
    lifecycle: '后台进程独立于窗口。stdio 自动启动或复用；HTTP 客户端需先启动 Bridge。停止服务会断开所有客户端。',
    repairs: ['启动后台', '检查版本与端口', '重新部署本项目加载项', '仅合并用户选择的客户端配置'],
    limitations: ['不承诺未实测的平台可用', '写入超时结果未知，先读回，禁止自动重放', 'MCP 已连接不等于办公软件已连接', '持久保存必须有明确保存结果']
  };
}
/**
 * 对外工具清单（P2.3）：定义已按类迁移到 `tools/`，本函数只做装配调用，保留原有导出名与顺序。
 * 注意诊断工具的顺序：`bridge_diagnose` 在 `bridge_get_capabilities` 之前，与迁移前一致。
 */
export function getTools() {
  return assembleTools();
}

export function validateArgs(schema: any, value: any, location = 'arguments'): void {
  if (!schema) return;
  if (schema.oneOf) {
    const matches = schema.oneOf.filter((part: any) => { try { validateArgs(part, value, location); return true; } catch { return false; } });
    if (matches.length !== 1) throw new Error(`${location}: 类型不匹配`);
  }
  if (schema.enum && !schema.enum.includes(value)) throw new Error(`${location}: 不在允许值中`);
  const types = Array.isArray(schema.type) ? schema.type : [schema.type];
  if (schema.type && !types.some((type: string) => type === 'null' ? value === null : type === 'array' ? Array.isArray(value) : type === 'object' ? value !== null && typeof value === 'object' && !Array.isArray(value) : type === 'integer' ? Number.isInteger(value) : typeof value === type)) throw new Error(`${location}: 类型不正确`);
  if (typeof value === 'number' && !Number.isFinite(value)) throw new Error(`${location}: 数字无效`);
  if (Array.isArray(value)) value.forEach((v, i) => validateArgs(schema.items, v, `${location}[${i}]`));
  else if (value && typeof value === 'object') {
    for (const key of schema.required || []) if (value[key] === undefined) throw new Error(`${location}.${key}: 必填`);
    for (const key of Object.keys(value)) {
      if (schema.additionalProperties === false && !schema.properties?.[key]) throw new Error(`${location}.${key}: 未知参数`);
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
  try {
    validateArgs(tool.inputSchema, args);
  } catch (error: any) {
    // 入参校验在调用宿主之前完成，可确认宿主未执行。
    throw new BridgeError('rejected', error?.message || String(error), { channel: 'bridge', method: name, executed: 'no', cause: error });
  }
  if (name === 'bridge_get_capabilities') return capabilities();
  if (name === 'bridge_diagnose') return { ...capabilities(), checks: [{ id: 'service', status: 'ok' }, ...Object.entries(bridgeServer.getState().components).map(([id, s]) => ({ id: `wps-${id}`, status: s.connected ? 'connected' : 'not-connected', next: s.connected ? '先读取目标文档' : '打开目标办公组件，并检查加载项是否安装、版本是否匹配' }))] };
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

  args = structuredClone(args);
  if (realName === 'wps_patch_cells') {
    for (const field of ['values', 'formulas']) if (args[field] != null && (!Array.isArray(args[field]) || args[field].length === 0 || args[field].some((row: any) => !Array.isArray(row) || row.length !== args[field][0].length))) throw new BridgeError('rejected', `${field} 必须为非空矩形二维数组`, { channel: 'bridge', method: name, executed: 'no' });
  }
  const result = await UniversalGateway.executeTool(realName, args, clientName);
  // 宿主已接收请求并明确报告失败：是否已部分生效未知，不得当作"未执行"重放。
  if (result?.success === false) throw new BridgeError('failed', result.error || result.message || '办公软件未完成请求', { channel: 'bridge', method: name });
  return result;
}
