/**
 * 契约一致性校验（改造计划 P2.5）。
 *
 * 目标：让两类漂移**被检查发现**，而不是靠人工核对——
 *   1. 「声明了但未实现」：能力声明或路由表里列了，却没有可调用工具或宿主实现；
 *   2. 「实现了但未注册」：网关有执行分支，却没有对应的对外工具 schema。
 *
 * 另覆盖审计归属：`clientName` 必须真正落到审计记录里（P2.2 会话语义补正时发现
 * stdio 路径的 clientName 曾被改错，当时没有自动化断言）。
 *
 * 结构类断言只校验"名字集合"，不比对实现文本，因此不会因函数内部改动而假失败。
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const home = fs.mkdtempSync(path.join(os.tmpdir(), 'bridge-contract-'));
process.env.WPS_BRIDGE_HOME = home;

const { getTools, capabilities, EXCEL_METHODS } = await import('../src/bridge/catalog.js');
const { executeCatalogTool } = await import('../src/bridge/catalog.js');
const { bridgeServer } = await import('../src/bridge/ws-server.js');
const { requestContext } = await import('../src/bridge/context.js');
const { auditStore } = await import('../src/bridge/audit-store.js');
const contracts = await import('../src/bridge/contracts/host-methods.js');

/**
 * 网关里的执行分支。
 *
 * 直接读**真实注册表**（`registeredGatewayBranches()`），不从源码正则抓 `case "..."`：
 * 分支已改为 `HANDLERS` 注册表，若继续抓文本，本测试就只是在校验注释，
 * 注释与真实执行映射脱节时不会失败——那是被削弱的检查，不是通过。
 */
function gatewayBranches(): Set<string> {
  return new Set(registeredGatewayBranches());
}

/** 不经网关 switch、由 catalog 本地处理的工具。 */
const LOCAL_TOOLS = new Set([
  'bridge_get_capabilities', 'bridge_diagnose',
  'wps_get_audit_history', 'wps_get_audit_record', 'wps_clear_audit_history'
]);

/**
 * 已知「实现了但未注册」的分支（债务台账，不是期望状态）。
 *
 * 这些分支在 gateway 里有实现，但 `getOpenAiTools()` 没给它们 schema，因此 AI 永远调不到——
 * 属"宿主能做但对外没有工具"，处置方式见计划 DP3（保持暂缓，P2 只登记）。
 * 台账必须只减不增：新增死分支会让测试失败。
 */
const KNOWN_UNREGISTERED_BRANCHES: readonly string[] = [
  // 2026-09-22：wps_clear_range / wps_get_style_token / wps_reload_addon / wps_word_capture_preview
  // 已补齐 schema 并注册（原为"宿主能做但对外没有工具"）。台账按约定只减不增。
  // 剩余两项是**有意不暴露**，不是遗漏：
  //   wps_eval_code —— 与 wps_execute_script 重复，通用入口已覆盖，暴露只会多一条无校验的代码执行路径；
  //   wps_ppt_add_chart —— PPT 原生图表在本机 WPS 上 AddChart 返回 null 且不建形状（ISS-80），
  //     暴露会得到"必定失败"的工具；上游改用 insert_native_chart（报错明确）与矢量形状绘制。
  'wps_eval_code', 'wps_ppt_add_chart'
];

const { toolClassOf } = await import('../src/bridge/tools/index.js');
const { registeredGatewayBranches } = await import('../src/bridge/gateway.js');
const { normalizeOfficeRequest, normalizeOfficeResponse } = await import('../src/bridge/office/normalizer.js');

/** 解析加载项 RPC 分发里的方法名 —— 这是**宿主侧真实实现**，不是同源清单比对。 */
function addonMethods(relPath: string): Set<string> {
  const src = fs.readFileSync(path.resolve(import.meta.dirname, '..', relPath), 'utf8');
  return new Set([...src.matchAll(/case "([a-z_0-9]+)":/g)].map(m => m[1]));
}

const TOOL_NAMES = getTools().map((t: any) => t.name);
const REGISTERED = new Set(TOOL_NAMES);
const BRANCHES = gatewayBranches();

/**
 * 覆盖范围声明（不要让本测试的名称超出它真正证明的东西）：
 * 本项只证明"每个对外工具名都能匹配到一个网关 `case`"。它**不**验证：
 *   - 宿主加载项是否真有对应 RPC 分支（宿主实现被删掉，本项仍会通过）；
 *   - 参数转换、返回值形状与错误语义是否正确。
 * 宿主侧实现漂移目前只有静态核对（`HOST_IMPLEMENTATION_GAPS`）与实机验收两条路，
 * 前者是同源清单比对、不能独立发现漂移，需在 P2.3 描述符注册里补齐。
 */
test('注册工具缺少网关分支必须为空（不覆盖宿主实现与参数转换）', () => {
  const missing = TOOL_NAMES
    .filter(name => !LOCAL_TOOLS.has(name) && !BRANCHES.has(name) && !BRANCHES.has(name.replace(/^excel_/, 'wps_')))
    .sort();
  assert.deepEqual(missing, [], `以下工具已注册但找不到网关执行分支（excel_* 会重写为 wps_*）：\n${missing.join('\n')}`);
});

test('「实现了但未注册」与台账一致，不得新增死分支', () => {
  const unregistered = [...BRANCHES].filter(branch => !REGISTERED.has(branch)).sort();
  const known = [...KNOWN_UNREGISTERED_BRANCHES].sort();
  const added = unregistered.filter(b => !known.includes(b));
  const resolved = known.filter(b => !unregistered.includes(b));
  assert.deepEqual(added, [], `出现新的"有分支无 schema"工具，应注册 schema 或从网关删除：\n${added.join('\n')}`);
  assert.deepEqual(resolved, [], `以下分支已不再是死分支，请从台账删除（台账须反映真实状态）：\n${resolved.join('\n')}`);
});

test('能力声明与宿主实现缺口同源且一致', () => {
  const caps: any = capabilities();
  const toolNames = new Set(TOOL_NAMES);

  // 声明为可调用的方法，必须真的有 wps_/excel_ 两个入口
  for (const host of ['wps', 'microsoft'] as const) {
    for (const method of caps.hosts[host].excel.implemented) {
      assert.ok(toolNames.has(`wps_${method}`), `${host}: 声明 implemented 但没有 wps_${method}`);
      assert.ok(toolNames.has(`excel_${method}`), `${host}: 声明 implemented 但没有 excel_${method}`);
    }
  }

  // 声明的"声明了但没有工具"必须与路由表差集一致
  const declared = new Set(caps.hosts.wps.excel.declaredNotCallable);
  const expected = EXCEL_METHODS.filter(m => !toolNames.has(`wps_${m}`) || !toolNames.has(`excel_${m}`));
  assert.deepEqual([...declared].sort(), [...expected].sort(), 'declaredNotCallable 与路由表差集不一致');

  // 宿主实现缺口：WPS 侧剩 export_shape_image（CAP-21 补上 update_chart 后换用它作样例）；
  // 断言与契约层 HOST_IMPLEMENTATION_GAPS 同源
  assert.deepEqual(caps.hosts.wps.excel.unimplementedOnHost, [...contracts.HOST_IMPLEMENTATION_GAPS.wps]);
  assert.deepEqual(caps.hosts.microsoft.excel.unimplementedOnHost, [...contracts.HOST_IMPLEMENTATION_GAPS.microsoft]);
  assert.equal(caps.hosts.wps.excel.implemented.includes('export_shape_image'), false, 'WPS 未实现的方法不得计入 implemented');
  assert.equal(caps.hosts.wps.excel.implemented.includes('update_chart'), true, 'CAP-21 后 WPS 已实现 update_chart');
  assert.equal(caps.hosts.microsoft.excel.implemented.includes('update_chart'), true, 'Microsoft 有实现，必须计入');
});

test('P2.3 分类：每个工具都能落到唯一一类，且 Excel 类与宿主路由表一致', () => {
  const byClass: Record<string, number> = {};
  for (const name of TOOL_NAMES) {
    const cls = toolClassOf(name);
    byClass[cls] = (byClass[cls] || 0) + 1;
  }
  assert.equal(Object.values(byClass).reduce((a, b) => a + b, 0), TOOL_NAMES.length, '每个工具必须恰好归类一次');

  // Excel 类（统一入口 + wps_ 表格兼容名）必须与宿主方法路由表一一对应
  const excelBare = new Set(
    TOOL_NAMES.filter(n => toolClassOf(n) === 'excel').map(n => n.replace(/^(wps|excel)_/, ''))
  );
  const routed = new Set([...EXCEL_METHODS]);
  assert.deepEqual(
    [...excelBare].filter(m => !routed.has(m as any)).sort(), [],
    'Excel 类里有工具不在宿主方法路由表中'
  );
  assert.ok(byClass.excel > 0 && byClass.word > 0 && byClass.ppt > 0, `分类覆盖异常: ${JSON.stringify(byClass)}`);
});

test('P2.5 分类口径：统一入口与兼容名成对，Microsoft 类只含原生脚本/预览/状态', () => {
  const byClass: Record<string, string[]> = {};
  for (const name of TOOL_NAMES) (byClass[toolClassOf(name)] ||= []).push(name);

  const unified = TOOL_NAMES.filter(n => n.startsWith('excel_'));
  const compatTable = byClass.excel.filter(n => !n.startsWith('excel_'));
  assert.equal(unified.length, compatTable.length, '统一 excel_* 与 wps_ 表格兼容名必须成对出现');
  assert.equal(unified.length, EXCEL_METHODS.filter(m => REGISTERED.has(`wps_${m}`)).length,
    '统一入口数量必须等于"路由表中已注册为工具的宿主方法数"');
  for (const name of unified) {
    const compat = name.replace(/^excel_/, 'wps_');
    assert.ok(REGISTERED.has(compat), `${name} 缺少对应的兼容入口 ${compat}（已发布别名不得丢失）`);
  }

  // Microsoft 专有类只应是原生驱动通道的三个工具；Excel 结构化能力走统一入口 + host=microsoft
  assert.deepEqual([...byClass.microsoft].sort(),
    ['office_capture_slide_preview', 'office_execute_script', 'office_get_status'],
    'Microsoft 类工具清单发生变化，请确认是否有新的原生驱动工具或误分类');
});

test('P2.5 参数转换边界：Microsoft 支持的宿主方法都有可用的请求/响应转换', () => {
  const msMethods = (capabilities() as any).hosts.microsoft.excel.implemented as string[];
  // 防止空清单让本项变成"空跑通过"：Microsoft 声明可调用的方法数应显著大于 0。
  assert.ok(msMethods.length >= 20, `Microsoft 可调用方法数异常偏少（${msMethods.length}），本项可能已失去覆盖意义`);
  for (const method of msMethods) {
    // 需要覆盖"宿主已执行后响应转换失败"这条路径，因此两侧都要至少能跑通空载荷。
    const request = normalizeOfficeRequest(method, {});
    assert.equal(typeof request.method, 'string', `${method}: 请求转换必须产出方法名`);
    assert.ok(request.params && typeof request.params === 'object' && !Array.isArray(request.params),
      `${method}: 请求参数必须收敛为对象`);
    assert.doesNotThrow(() => normalizeOfficeResponse(method, { success: true }, {}),
      `${method}: 响应转换边界不得对合法成功响应抛错`);
  }
});

test('P2.5 宿主实现独立校验：路由表的方法必须在加载项里真有 RPC 分支', () => {
  // 关键：这里读的是**加载项源码**（宿主侧真实实现），不是 HOST_IMPLEMENTATION_GAPS 或 capabilities。
  // 删掉某个加载项分支时，本项必须失败——同源清单比对做不到这一点。
  const wpsAddon = addonMethods('wps-addon/src/dispatch.js');
  const declaredWpsGaps = new Set(contracts.HOST_IMPLEMENTATION_GAPS.wps);

  const missing = EXCEL_METHODS.filter(m => !wpsAddon.has(m));
  const undeclared = missing.filter(m => !declaredWpsGaps.has(m));
  assert.deepEqual(undeclared, [],
    `以下方法在路由表里，但 WPS 加载项没有 RPC 分支、也没声明为缺口：\n${undeclared.join('\n')}`);

  const staleGaps = [...declaredWpsGaps].filter(m => wpsAddon.has(m));
  assert.deepEqual(staleGaps, [],
    `以下方法已声明为"WPS 未实现"，但加载项里其实有分支，声明已过期：\n${staleGaps.join('\n')}`);

  // Microsoft 侧：声明无缺口，则每个可调用方法都必须在 Office.js 分发里有分支
  const msAddon = addonMethods('office-addon/src/rpc.js');
  const msMissing = (capabilities() as any).hosts.microsoft.excel.implemented.filter((m: string) => !msAddon.has(m));
  assert.deepEqual(msMissing, [],
    `以下方法声明可在 Microsoft 使用，但 Office.js 加载项没有对应分支：\n${msMissing.join('\n')}`);
});

test('只读标注与重放安全性与契约层同源', () => {
  for (const tool of getTools() as any[]) {
    assert.equal(tool.annotations.readOnlyHint, contracts.isReadOnlyTool(tool.name), `${tool.name} 的 readOnlyHint 与契约层不一致`);
  }
});

test('审计归属：clientName 真正落到审计记录', async () => {
  const original = bridgeServer.callWps;
  (bridgeServer as any).callWps = async (method: string) => ({
    success: true, sheetName: 'Sheet1', address: 'A1:B1', rowCount: 1, columnCount: 2,
    modifiedCount: 2, diff: [], workbookName: 'audit.xlsx', message: `stub ${method}`
  });
  try {
    const result: any = await requestContext.run({ sessionId: 'audit-session-a', host: 'wps' }, () =>
      executeCatalogTool('wps_patch_cells', { address: 'A1:B1', values: [[1, 2]], workbookName: 'audit.xlsx' }, '客户端A'));
    assert.ok(result.auditId, 'patch_cells 应返回 auditId');

    const record = auditStore.getRecordById(result.auditId);
    assert.ok(record, '审计记录应可读回');
    assert.equal((record as any).clientName, '客户端A', 'clientName 必须写入审计记录，不能被默认值替换');
    assert.equal((record as any).host, 'wps');
    assert.equal((record as any).workbookName, 'audit.xlsx');
  } finally {
    (bridgeServer as any).callWps = original;
  }
});

test('审计归属：会话隔离，记录归属各自客户端', async () => {
  const original = bridgeServer.callWps;
  (bridgeServer as any).callWps = async () => ({
    success: true, sheetName: 'Sheet1', address: 'C1', rowCount: 1, columnCount: 1,
    modifiedCount: 1, diff: [], workbookName: 'iso.xlsx'
  });
  try {
    const idA = (await requestContext.run({ sessionId: 'iso-a', host: 'wps' }, () =>
      executeCatalogTool('wps_patch_cells', { address: 'C1', values: [[1]], workbookName: 'iso.xlsx' }, 'client-A')) as any).auditId;
    const idB = (await requestContext.run({ sessionId: 'iso-b', host: 'wps' }, () =>
      executeCatalogTool('wps_patch_cells', { address: 'C1', values: [[2]], workbookName: 'iso.xlsx' }, 'client-B')) as any).auditId;

    assert.notEqual(idA, idB, '两次调用应产生不同审计记录');
    assert.equal((auditStore.getRecordById(idA) as any).clientName, 'client-A');
    assert.equal((auditStore.getRecordById(idB) as any).clientName, 'client-B', '不同客户端的审计归属不得互相覆盖');
  } finally {
    (bridgeServer as any).callWps = original;
  }
});
