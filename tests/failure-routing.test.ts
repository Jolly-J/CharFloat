/**
 * 失败分类与跨通道回退的故障注入测试（改造计划 P1.2 / P1.3 / P1.4）。
 *
 * 覆盖的真实风险：Office.js 通道在执行**已经发生之后**才失败（响应超时、断连或响应无法解析），
 * 旧实现会把这类失败也当作"可以直接回退"，从而在 Windows 原生 COM 通道上重复执行同一次写入。
 *
 * 本文件用桩替换通道边界（bridgeServer.callOfficeAddon / MsOfficeDriver.windows）来统计
 * "宿主实际执行次数"和"原生通道回退次数"，因此可以断言写入只发生一次。
 * 不依赖真实 Office 宿主，也不修改任何用户文档。
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  BridgeError, isReadOnlyTool, isReplaySafeMethod, mayReplayOnAnotherChannel, routeOfficeFailure
} from '../src/bridge/errors.js';
import { bridgeServer } from '../src/bridge/ws-server.js';
import { requestContext } from '../src/bridge/context.js';
import { getTools } from '../src/bridge/catalog.js';
import { MsOfficeDriver } from '../src/bridge/office/ms-office-driver.js';
import { callOffice } from '../src/bridge/office/adapter.js';

const ROUTING = { platform: 'win32' as NodeJS.Platform, supportedMethods: ['patch_cells', 'read_range', 'find_and_replace', 'inspect_api'] };

function failure(kind: 'unavailable' | 'rejected' | 'failed' | 'unknown', method = 'patch_cells', message?: string) {
  return new BridgeError(kind, message ?? `${kind} 注入`, { channel: 'microsoft-officejs', method });
}

test('四种失败种类默认的"是否已执行"判定是保守的', () => {
  assert.equal(failure('unavailable').executed, 'no', '通道未连接可确认未执行');
  assert.equal(failure('rejected').executed, 'no', '入参拒绝可确认未执行');
  assert.equal(failure('failed').executed, 'unknown', '宿主报错不得当作未执行');
  assert.equal(failure('unknown').executed, 'unknown', '超时/断连不得当作未执行');
  assert.equal(failure('unknown').kind, 'unknown');
});

test('只有确认未执行或不修改文档的方法才允许跨通道重放', () => {
  assert.equal(isReplaySafeMethod('read_range'), true);
  assert.equal(isReplaySafeMethod('wps_read_range'), true, '带兼容前缀的写法应归一');
  assert.equal(isReplaySafeMethod('ppt_get_slide_shapes'), true);
  assert.equal(isReplaySafeMethod('patch_cells'), false, '写入默认不可重放');
  assert.equal(isReplaySafeMethod('execute_script'), false, '任意脚本默认不可重放');
  assert.equal(isReplaySafeMethod('未知方法'), false, '未知方法按可能写文档处理');

  assert.equal(mayReplayOnAnotherChannel('patch_cells', failure('unavailable')), true, '确认未执行可回退');
  assert.equal(mayReplayOnAnotherChannel('patch_cells', failure('rejected')), true, '拒绝且未执行可回退');
  assert.equal(mayReplayOnAnotherChannel('patch_cells', failure('failed')), false, '宿主已报错不可回退');
  assert.equal(mayReplayOnAnotherChannel('patch_cells', failure('unknown')), false, '结果未知不可回退');
  assert.equal(mayReplayOnAnotherChannel('read_range', failure('unknown')), true, '读取重放不改文档，允许回退');
});

test('routeOfficeFailure 决策表：写操作在结果不可判定时拒绝回退', () => {
  for (const kind of ['failed', 'unknown'] as const) {
    const route = routeOfficeFailure('patch_cells', failure(kind), ROUTING);
    assert.equal(route.action, 'reject', `${kind} 的写法不应回退`);
    if (route.action === 'reject') {
      assert.match(route.error.message, /已阻止自动改用 Windows 原生 COM 通道重放/);
      assert.match(route.error.message, /先读回/, '必须给出读回指引');
      assert.equal(route.error.kind, kind, '保留原始失败种类');
    }
  }
  for (const kind of ['unavailable', 'rejected'] as const) {
    assert.equal(routeOfficeFailure('patch_cells', failure(kind), ROUTING).action, 'fallback', `${kind} 可回退`);
  }
  assert.equal(routeOfficeFailure('read_range', failure('unknown'), ROUTING).action, 'fallback', '读取可回退');
  assert.equal(routeOfficeFailure('patch_cells', failure('unknown'), { ...ROUTING, platform: 'darwin' as NodeJS.Platform }).action, 'reject', '非 Windows 不进入原生通道');
});

test('替代通道不支持该方法时，保留前一通道的原始错误与 executed 状态', () => {
  const unsupported = { platform: 'win32' as NodeJS.Platform, supportedMethods: ['patch_cells'] };
  const timeout = failure('unknown', 'ppt_generate_deck', 'Microsoft Office ppt_generate_deck 超时，请检查 Excel 任务窗格连接');
  const route = routeOfficeFailure('ppt_generate_deck', timeout, unsupported);
  assert.equal(route.action, 'reject');
  if (route.action === 'reject') {
    assert.equal(route.error.kind, 'unknown', '不得把结果未知改写成"明确拒绝"');
    assert.equal(route.error.executed, 'unknown', '关键断言：不得把"可能已执行"改写成"未执行"');
    assert.equal(route.error.message.startsWith(timeout.message), true, '必须保留前一通道的原始错误文本');
    assert.match(route.error.message, /不在 Windows 原生通道支持列表内/);
    assert.match(route.error.message, /bridge_get_capabilities/);
    assert.equal(route.error.method, 'ppt_generate_deck');
  }

  // 前一通道确认未执行时，no 也要原样保留
  const unavailable = failure('unavailable', 'ppt_generate_deck', '加载项未连接');
  const route2 = routeOfficeFailure('ppt_generate_deck', unavailable, unsupported);
  if (route2.action === 'reject') {
    assert.equal(route2.error.kind, 'unavailable');
    assert.equal(route2.error.executed, 'no');
  } else assert.fail('应当拒绝');

  // 已执行后报错的情形同样保留 failed / unknown
  const reported = failure('failed', 'ppt_generate_deck', '宿主返回失败');
  const route3 = routeOfficeFailure('ppt_generate_deck', reported, unsupported);
  if (route3.action === 'reject') {
    assert.equal(route3.error.kind, 'failed');
    assert.equal(route3.error.executed, 'unknown');
  } else assert.fail('应当拒绝');
});

/** 在 win32 语义下运行一段使用桩通道的代码，结束后恢复平台与桩。 */
async function withWin32Stubs<T>(stubs: { addon: (method: string, params: any) => Promise<any>; native: () => Promise<any> }, run: () => Promise<T>) {
  const originalPlatform = Object.getOwnPropertyDescriptor(process, 'platform');
  const originalAddon = bridgeServer.callOfficeAddon;
  const originalNative = MsOfficeDriver.windows;
  Object.defineProperty(process, 'platform', { value: 'win32', configurable: true });
  (bridgeServer as any).callOfficeAddon = stubs.addon;
  (MsOfficeDriver as any).windows = stubs.native;
  try {
    return await run();
  } finally {
    if (originalPlatform) Object.defineProperty(process, 'platform', originalPlatform);
    (bridgeServer as any).callOfficeAddon = originalAddon;
    (MsOfficeDriver as any).windows = originalNative;
  }
}

test('故障注入：写入已发生但响应超时，不得在原生通道重复执行', async () => {
  let hostExecutions = 0;
  let nativeFallbacks = 0;
  const stats = { get hostExecutions() { return hostExecutions; }, get nativeFallbacks() { return nativeFallbacks; } };

  const error = await withWin32Stubs({
    addon: async (method: string) => {
      // 宿主已经执行了写入，只是响应没有回来。
      hostExecutions += 1;
      throw new BridgeError('unknown', `Microsoft Office ${method} 超时，请检查 Excel 任务窗格连接`, { channel: 'microsoft-officejs', method });
    },
    native: async () => { nativeFallbacks += 1; return { success: true }; }
  }, async () => {
    try {
      await requestContext.run({ sessionId: 'fault-timeout', host: 'microsoft' }, () =>
        callOffice('patch_cells', { address: 'A1:B2', values: [[1, 2], [3, 4]], workbookName: 'fault.xlsx' }));
      return null;
    } catch (e) { return e as BridgeError; }
  });

  assert.ok(error, '必须抛出错误而不是静默成功');
  assert.equal(error!.kind, 'unknown', '超时应归类为结果未知');
  assert.equal(stats.hostExecutions, 1, '宿主只应被调用一次');
  assert.equal(stats.nativeFallbacks, 0, '关键断言：不得改用原生通道重复执行同一次写入');
  assert.match(error!.message, /先读回/);
});

test('故障注入：写入已完成但响应无法解析，同样不得回退重放', async () => {
  let hostExecutions = 0;
  let nativeFallbacks = 0;
  // rowCount 取值即抛错：模拟宿主返回了无法解析的响应，而写入本身已经落盘。
  const poisoned = { get rowCount(): number { throw new Error('响应格式无法解析'); } };

  const error = await withWin32Stubs({
    addon: async () => { hostExecutions += 1; return poisoned; },
    native: async () => { nativeFallbacks += 1; return { success: true }; }
  }, async () => {
    try {
      await requestContext.run({ sessionId: 'fault-convert', host: 'microsoft' }, () =>
        callOffice('patch_cells', { address: 'A1:B2', values: [[1, 2], [3, 4]], workbookName: 'fault.xlsx' }));
      return null;
    } catch (e) { return e as BridgeError; }
  });

  assert.ok(error, '必须抛出错误而不是静默成功');
  assert.equal(error!.kind, 'failed', '响应转换失败发生在宿主执行之后，归类为执行失败');
  assert.equal(error!.message, '响应格式无法解析', '保留原始失败原因');
  assert.equal(hostExecutions, 1, '宿主只应被调用一次');
  assert.equal(nativeFallbacks, 0, '关键断言：响应转换失败不得触发原生通道重复写入');
});

test('故障注入：加载项未连接属于执行前不可用，仍允许原生通道接管', async () => {
  let nativeFallbacks = 0;
  const result = await withWin32Stubs({
    addon: async (method: string) => {
      throw new BridgeError('unavailable', `Microsoft Office (ms-excel) 加载项未连接`, { channel: 'microsoft-officejs', method, executed: 'no' });
    },
    native: async () => { nativeFallbacks += 1; return { success: true, via: 'native' }; }
  }, () => requestContext.run({ sessionId: 'fault-unavailable', host: 'microsoft' }, () =>
    callOffice('patch_cells', { address: 'A1:B2', values: [[1, 2], [3, 4]], workbookName: 'fault.xlsx' })));

  assert.equal(nativeFallbacks, 1, '确认未执行时仍应保留原生通道回退能力');
  assert.deepEqual(result, { success: true, via: 'native' });
});

test('通道边界在真实调用路径上抛出已分类错误', async () => {
  // 未启动服务、无加载项连接：属于执行前不可用，且可确认宿主未执行。
  const error = await bridgeServer.callWps('read_range', { address: 'A1' }, 200).then(() => null, (e: unknown) => e as BridgeError);
  assert.ok(error instanceof BridgeError, '通道边界必须抛出 BridgeError');
  assert.equal(error!.kind, 'unavailable');
  assert.equal(error!.executed, 'no');
  assert.equal(error!.channel, 'wps-addon');
  assert.equal(error!.method, 'read_range');
  assert.equal(JSON.parse(JSON.stringify(error)).kind, 'unavailable', '结构化描述可用于日志，且不含文档内容');
});

test('MCP readOnlyHint 与重放安全性是两套命名清晰的判定', () => {
  assert.equal(isReadOnlyTool('bridge_get_capabilities'), true);
  assert.equal(isReadOnlyTool('office_get_status'), true);
  assert.equal(isReadOnlyTool('excel_read_range'), true);
  assert.equal(isReadOnlyTool('wps_get_audit_history'), true);
  assert.equal(isReadOnlyTool('wps_get_style_token'), true, '取值不写文档，应标为只读');
  assert.equal(isReadOnlyTool('excel_patch_cells'), false);
  assert.equal(isReadOnlyTool('wps_capture_sheet_preview'), false, '截图会写本地文件，不属于 MCP 只读');
  assert.equal(isReplaySafeMethod('capture_sheet_preview'), false, '截图未列入重放白名单，保守处理');
});

test('带副作用的表达式探测既不算只读，也不允许跨通道重放', () => {
  // inspect_api 接收调用方给出的任意表达式，无法静态保证只读。
  // 会改文档的表达式示例（真实可传给 inspect_api 的写法）：
  const sideEffectingExpressions = [
    'app.ActiveWorkbook.Worksheets.Add()',      // 新建工作表
    'wb.Worksheets.Item(1).Delete()',           // 删除工作表
    'pres.Slides.Add(1, 12)',                   // 新建幻灯片
    'doc.Content.InsertAfter("x")'              // 写入正文
  ];
  assert.equal(sideEffectingExpressions.length > 0, true);

  assert.equal(isReadOnlyTool('wps_inspect_api'), false, '表达式可能改文档，不得标为只读');
  assert.equal(isReplaySafeMethod('inspect_api'), false, '表达式可能改文档，不得跨通道重放');
  assert.equal(isReplaySafeMethod('wps_inspect_api'), false, '带兼容前缀同样处理');

  // 前一通道结果不可判定时不回退：否则副作用表达式会在另一通道再执行一次
  const route = routeOfficeFailure('inspect_api', failure('unknown', 'inspect_api', '探测超时'), ROUTING);
  assert.equal(route.action, 'reject', '结果未知的表达式探测不得重放');
  if (route.action === 'reject') {
    assert.equal(route.error.executed, 'unknown');
    assert.match(route.error.message, /已阻止自动改用 Windows 原生 COM 通道重放/);
  }

  // 确认未执行时仍允许回退（不因收紧而损失正常能力）
  assert.equal(routeOfficeFailure('inspect_api', failure('unavailable', 'inspect_api'), ROUTING).action, 'fallback');

  // 工具快照层面同步：wps_inspect_api 的 readOnlyHint 必须为 false
  const inspectTool = getTools().find(t => t.name === 'wps_inspect_api');
  assert.ok(inspectTool, 'wps_inspect_api 应仍注册（只改标注，不删工具）');
  assert.equal(inspectTool!.annotations?.readOnlyHint, false);
});

test('DP4：声明了但目标宿主未实现的操作，在调用宿主之前明确拒绝', async () => {
  const { executeCatalogTool, capabilities } = await import('../src/bridge/catalog.js');

  // host=wps（wps_* 兼容名固定指向 WPS）
  // 关键：此刻没有任何加载项连接。若检查发生在宿主调用之后，错误会是 unavailable；
  // 得到 rejected 说明确实在调用宿主之前就拒绝了。
  // 样例用 export_shape_image：CAP-21 补上 update_chart 后，它是 WPS 侧仍在缺口表里的方法。
  const wpsError = await requestContext.run({ sessionId: 'dp4-wps', host: 'wps' }, () =>
    executeCatalogTool('wps_export_shape_image', { workbookName: 'x.xlsx', sheetName: 'S1', shapeName: 'Shape 1' }, '测试')
      .then(() => null, (e: unknown) => e as BridgeError));

  assert.ok(wpsError instanceof BridgeError, '应抛出已分类错误');
  assert.equal(wpsError!.kind, 'rejected', '属于明确拒绝');
  assert.equal(wpsError!.executed, 'no', '可确认宿主未执行');
  assert.match(wpsError!.message, /未实现/, '错误文本应说明未实现');
  assert.match(wpsError!.message, /wps_execute_script/, '应给出可行的替代路径');

  // 同一能力声明与行为一致
  const caps: any = capabilities();
  // 2026-09-22：CAP-07 之后 WPS 表格侧已实现形状增删改查/分组/层级，
  // 仍缺的是形状导图（export_shape_image）与 MS 独有的 get_active_shape。
  assert.deepEqual(caps.hosts.wps.excel.unimplementedOnHost, ['export_shape_image']);

  // Microsoft 侧有实现，不得被这条前置检查误伤：未连接时仍是 unavailable（说明走到了宿主调用）
  const msError = await requestContext.run({ sessionId: 'dp4-ms', host: 'microsoft' }, () =>
    executeCatalogTool('excel_update_chart', { host: 'microsoft', workbookName: 'x.xlsx', chartName: 'Chart 1', title: '新标题' }, '测试')
      .then(() => null, (e: unknown) => e as BridgeError));
  assert.ok(msError instanceof BridgeError);
  assert.equal(msError!.kind, 'unavailable', 'Microsoft 侧不得被误判为未实现，应表现为加载项未连接');
});

test('CAP-50：参数安全护栏（空搜索串）不得换通道重放，能力缺口仍可回退', async () => {
  // 破坏链：schema 允许 searchQuery=''（空串通过 JSON-schema）→ 桥接在调用宿主前判为破坏性并阻断。
  // 如果这次"执行前拒绝"被当成普通拒绝而转交 Windows 原生 COM，excel.ps1 的 find_and_replace
  // 会用 `IndexOf('')` 命中整片区域、并按空正则逐格替换 —— 绕过护栏，把真实数据改坏。
  const unsafe = new BridgeError('rejected', '已阻断不安全的参数组合（空搜索串的查找替换）', {
    channel: 'microsoft-officejs', method: 'find_and_replace', executed: 'no', unsafe: true
  });
  const route = routeOfficeFailure('find_and_replace', unsafe, ROUTING);
  assert.equal(route.action, 'reject', 'find_and_replace 在 COM 白名单里，但安全护栏拒绝不得回退');
  if (route.action === 'reject') {
    assert.equal(route.error.kind, 'rejected');
    assert.equal(route.error.executed, 'no');
    assert.equal(route.error.unsafe, true, '安全护栏标记必须保留，否则日志与后续判定会看不出来');
    assert.match(route.error.message, /参数安全护栏/);
    assert.match(route.error.message, /空搜索串/);
  }
  // 反向：同方法的普通"执行前拒绝"（能力/入参类）仍可回退，不因收紧而损失正常回退能力
  assert.equal(routeOfficeFailure('find_and_replace', failure('rejected', 'find_and_replace'), ROUTING).action, 'fallback');

  // 真实调用路径：空搜索串的查找替换在 win32 语义下既不碰 Office.js 宿主，也不碰原生 COM
  let addonCalls = 0;
  let nativeCalls = 0;
  const error = await withWin32Stubs({
    addon: async () => { addonCalls += 1; return { success: true }; },
    native: async () => { nativeCalls += 1; return { success: true }; }
  }, async () => {
    try {
      await requestContext.run({ sessionId: 'cap50-empty-search', host: 'microsoft' }, () =>
        callOffice('find_and_replace', { searchQuery: '', replaceText: 'X', workbookName: 'cap50.xlsx' }));
      return null;
    } catch (e) { return e as BridgeError; }
  });
  assert.ok(error, '必须抛出错误而不是静默成功');
  assert.equal(error!.kind, 'rejected', '属于执行前拒绝');
  assert.equal(error!.executed, 'no');
  assert.equal(addonCalls, 0, '不得把已判定为破坏性的请求交给 Office.js 宿主');
  assert.equal(nativeCalls, 0, '关键断言：不得转手 Windows 原生 COM（那边没有等价护栏）');

  // 能力缺口类的执行前拒绝仍应保留原生通道接管能力（COM 的 set_filter_and_sort 支持"只开筛选按钮"）
  let nativeFallbacks = 0;
  const fallback = await withWin32Stubs({
    addon: async () => { throw new Error('能力缺口应在参数转换阶段被拒绝，不应走到宿主'); },
    native: async () => { nativeFallbacks += 1; return { success: true, via: 'native' }; }
  }, () => requestContext.run({ sessionId: 'cap50-capability-gap', host: 'microsoft' }, () =>
    callOffice('set_filter_and_sort', { range: 'A1:E20', enableAutoFilter: true })));
  assert.equal(nativeFallbacks, 1, 'Office.js 无法只开筛选按钮，应仍由 COM 接管（COM 脚本支持该参数）');
  assert.deepEqual(fallback, { success: true, via: 'native' });
});

test('ISS-126：残留目标锁失效时自动清理并给出可操作的重新锁定提示', async () => {
  const { executeCatalogTool } = await import('../src/bridge/catalog.js');
  const { TargetLockStore } = await import('../src/bridge/gateway.js');
  const original = bridgeServer.callWps;
  const sessionId = 'iss126-stale-lock';
  const stale = () => new BridgeError(
    'failed',
    '未在 WPS 中找到目标 Word 文档 [agent-word.docx]。当前已打开: 测试文字文稿.docx',
    { channel: 'wps-addon', method: 'word_read_document' }
  );
  (bridgeServer as any).callWps = async () => { throw stale(); };
  try {
    // 场景一：目标由**会话锁注入**（宿主重启后锁指向已不存在的文档）
    await requestContext.run({ sessionId, host: 'wps' }, () => TargetLockStore.lock('word', 'agent-word.docx'));
    const error = await requestContext.run({ sessionId, host: 'wps' }, () =>
      executeCatalogTool('wps_word_read_document', {}, '测试').then(() => null, (e: unknown) => e as Error));
    assert.ok(error, '锁目标不存在时必须失败，不得静默返回空文档');
    assert.match(error!.message, /未在 WPS 中找到目标 Word 文档/, '必须保留宿主原始错误（含当前已打开列表）');
    assert.match(error!.message, /目标锁已失效并自动解除/, '必须点明是残留锁而不是工具坏了');
    assert.match(error!.message, /wps_lock_target_document/, '必须给出重新锁定的可操作指引');
    assert.match(error!.message, /wps_get_locked_status/, '必须给出查询当前打开的文档的入口');
    assert.equal((error as any).kind, 'failed', '保留原失败分类，调用方仍能判断能否重试');

    const locks = await requestContext.run({ sessionId, host: 'wps' }, () => TargetLockStore.getLocks());
    assert.deepEqual(locks, {}, '失效锁必须被自动清除，否则后续不带目标的调用继续被它带偏');

    // 场景二：调用方**显式传名**失败时不得动会话锁（不能把用户刚设好的锁误删）
    await requestContext.run({ sessionId, host: 'wps' }, () => TargetLockStore.lock('word', 'good.docx'));
    const explicit = await requestContext.run({ sessionId, host: 'wps' }, () =>
      executeCatalogTool('wps_word_read_document', { documentName: 'typo.docx' }, '测试').then(() => null, (e: unknown) => e as Error));
    assert.ok(explicit);
    assert.doesNotMatch(explicit!.message, /目标锁已失效并自动解除/, '显式传名失败不得改写成锁失效');
    const still = await requestContext.run({ sessionId, host: 'wps' }, () => TargetLockStore.getLocks());
    assert.deepEqual(still, { word: 'good.docx' }, '显式传错名字不得清掉会话锁');

    // 场景三：与目标无关的失败不得触发清理（例如样式/参数错误）
    (bridgeServer as any).callWps = async () => { throw new BridgeError('failed', '段落下标超出范围', { channel: 'wps-addon', method: 'word_read_document' }); };
    const other = await requestContext.run({ sessionId, host: 'wps' }, () =>
      executeCatalogTool('wps_word_read_document', {}, '测试').then(() => null, (e: unknown) => e as Error));
    assert.ok(other);
    assert.doesNotMatch(other!.message, /目标锁已失效并自动解除/);
    const kept = await requestContext.run({ sessionId, host: 'wps' }, () => TargetLockStore.getLocks());
    assert.deepEqual(kept, { word: 'good.docx' }, '只有"目标不存在"才清理锁，其他失败不得误清');
  } finally {
    (bridgeServer as any).callWps = original;
    await requestContext.run({ sessionId, host: 'wps' }, () => TargetLockStore.unlock());
  }
});
