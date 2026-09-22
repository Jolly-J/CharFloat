/**
 * 契约层边界与依赖回环守卫（改造计划 P2.1 / P2.2）。
 *
 * 覆盖的风险：`catalog.ts → gateway.ts → office/adapter.ts → catalog.ts`、
 * `catalog.ts → gateway.ts → ws-server.ts → catalog.ts` 等回环曾真实存在，
 * 让"注册/协议层"与"执行层"互相依赖，依赖方向无法单独理解。
 *
 * 本测试只做**结构约束**校验（谁 import 谁），不比对实现文本、不检查行为细节。
 *
 * 完整性问题（评审指出）：用 DFS + "已访问即跳过"只能**发现**回环，无法**枚举**全部回环路径——
 * 先走到的一条路径会把中间节点标记为完成，另一条通往同一节点的路径就被跳过
 * （例如先发现 adapter→ws-server，就再也看不到 gateway→ws-server）。
 * 因此这里改用 **Tarjan 强连通分量（SCC）**：先把图分解为强连通分量，
 * 再列出每个非平凡分量内部的依赖边。这是完整且与遍历顺序无关的口径。
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const BRIDGE_DIR = path.resolve(import.meta.dirname, '../src/bridge');

/**
 * 允许存在的强连通分量（债务台账）。
 *
 * P2.2 已完成：协议层改为通过注入的 `ToolService` 取工具，`ws-server.ts` / `mcp-server.ts`
 * 不再静态引用 `catalog.ts`，原先那个含 5 个模块的强连通分量已分解，**当前无循环依赖**。
 *
 * 本数组必须保持为空。若将来出现环，测试会报出该分量的成员与内部依赖边；
 * 不允许把新环直接加进来"放行"。
 */
const KNOWN_CYCLIC_COMPONENTS: readonly string[] = [];

/** 收集 src/bridge 下所有 .ts 文件（相对 BRIDGE_DIR 的 POSIX 风格路径）。 */
function listSources(dir = BRIDGE_DIR, prefix = ''): string[] {
  const out: string[] = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.isDirectory()) out.push(...listSources(path.join(dir, entry.name), `${prefix}${entry.name}/`));
    else if (entry.name.endsWith('.ts')) out.push(`${prefix}${entry.name}`);
  }
  return out;
}

/** 解析文件里的静态 import/export-from，返回其指向的模块路径（相对 BRIDGE_DIR，POSIX 风格）。 */
function importsOf(relFile: string): string[] {
  const text = fs.readFileSync(path.join(BRIDGE_DIR, relFile), 'utf8');
  const specs: string[] = [];
  const patterns = [/^\s*import\s[^'"]*from\s*['"]([^'"]+)['"]/gm, /^\s*export\s[^'"]*from\s*['"]([^'"]+)['"]/gm];
  for (const pattern of patterns) for (const match of text.matchAll(pattern)) specs.push(match[1]);
  return specs
    .filter(spec => spec.startsWith('.'))
    .map(spec => path.posix.normalize(path.posix.join(path.posix.dirname(relFile), spec.replace(/\.js$/, '.ts'))));
}

const SOURCES = listSources();
const GRAPH = new Map(SOURCES.map(file => [file, importsOf(file).filter(dep => GRAPH_HAS(dep))]));
function GRAPH_HAS(dep: string) { return SOURCES.includes(dep); }

/** Tarjan 强连通分量（迭代实现，避免深图递归爆栈）。 */
function stronglyConnectedComponents(graph: Map<string, string[]>): string[][] {
  const index = new Map<string, number>();
  const low = new Map<string, number>();
  const onStack = new Set<string>();
  const stack: string[] = [];
  const components: string[][] = [];
  let counter = 0;

  for (const root of graph.keys()) {
    if (index.has(root)) continue;
    const work: { node: string; edge: number }[] = [{ node: root, edge: 0 }];
    index.set(root, counter); low.set(root, counter); counter++;
    stack.push(root); onStack.add(root);

    while (work.length) {
      const frame = work[work.length - 1];
      const edges = graph.get(frame.node) || [];
      if (frame.edge < edges.length) {
        const next = edges[frame.edge++];
        if (!index.has(next)) {
          index.set(next, counter); low.set(next, counter); counter++;
          stack.push(next); onStack.add(next);
          work.push({ node: next, edge: 0 });
        } else if (onStack.has(next)) {
          low.set(frame.node, Math.min(low.get(frame.node)!, index.get(next)!));
        }
      } else {
        work.pop();
        if (work.length) {
          const parent = work[work.length - 1].node;
          low.set(parent, Math.min(low.get(parent)!, low.get(frame.node)!));
        }
        if (low.get(frame.node) === index.get(frame.node)) {
          const component: string[] = [];
          let member: string;
          do { member = stack.pop()!; onStack.delete(member); component.push(member); } while (member !== frame.node);
          components.push(component);
        }
      }
    }
  }
  return components;
}

/** 把强连通分量渲染成稳定的台账条目：成员排序 + 分量内部边排序。 */
function describeComponent(component: string[]): string {
  const members = [...component].sort();
  const memberSet = new Set(members);
  const internalEdges: string[] = [];
  for (const node of members) {
    for (const dep of GRAPH.get(node) || []) if (memberSet.has(dep)) internalEdges.push(`${node}→${dep}`);
  }
  // 自环也算环：单成员分量只有在存在自引用时才是循环依赖。
  if (members.length === 1 && internalEdges.length === 0) return '';
  return `members=[${members.join(',')}] edges=[${[...internalEdges].sort().join(',')}]`;
}

function cyclicComponents(): string[] {
  return stronglyConnectedComponents(GRAPH)
    .map(describeComponent)
    .filter(Boolean)
    .sort();
}

const CYCLIC = cyclicComponents();

test('强连通分量台账只减不增，且反映真实状态', () => {
  const known = [...KNOWN_CYCLIC_COMPONENTS].sort();
  const added = CYCLIC.filter(c => !known.includes(c));
  const fixed = known.filter(c => !CYCLIC.includes(c));

  assert.deepEqual(added, [], `出现未登记的循环依赖分量，请先消除而不是加入台账：\n${added.join('\n')}`);
  assert.deepEqual(fixed, [], `下列循环依赖已不存在，请从 KNOWN_CYCLIC_COMPONENTS 中删除（台账必须反映真实状态）：\n${fixed.join('\n')}`);
});

test('P2.1 验收：执行层 adapter 不再反向依赖注册入口 catalog', () => {
  const imports = GRAPH.get('office/adapter.ts') || [];
  assert.ok(!imports.includes('catalog.ts'), 'office/adapter.ts 不应 import catalog.ts');
  assert.ok(imports.includes('contracts/host-methods.ts'), 'office/adapter.ts 应从契约层取宿主方法标识');
});

test('P2.2 验收：协议层通过注入的接口取工具，不静态反向引用 catalog', () => {
  for (const file of ['ws-server.ts', 'mcp-server.ts']) {
    const imports = GRAPH.get(file) || [];
    assert.ok(!imports.includes('catalog.ts'), `${file} 不应静态 import catalog.ts（应通过注入的 ToolService）`);
  }
});

test('P2.3 验收：分类模块不反向依赖注册入口 catalog', () => {
  for (const file of SOURCES.filter(f => f.startsWith('tools/'))) {
    const imports = GRAPH.get(file) || [];
    assert.ok(!imports.includes('catalog.ts'), `${file} 不应 import catalog.ts（装配方向为 catalog → tools）`);
  }
});

test('契约层不依赖执行层、协议层与服务实例', () => {
  const contractFiles = SOURCES.filter(file => file.startsWith('contracts/'));
  assert.ok(contractFiles.length > 0, '应存在 contracts/ 契约模块');
  const forbidden = new Set([
    'gateway.ts', 'catalog.ts', 'ws-server.ts', 'mcp-server.ts', 'service-client.ts',
    'office/adapter.ts', 'office/ms-office-driver.ts', 'runtime.ts', 'audit-store.ts', 'process-runner.ts', 'compose.ts'
  ]);
  for (const file of contractFiles) {
    for (const dep of GRAPH.get(file) || []) {
      assert.ok(!forbidden.has(dep), `契约层 ${file} 不应依赖 ${dep}`);
    }
  }
});

test('契约层标识与既有兼容导出同源', async () => {
  const contract = await import('../src/bridge/contracts/host-methods.js');
  const catalog = await import('../src/bridge/catalog.js');
  assert.equal(catalog.EXCEL_METHODS, contract.EXCEL_METHODS, 'catalog 的兼容导出应与契约层同一对象');
  assert.equal(catalog.READ_TOOLS, contract.MCP_READ_ONLY_TOOLS);
  assert.equal(catalog.HOST_IMPLEMENTATION_GAPS, contract.HOST_IMPLEMENTATION_GAPS);
  assert.equal(catalog.isReadOnlyTool, contract.isReadOnlyTool);
});

test('契约层判定在搬迁中保持不变（回归护栏）', async () => {
  const c = await import('../src/bridge/contracts/host-methods.js');
  // 2026-09-22 由 30 改 31：新增 get_style_token（读原表设计语言），它是**真实存在的宿主方法**——
  // wps-addon/src/dispatch.js 有 `get_style_token` RPC 分支、网关有 getStyleToken 处理器，
  // 此前只因没有 schema 而成为"实现了但未注册"的死分支（ISS-95）。
  // 本护栏的用意是"搬迁过程中不得顺手增减"，**有意扩展能力时按此格式写明理由再改**，不是禁止增长。
  // 2026-09-22 由 30 依次扩展：format_text_segment（富文本，CAP-03）、
  // configure_print_layout（打印与分页，CAP-01）、export_sheet_pdf（导出 PDF，CAP-02）、
  // set_sheet_view（画布页必须先隐藏网格线，CAP-07 验收暴露的缺口）、
  // add_shape/list_shapes/update_shape/group_shapes/ungroup_shapes/set_shape_zorder/export_shape_image
  // （矢量图形，CAP-07 WPS 表格侧 + CAP-08 Office.js 侧）——
  // 三者都是**真实存在的宿主方法**：dispatch.js 有 RPC 分支、excel.js 有实现、网关有处理器与注册项。
  // 同样是**真实存在的宿主方法**：wps-addon/src/dispatch.js 有 RPC 分支、excel.js 有
  // formatTextSegment 实现、网关有对应处理器与注册项。
  // 2026-09-22（CAP-14）：新增 `manage_workbook_views` —— WPS 侧「自定义视图」宿主方法。
  // 真机确认：`wb.CustomViews` 可用（实测能建视图并读回 Name），而 `wb.LinkedDataTypes` 不存在，
  // 所以「链接数据类型（富值）」在本机**做不到**（工具会如实拒绝），只实现自定义视图与切片器清单。
  // 该能力**仅 WPS 实现**，已同时声明进 `HOST_IMPLEMENTATION_GAPS.microsoft`。**有意扩展。**
  assert.equal(c.EXCEL_METHODS.length, 51, '宿主方法路由表不得在搬迁中增减（有意扩展需在此写明理由）');
  assert.equal(c.isReadOnlyTool('excel_read_range'), true);
  assert.equal(c.isReadOnlyTool('wps_inspect_api'), false, '表达式探测不是只读');
  assert.equal(c.isReplaySafeMethod('read_range'), true);
  assert.equal(c.isReplaySafeMethod('wps_inspect_api'), false, '表达式探测不可重放');
  assert.equal(c.isReplaySafeMethod('patch_cells'), false);
  assert.equal(c.bareName('wps_read_range'), 'read_range');
  assert.equal(c.bareName('excel_read_range'), 'read_range');
  assert.equal(c.bareName('read_range'), 'read_range');
  // 2026-09-22：microsoft 侧新增 get_style_token（WPS 独有能力的对侧缺口，声明后失败被归类为
  // "宿主未实现"而不是未知工具）。同样是有意扩展。
  // 2026-09-22：microsoft 侧再增 format_text_segment（Office.js 无字符级富文本入口）、
  // configure_print_layout 与 export_sheet_pdf（Office.js 交付面无对应 API）。
  assert.deepEqual(c.HOST_IMPLEMENTATION_GAPS, {
    // CAP-07 后 WPS 侧已实现形状增删改查/分组/层级；仍缺形状导图与 MS 独有的 get_active_shape
    wps: ['export_shape_image'],
    microsoft: ['get_style_token', 'format_text_segment', 'configure_print_layout', 'export_sheet_pdf', 'set_sheet_view',
      'manage_hyperlink', 'manage_named_range', 'manage_document_properties', 'manage_table', 'manage_pictures',
      // CAP-14：自定义视图（wb.CustomViews）在 WPS 可用、Office.js 侧未实现 → 声明为 MS 缺口
      'manage_workbook_views',
      'create_workbook']
  });
});
