#!/usr/bin/env node
/**
 * scripts/check-param-forwarding.ts
 *
 * 检查"**schema 声明了参数、网关处理器却没读取**"的静默丢参问题。
 *
 * 为什么需要这个检查：
 *   工具 schema 是给调用方看的契约，网关处理器是把参数送到宿主的唯一通道。
 *   两者脱节时，调用方**以为参数生效了**，实际它从未离开桥接进程——
 *   典型的"静默无效"，比报错更难发现。已经踩过两次：
 *     - ISS-106 `manage_slides` 的 `filePath`/`format` 未转发 → save_as 落宿主默认路径
 *     - ISS-89  `inspect_api` 的 `evaluate`/`maxMembers` 未转发 → 反射护栏按默认走
 *
 * 判定方式（启发式，允许少量误报）：
 *   1. 从 `src/bridge/tools/definitions/**` 收集每个工具 schema 的属性名；
 *   2. 从 `src/bridge/gateway*.ts` 与 `src/bridge/gateway/**` 里找该工具对应的处理器，
 *      收集处理器正文中出现的 `args?.X` / `args.X` / `args?.X ??` 等读取（以及
 *      `const { ... } = args` 解构）；
 *   3. schema 有、处理器正文完全没提到的属性 = 可疑丢参。
 *
 * 豁免：`host`（由统一入口在别处处理）与在 `PAIRS` 里显式列出的已知例外。
 *
 * 退出码：发现可疑丢参 → 1；否则 0。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const DEFINITIONS_DIR = path.join(ROOT, 'src/bridge/tools/definitions');
const GATEWAY_DIR = path.join(ROOT, 'src/bridge/gateway');
const GATEWAY_FACADE = path.join(ROOT, 'src/bridge/gateway.ts');

/** 由公共逻辑处理、不属于任何单个处理器的参数。 */
const GLOBAL_PARAMS = new Set(['host']);

/** 已知例外：schema 里有、但确实不由处理器转发（写明原因，不要为了过检查而删条目）。 */
const KNOWN_EXCEPTIONS: Record<string, string> = {
  // 键格式：`工具名.参数名`。值是**为什么故意不转发**（必须写清楚，否则以后没人敢删）。
  'wps_execute_script.readOnly':
    '该参数由**网关层**消费（gateway.ts 的 recordIdentityChangingOperation 读它来决定要不要把本次调用' +
    '登记为「未快照操作」），不需要转发给宿主加载项。宿主根本不关心这个标志。',
};

interface ToolDef { tool: string; params: string[] }

/**
 * 按大括号配平取出从 `from` 处 `{` 开始的对象字面量文本（含两端大括号）。
 * 用配平而不是正则，是因为参数 schema 里会有嵌套对象（如 `iconThresholds` 的元素结构），
 * 惰性正则会在内层 `}` 处提前收尾——**这曾让本检查漏报过真实丢参**，所以改成配平。
 */
function sliceBalanced(text: string, openIndex: number): string | null {
  if (text[openIndex] !== '{') return null;
  let depth = 0;
  for (let i = openIndex; i < text.length; i++) {
    const ch = text[i];
    if (ch === '{') depth++;
    else if (ch === '}') {
      depth--;
      if (depth === 0) return text.slice(openIndex, i + 1);
    }
  }
  return null;
}

export function collectToolDefs(): ToolDef[] {
  const out: ToolDef[] = [];
  for (const file of fs.readdirSync(DEFINITIONS_DIR).filter(f => f.endsWith('.ts'))) {
    const text = fs.readFileSync(path.join(DEFINITIONS_DIR, file), 'utf8');
    for (const m of text.matchAll(/name:\s*"([a-z0-9_]+)"/g)) {
      const tool = m[1];
      const after = text.indexOf('parameters:', m.index);
      if (after < 0) continue;
      const paramsOpen = text.indexOf('{', after);
      if (paramsOpen < 0) continue;
      const paramsBlock = sliceBalanced(text, paramsOpen);
      if (!paramsBlock) continue;
      const propsOpen = paramsBlock.indexOf('properties:');
      if (propsOpen < 0) continue;
      const propsBrace = paramsBlock.indexOf('{', propsOpen);
      const propsBlock = sliceBalanced(paramsBlock, propsBrace);
      if (!propsBlock) continue;
      // 只取 properties 的**第一层**键
      const names: string[] = [];
      let depth = 0;
      for (let i = 1; i < propsBlock.length - 1; i++) {
        const ch = propsBlock[i];
        if (ch === '{' || ch === '[') depth++;
        else if (ch === '}' || ch === ']') depth--;
        else if (depth === 0) {
          const rest = propsBlock.slice(i);
          const km = rest.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*:/);
          if (km && (i === 1 || /[\s,{]/.test(propsBlock[i - 1]))) {
            names.push(km[1]);
            i += km[1].length;
          }
        }
      }
      const params = [...new Set(names)];
      if (params.length) out.push({ tool, params });
    }
  }
  return out;
}

function collectHandlers(): { byModule: Map<string, string>; registry: Map<string, string> } {
  const byModule = new Map<string, string>();
  const registry = new Map<string, string>();
  const files = [GATEWAY_FACADE, ...fs.readdirSync(GATEWAY_DIR).filter(f => f.endsWith('.ts')).map(f => path.join(GATEWAY_DIR, f))];
  for (const file of files) {
    if (!fs.existsSync(file)) continue;
    const text = fs.readFileSync(file, 'utf8');
    const moduleName = path.basename(file).replace(/\.ts$/, '');
    // 处理器形如 export const xxx: Handler = async (ctx) => { ... }; 取到下一个 export const 为止。
    // **按 模块:函数名 建键**：拆分后不同类文件会重用同名处理器（如 ppt.ts 与 word.ts 都有
    // manageTable），只按函数名建键会被后读的文件覆盖、拿错处理器造成误报——这是第三个盲点。
    const re = /export const\s+(\w+)\s*:\s*Handler\s*=\s*async[\s\S]*?(?=\nexport const |\n\/\*\*|\Z)/g;
    for (const m of text.matchAll(re)) byModule.set(`${moduleName}:${m[1]}`, m[0]);
    // 门面里的注册表： "wps_xxx": <handler> 或 "wps_xxx": excel.<handler>
    // **必须允许点号**：按类拆分后注册值带模块前缀（如 `excel.addConditionalFormatting`），
    // 只认裸标识符会让这批工具整体漏检——这是第二个盲点。
    const regRe = /"([a-z0-9_]+)"\s*:\s*([A-Za-z_][A-Za-z0-9_.]*)\s*,/g;
    for (const m of text.matchAll(regRe)) registry.set(m[1], m[2]);
  }
  return { byModule, registry };
}

/** 解析注册值为处理器正文：`excel.foo` → gateway/excel.ts 的 foo；裸名则全局唯一查找。 */
function resolveHandler(ref: string, collected: { byModule: Map<string, string>; registry: Map<string, string> }): { body: string; label: string } | null {
  const bare = ref.includes('.') ? ref.split('.').pop()! : ref;
  const moduleName = ref.includes('.') ? ref.split('.')[0] : (ref === 'script' || ref === 'audit' ? ref : '');
  if (moduleName) {
    const hit = collected.byModule.get(`${moduleName}:${bare}`);
    if (hit) return { body: hit, label: `${moduleName}.${bare}` };
  }
  const matches = [...collected.byModule.entries()].filter(([k]) => k.endsWith(`:${bare}`));
  if (matches.length === 1) return { body: matches[0][1], label: matches[0][0] };
  if (matches.length > 1) return { body: matches.map(m => m[1]).join('\n'), label: `${bare}(合并${matches.length}处)` };
  return null;
}

/** 处理器正文里被读取的参数名。 */
function readParams(body: string): Set<string> {
  const found = new Set<string>();
  for (const m of body.matchAll(/args\s*\??\.\s*([A-Za-z_][A-Za-z0-9_]*)/g)) found.add(m[1]);
  for (const m of body.matchAll(/args\s*\??\.\s*\[\s*['"]([A-Za-z0-9_]+)['"]\s*\]/g)) found.add(m[1]);
  // const { a, b } = args || {}
  for (const m of body.matchAll(/const\s*\{([^}]*)\}\s*=\s*args\b/g)) {
    for (const part of m[1].split(',')) {
      const name = part.split(':')[0].trim().replace(/\.\.\./, '');
      if (name) found.add(name);
    }
  }
  return found;
}

/**
 * 核心判定：给定"schema 参数集合"与"处理器正文"，返回未被读取的参数。
 * 独立成函数是为了能跑自检（见文件末尾）——**一个永远通过的检查没有价值**。
 */
export function findUnreadParams(params: string[], handlerBody: string): string[] {
  const read = readParams(handlerBody);
  return params.filter(p => !GLOBAL_PARAMS.has(p) && !read.has(p));
}

/**
 * 自检：用一对人造输入证明"能抓到丢参"，并证明"读了就不报"。
 * 检查脚本自身失守（正则写错→永远 0 命中）比不检查更危险。
 */
function selfTest(): boolean {
  const def = { tool: '__self_test', params: ['alpha', 'beta'] };
  const handlerMissing = `export const t: Handler = async (ctx) => { const { args } = ctx; return args?.alpha; };`;
  const handlerFull = `export const t: Handler = async (ctx) => { const { args } = ctx; return [args?.alpha, args?.beta]; };`;
  const caught = findUnreadParams(def.params, handlerMissing);
  const clean = findUnreadParams(def.params, handlerFull);
  const ok = caught.length === 1 && caught[0] === 'beta' && clean.length === 0;
  console.log(
    ok
      ? '[check:params] 自检通过：能抓到未转发的参数，且读过的不误报'
      : `[check:params] ✖ 自检失败（漏报=${JSON.stringify(caught)} 误报=${JSON.stringify(clean)}）——检查本身不可信`
  );
  return ok;
}

function main() {
  if (!selfTest()) return 1;
  const defs = collectToolDefs();
  const collected = collectHandlers();
  const problems: { tool: string; param: string; handler: string }[] = [];
  let checked = 0;
  let unresolved = 0;

  for (const def of defs) {
    const ref = collected.registry.get(def.tool);
    if (!ref) { unresolved++; continue; } // 由统一入口派生（excel_*）或未注册，交给别的检查
    const resolved = resolveHandler(ref, collected);
    if (!resolved) { unresolved++; continue; }
    checked++;
    for (const param of findUnreadParams(def.params, resolved.body)) {
      // 例外必须显式登记并写明原因，否则一律算问题（KNOWN_EXCEPTIONS 原先只声明未使用，
      // 等于这个逃生口是假的——真去登记也不生效）。
      if (KNOWN_EXCEPTIONS[`${def.tool}.${param}`]) continue;
      problems.push({ tool: def.tool, param, handler: resolved.label });
    }
  }

  console.log(`[check:params] 扫描 ${defs.length} 个工具定义；已解析到处理器的 ${checked} 个，未解析 ${unresolved} 个（多为派生入口）`);
  if (problems.length === 0) {
    console.log('[check:params] 未发现"schema 有、处理器不读"的参数');
    return 0;
  }
  console.log('');
  console.log('可疑丢参（schema 声明了，但处理器正文里没有读取）：');
  for (const p of problems) console.log(`  ${p.tool}.${p.param}   (handler: ${p.handler})`);
  console.log('');
  console.log('处置：确实要转发就补上转发；确实不转发就从 schema 移除或写入 KNOWN_EXCEPTIONS 并写明原因。');
  return 1;
}

// 仅在直接执行时运行（被 import 做调试/单测时不自动跑）
const invokedDirectly = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (invokedDirectly) process.exit(main());
