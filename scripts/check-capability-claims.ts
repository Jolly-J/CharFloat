/**
 * 能力声明一致性检查（改造计划 P2.6）。
 *
 * 用途：skill 与官网里的人工说明必须与**运行时元数据源**一致，不能各写一套数字。
 * 元数据源就是 `catalog.getTools()` 与 `catalog.capabilities()`——MCP 的 `bridge://capabilities`
 * 资源、`bridge_get_capabilities` 工具、HTTP `/api/v1/capabilities` 三处都读它。
 *
 * 本脚本只检查**可机械判定的部分**：文档中出现的"能力数量"表述是否与元数据源相符。
 * 其余定性说明（哪个宿主支持什么、验证到什么程度）仍需人工核对，脚本不能代替。
 *
 * 副作用：只读。不监听端口、不连接宿主。
 *
 * 用法：npm run check:claims
 * 退出码：0 = 一致；1 = 发现与元数据源矛盾的表述。
 */
import fs from 'node:fs';
import path from 'node:path';
import { getTools, capabilities, EXCEL_METHODS } from '../src/bridge/catalog.js';

const root = path.resolve(import.meta.dirname, '..');

/**
 * 扫描范围：人工维护的说明文本。生成物与依赖不参与。
 *
 * 为什么包含 `office-addon/public` 与 `wps-addon`：任务窗格与加载项页面里有**面向用户**的能力文案，
 * 2026-09-22 总验收审计发现 `office-addon/public/taskpane.html` 残留过 "42 项能力已就绪"（无任何口径对应），
 * 而该文件是手写文件（`build:office-addon` 只生成 `taskpane.js`），此前不在扫描范围内。扩展名限定为
 * `.md/.tsx/.ts/.html`，因此生成物 `taskpane.js`/`addon-core.js` 仍不会被扫到。
 */
const TARGET_DIRS = ['skills', path.join('website', 'src'), path.join('office-addon', 'public'), 'wps-addon'];
const EXTENSIONS = new Set(['.md', '.tsx', '.ts', '.html']);

/**
 * 能力口径（互不相等，混用会得出误导结论）：
 * - 路由数：`EXCEL_METHODS` 长度，是宿主方法路由表，**不代表任何宿主真的支持**；
 * - 工具数：对 AI 实际可调用的对外工具总数；
 * - 宿主可调用数：某个宿主上"声明已实现且有可调用工具"的方法数。
 */
interface Denominators {
  routing: number;
  tools: number;
  wpsCallable: number;
  microsoftCallable: number;
}

type ClaimKind =
  /** 口径明确：必须等于指定的量 */
  | { kind: 'exact'; pick: 'tools' }
  /** 口径含糊：必须能落在某个宿主可调用数上，否则说明它只是路由数或凭空写的 */
  | { kind: 'host-capability' };

interface ClaimPattern { re: RegExp; claim: ClaimKind; label: string }

const CLAIM_PATTERNS: ClaimPattern[] = [
  { re: /(\d+)\s*个对外工具/g, claim: { kind: 'exact', pick: 'tools' }, label: '对外工具总数' },
  { re: /工具总数\s*[:：]?\s*(\d+)/g, claim: { kind: 'exact', pick: 'tools' }, label: '对外工具总数' },
  { re: /(\d+)\s*项结构化能力/g, claim: { kind: 'host-capability' }, label: '结构化能力数' },
  { re: /(\d+)\s*项[^，。；、\n]{0,16}?能力/g, claim: { kind: 'host-capability' }, label: '能力数' },
  { re: /全量\s*(\d+)\s*项/g, claim: { kind: 'host-capability' }, label: '全量能力数' }
];

function walk(dir: string, out: string[] = []): string[] {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.name.startsWith('.') || entry.name === 'node_modules' || entry.name === 'dist') continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(full, out);
    else if (EXTENSIONS.has(path.extname(entry.name))) out.push(full);
  }
  return out;
}

const tools = getTools();
const caps: any = capabilities();
const denominators: Denominators = {
  routing: EXCEL_METHODS.length,
  tools: tools.length,
  wpsCallable: caps.hosts.wps.excel.implemented.length,
  microsoftCallable: caps.hosts.microsoft.excel.implemented.length
};

/** 行内宿主识别：不能仅凭数字反推宿主，必须由文本明确指出。 */
const HOST_HINTS: { host: 'wps' | 'microsoft'; re: RegExp }[] = [
  { host: 'microsoft', re: /Microsoft|微软|MS Office/i },
  { host: 'wps', re: /WPS|金山/i }
];

function hostOf(line: string): 'wps' | 'microsoft' | null {
  const hits = HOST_HINTS.filter(h => h.re.test(line));
  // 同一行同时提到两个宿主时无法归属，按"未指明"处理，交人工。
  return hits.length === 1 ? hits[0].host : null;
}

/**
 * 已证实错误的定性表述（回归守卫，不是风格检查）。
 *
 * 每一条都对应一次真实缺陷：文档把"未验收"写成"已统一验收"，或把当前
 * Office.js 优先 / Windows COM 回退的通道说成只有 COM。数量检查发现不了这类问题，
 * 因为它们不含数字。
 *
 * 新增条目时必须写明"为什么"——只写"这是禁词"会把有依据的表述也禁掉。
 */
const BANNED_CLAIMS: { pattern: RegExp; why: string }[] = [
  {
    pattern: /macOS\s*\/\s*Windows\s*统一/,
    why: '把未实机验收说成两平台已统一验收；本候选版本没有任何实机证据',
  },
  {
    pattern: /面向\s*Windows\s*COM/,
    why: 'Microsoft Excel 结构化工具走 Office.js 任务窗格通道（macOS 与 Windows 同一路径），COM 只是 Windows 回退',
  },
  {
    pattern: /macOS\s*Microsoft\s*Excel[^\n]*未实现结构化/,
    why: 'macOS 与 Windows 走同一 Office.js 通道，不是"macOS 未实现"',
  },
];

const errors: string[] = [];
const notes: string[] = [];
let scanned = 0;

for (const target of TARGET_DIRS) {
  const dir = path.join(root, target);
  if (!fs.existsSync(dir)) continue;
  for (const file of walk(dir)) {
    scanned++;
    const rel = path.relative(root, file);
    const lines = fs.readFileSync(file, 'utf8').split('\n');
    lines.forEach((line, index) => {
      for (const { pattern, why } of BANNED_CLAIMS) {
        if (pattern.test(line)) {
          errors.push(`${rel}:${index + 1} 出现已证实错误的表述「${line.trim().slice(0, 60)}」——${why}`);
        }
      }
      for (const { re, claim, label } of CLAIM_PATTERNS) {
        for (const match of line.matchAll(re)) {
          const value = Number(match[1]);
          const at = `${rel}:${index + 1}`;
          if (claim.kind === 'exact') {
            if (value !== denominators[claim.pick]) {
              errors.push(`${at} 出现「${match[0]}」，但${label}实际为 ${denominators[claim.pick]}`);
            }
            continue;
          }
          // 口径含糊的表述：必须由文本明确指出宿主，再按该宿主的口径核对。
          const host = hostOf(line);
          if (!host) {
            errors.push(
              `${at} 出现「${match[0]}」但未指明宿主，无法判定；请写明宿主（如「WPS ${denominators.wpsCallable} 项」）` +
              `，或改为以 bridge_get_capabilities 为准。不得仅凭数字反推宿主`
            );
            continue;
          }
          const expected = host === 'wps' ? denominators.wpsCallable : denominators.microsoftCallable;
          if (value !== expected) {
            errors.push(`${at} 出现「${match[0]}」，按 ${host === 'wps' ? 'WPS' : 'Microsoft'} 口径应为 ${expected} 项`);
          } else {
            notes.push(`${at} 「${match[0]}」按 ${host === 'wps' ? 'WPS' : 'Microsoft'} 口径核对通过`);
          }
        }
      }
    });
  }
}

if (errors.length) {
  console.error('能力声明与元数据源不一致：');
  for (const line of errors) console.error(`  ✖ ${line}`);
  console.error('\n请改写文档，或先确认元数据源是否应变更（不要只改数字）。');
  process.exitCode = 1;
} else {
  console.log(`能力声明检查通过：扫描 ${scanned} 个文件，未发现与元数据源矛盾的数量表述。`);
  console.log(`  口径：路由表 ${denominators.routing} 项 | 对外工具 ${denominators.tools} 个 | WPS 可调用 ${denominators.wpsCallable} 项 | Microsoft 可调用 ${denominators.microsoftCallable} 项。`);
  for (const note of notes) console.log(`  提示：${note}`);
  console.log('  说明：定性表述本检查只覆盖「已证实错误表述」黑名单（每轮发现的错误表述会加入），其余仍需人工核对。');
}
