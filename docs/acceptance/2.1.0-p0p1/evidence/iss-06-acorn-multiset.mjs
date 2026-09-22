// ISS-06-A：acorn 顶层语句多重集比对
// 用法: node .scratch/evidence/acorn-multiset.mjs <改造前文件> <改造后文件>
// 说明: 两个加载项入口都是 "外层 IIFE + 顶层语句序列" 结构，
//       因此把 IIFE 函数体的 body 视为「顶层语句」，用 acorn 的 node.start/node.end
//       切出源码字节片段做多重集比对（逐字节，零归一化）。
import fs from 'node:fs';
import crypto from 'node:crypto';
import * as acorn from 'acorn';

function topLevelStatements(file) {
  const src = fs.readFileSync(file, 'utf8');
  const ast = acorn.parse(src, { ecmaVersion: 'latest', sourceType: 'script', locations: true });
  // Program.body[0] = ExpressionStatement( CallExpression( FunctionExpression ) )
  let stmt = ast.body[0];
  let inner = null;
  if (stmt.type === 'ExpressionStatement' && stmt.expression.type === 'CallExpression') {
    const callee = stmt.expression.callee;
    if (callee.type === 'FunctionExpression' || callee.type === 'ArrowFunctionExpression') inner = callee.body;
  }
  if (!inner || inner.type !== 'BlockStatement') {
    throw new Error(`${file}: 顶层不是 IIFE，无法定位顶层语句`);
  }
  return {
    src,
    programStatements: ast.body.length,
    statements: inner.body.map((node) => ({
      type: node.type,
      kind: node.kind || '',
      name: node.id && node.id.name ? node.id.name : '',
      line: node.loc.start.line,
      endLine: node.loc.end.line,
      code: src.slice(node.start, node.end),
      sha: crypto.createHash('sha256').update(src.slice(node.start, node.end)).digest('hex').slice(0, 12)
    }))
  };
}

const [beforeFile, afterFile] = process.argv.slice(2);
const A = topLevelStatements(beforeFile);
const B = topLevelStatements(afterFile);

const line = (s) => s.replace(/\s+/g, ' ').trim().slice(0, 96);
const label = (s) => `${s.type}${s.kind ? ':' + s.kind : ''}${s.name ? ' ' + s.name : ''}`;

// 多重集比对：按 sha 计数
function multiset(list) {
  const m = new Map();
  for (const s of list) m.set(s.sha, (m.get(s.sha) || 0) + 1);
  return m;
}
const ma = multiset(A.statements);
const mb = multiset(B.statements);

const onlyBefore = [];
const onlyAfter = [];
for (const s of A.statements) if ((mb.get(s.sha) || 0) === 0) onlyBefore.push(s);
for (const s of B.statements) if ((ma.get(s.sha) || 0) === 0) onlyAfter.push(s);

const common = new Set();
for (const s of A.statements) if ((mb.get(s.sha) || 0) > 0) common.add(s.sha);

console.log('=== ISS-06-A · acorn 顶层语句多重集比对 ===');
console.log(`改造前: ${beforeFile}  (${A.src.split('\n').length} 行, Program.body=${A.programStatements})`);
console.log(`改造后: ${afterFile}  (${B.src.split('\n').length} 行, Program.body=${B.programStatements})`);
console.log(`顶层语句数: 改造前 ${A.statements.length} / 改造后 ${B.statements.length}`);
console.log(`逐字节相同的语句（多重集去重后）: ${new Set([...ma.keys()].filter((k) => mb.has(k))).size}`);
console.log('');
console.log(`改造前有、改造后无（按 sha 计）: ${onlyBefore.length}`);
for (const s of onlyBefore) console.log(`  - L${s.line}-${s.endLine} [${label(s)}] ${line(s.code)}`);
console.log(`改造后有、改造前无（按 sha 计）: ${onlyAfter.length}`);
for (const s of onlyAfter) console.log(`  + L${s.line}-${s.endLine} [${label(s)}] ${line(s.code)}`);
console.log('');
console.log('--- 关键计数：改造前语句中能在改造后逐字节找到的数量 ---');
let matched = 0;
const mbCopy = new Map(mb);
for (const s of A.statements) {
  const n = mbCopy.get(s.sha) || 0;
  if (n > 0) { matched++; mbCopy.set(s.sha, n - 1); }
}
console.log(`改造前语句总数 ${A.statements.length}，其中逐字节可在改造后找到 ${matched} → ${matched}/${A.statements.length}`);
console.log(`改造后语句总数 ${B.statements.length}，多出 ${B.statements.length - A.statements.length} 条`);
console.log('');
console.log('--- 每条「改造前独有」语句的最相似改造后语句（按名字/类型配对） ---');
for (const s of onlyBefore) {
  const cand = B.statements.find((t) => t.type === s.type && t.name === s.name && s.name);
  console.log(`  · ${label(s)} @ L${s.line}-${s.endLine}`);
  if (cand) console.log(`      配对改造后: @ L${cand.line}-${cand.endLine} sha ${cand.sha}`);
  else console.log('      未按名字找到配对');
}
console.log('');
console.log('--- 每条「改造后独有」语句的同名改造前语句（若有） ---');
for (const s of onlyAfter) {
  const cand = A.statements.find((t) => t.type === s.type && t.name === s.name && s.name);
  console.log(`  · ${label(s)} @ L${s.line}-${s.endLine}`);
  if (cand) console.log(`      配对改造前: @ L${cand.line}-${cand.endLine} sha ${cand.sha}`);
  else console.log('      改造前无同名语句');
}
