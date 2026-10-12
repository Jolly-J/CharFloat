// 只读：从固定提交恢复原声明，与当前源码对照；不部署、不启动宿主。
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('../../../../', import.meta.url));
const dir = path.dirname(fileURLToPath(import.meta.url));
const ts = createRequire(path.join(root, 'package.json'))('typescript');
const baseline = JSON.parse(fs.readFileSync(path.join(dir, 'baseline.json')));
const ownership = JSON.parse(fs.readFileSync(path.join(dir, 'ownership.json')));
const hash = text => crypto.createHash('sha256').update(text).digest('hex');
const old = execFileSync('git', ['show', `${baseline.commit}:${ownership.original}`], { cwd: root, encoding: 'utf8', maxBuffer: 2 ** 22 });
assert.equal(hash(old), baseline.files[ownership.original]);
function declarations(file, text) {
  const source = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
  assert.equal(source.parseDiagnostics.length, 0, `语法错误: ${file}`);
  return [...source.statements].map(node => {
    const kind = ts.isFunctionDeclaration(node) ? 'function' : 'initializer';
    const name = ts.isFunctionDeclaration(node) ? node.name.text : ts.isVariableStatement(node) ? node.declarationList.declarations.map(d => d.name.getText(source)).join(',') : null;
    assert.ok(name, `意外顶层语句: ${file}`);
    return { name, kind, sha256: hash(node.getText(source)), file };
  });
}
const before = declarations(ownership.original, old);
const builder = fs.readFileSync(path.join(root, 'scripts/build-wps-addon.mjs'), 'utf8');
const buildSource = ts.createSourceFile('builder.mjs', builder, ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
const array = [...buildSource.statements].filter(ts.isVariableStatement).flatMap(n => [...n.declarationList.declarations]).find(n => n.name.getText(buildSource) === 'MODULES').initializer;
const modules = array.elements.map(n => n.text).filter(n => n.startsWith('excel/'));
assert.deepEqual([...modules].sort(), [...ownership.modules].sort());
const after = modules.flatMap(module => {
  const file = 'wps-addon/src/' + module;
  const text = fs.readFileSync(path.join(root, file), 'utf8').replace(/^[^\n]*@build-strip:start[\s\S]*?^[^\n]*@build-strip:end[^\n]*\n?/gm, '');
  return declarations(file, text);
});
assert.equal(new Set(after.map(n => n.name)).size, after.length, '顶层声明不得重名');
const normalized = list => list.map(({ file, ...n }) => n).sort((a, b) => a.name.localeCompare(b.name));
assert.deepEqual(normalized(after), normalized(before), '函数体与初始化表达式必须逐字保持一致');
assert.deepEqual(after.filter(n => n.kind === 'initializer').map(n => n.name), before.filter(n => n.kind === 'initializer').map(n => n.name), '初始化顺序必须保持一致');
for (const entry of ownership.declarations) assert.equal(after.find(n => n.name === entry.name)?.file, 'wps-addon/src/' + entry.module, `归属错误: ${entry.name}`);
for (const [file, expected] of Object.entries(baseline.files)) {
  if (file.startsWith('wps-addon/src/') && file !== ownership.original || file.startsWith('office-addon/')) assert.equal(hash(fs.readFileSync(path.join(root, file))), expected, `非目标源码变化: ${file}`);
}
console.log(JSON.stringify({ anchor: baseline.commit, functions: before.filter(n => n.kind === 'function').length, initializers: before.filter(n => n.kind === 'initializer').length, modules: modules.length, declarationsUnchanged: true, initializationOrderUnchanged: true, otherAddonSourcesUnchanged: true }, null, 2));
