import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
// @ts-ignore 构建器共用的只读清单校验。
import { assertWpsModuleManifest, listWpsModules } from '../scripts/lib/wps-addon-modules.mjs';

function fixture(run: (root: string) => void) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'wps-manifest-'));
  try {
    fs.mkdirSync(path.join(root, 'excel/deep'), { recursive: true });
    for (const file of ['shared.js', 'excel/range.js', 'excel/deep/pure.js', 'excel/notes.md']) {
      fs.writeFileSync(path.join(root, file), '// fixture\n');
    }
    run(root);
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
}

const modules = ['shared.js', 'excel/range.js', 'excel/deep/pure.js'];
test('合法多层源码清单通过且校验不改变文件', () => fixture(root => {
  const before = modules.map(file => fs.readFileSync(path.join(root, file)));
  assert.deepEqual(listWpsModules(root), [...modules].sort());
  assertWpsModuleManifest(root, modules);
  modules.forEach((file, index) => assert.deepEqual(fs.readFileSync(path.join(root, file)), before[index]));
}));
test('子目录新增未登记片段必须拒绝', () => fixture(root => {
  fs.writeFileSync(path.join(root, 'excel/deep/forgotten.js'), 'function missed() {}');
  assert.throws(() => assertWpsModuleManifest(root, modules), /未登记模块.*excel\/deep\/forgotten.js/);
}));
test('缺失子目录片段必须拒绝', () => fixture(root => {
  fs.unlinkSync(path.join(root, 'excel/range.js'));
  assert.throws(() => assertWpsModuleManifest(root, modules), /文件不存在.*excel\/range.js/);
}));
test('重复登记的片段必须拒绝', () => fixture(root => {
  assert.throws(() => assertWpsModuleManifest(root, [...modules, 'excel/range.js']), /重复登记.*excel\/range.js/);
}));
