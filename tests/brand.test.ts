import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
// @ts-ignore JS 检查器与实际构建使用同一入口。
import { checkBrand } from '../scripts/check-brand.mjs';

test('品牌检查覆盖源码、插件清单、路径及 UTF-16，精确排除历史范围', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-brand-'));
  const words = ['Office', 'Agent', 'Bridge'];
  const write = (name: string, value: string | Buffer) => {
    fs.mkdirSync(path.dirname(path.join(root, name)), { recursive: true });
    fs.writeFileSync(path.join(root, name), value);
  };
  try {
    write('src/current.ts', '字浮 CharFloat');
    write('docs/acceptance/prior.md', words.join(' '));
    write('release/prior.txt', words.join('-'));
    write('node_modules/third-party/index.js', words.join('_'));
    write('dist/main/index.js', words.join(''));
    assert.deepEqual(checkBrand(root).errors, []);

    write('wps-addon/ribbon.xml', words.join(''));
    write('office-addon/excel/manifest.xml', words.join(' '));
    write('src/legacy.ts', words.join('_').toUpperCase());
    write('website/src/command.ts', words.join('-').toLowerCase());
    write('docs/encoded.md', words.join('%20'));
    write('resources/metadata.txt', Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(words.join(' '), 'utf16le')]));
    write(`${words.join('-')}/fixture.txt`, '字浮');
    const errors: string[] = checkBrand(root).errors;
    for (const file of ['wps-addon/ribbon.xml', 'office-addon/excel/manifest.xml', 'src/legacy.ts', 'website/src/command.ts', 'docs/encoded.md', 'resources/metadata.txt']) {
      assert.ok(errors.some(error => error.startsWith(`${file}:`)), file);
    }
    assert.ok(errors.some(error => error.includes('目录包含旧品牌标识')));
    assert.ok(!errors.some(error => error.startsWith('docs/acceptance/')));
    assert.ok(!errors.some(error => error.startsWith('release/')));
    const built = checkBrand(root, 'generated').errors;
    assert.ok(built.some((error: string) => error.startsWith('dist/main/index.js:')));
    assert.ok(built.some((error: string) => error.includes('缺少构建产物')));
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});
