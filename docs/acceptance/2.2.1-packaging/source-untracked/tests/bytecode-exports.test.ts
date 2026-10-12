import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { execFileSync } from 'node:child_process';
import { buildSync } from 'esbuild';

test('Electron 字节码运行时核验 PDF 与图表落盘，缺失文件不伪装成功', () => {
  const root = path.resolve(import.meta.dirname, '..');
  const runtime = process.platform === 'darwin'
    ? path.join(root, 'node_modules/electron/dist/Electron.app/Contents/MacOS/Electron')
    : path.join(root, 'node_modules/electron/dist', process.platform === 'win32' ? 'electron.exe' : 'electron');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-bytecode-export-'));
  const source = path.join(tmp, 'exports.cjs');
  const compiled = path.join(tmp, 'exports.jsc');
  const bytenode = path.join(root, 'node_modules/bytenode');
  const pdf = path.join(tmp, 'example.pdf'), png = path.join(tmp, 'example.png'), missing = path.join(tmp, 'missing.png');
  fs.writeFileSync(pdf, '%PDF-test'); fs.writeFileSync(png, 'test-image');
  try {
    buildSync({
      stdin: { contents: `
        import assert from 'node:assert/strict';
        import { exportSheetPdf, exportChartImage } from ${JSON.stringify(path.join(root, 'src/bridge/gateway/excel.ts'))};
        let calls = 0;
        const ctx = { args: {}, previewPath: () => ${JSON.stringify(pdf)}, callOffice: async (method, args) => {
          calls++; return { success: true, outputPath: args.outputPath };
        }};
        (async () => {
          const pdf = await exportSheetPdf(ctx);
          assert.equal(pdf.verifiedOnDisk, true); assert.equal(pdf.fileSizeBytes, 9);
          const image = await exportChartImage({ ...ctx, args: { outputPath: ${JSON.stringify(png)} } });
          assert.equal(image.fileWritten, true); assert.equal(image.fileSizeBytes, 10);
          const missingImage = await exportChartImage({ ...ctx, args: { outputPath: ${JSON.stringify(missing)} } });
          assert.equal(missingImage.fileWritten, false); assert.ok(missingImage.warnings.length);
          await assert.rejects(exportSheetPdf({ ...ctx, previewPath: () => ${JSON.stringify(missing)} }), /未落盘/);
          assert.equal(calls, 4, '落盘核验不得重新调用宿主导出');
          console.log(JSON.stringify({ pdf: true, image: true, missingFilesRejected: true, calls }));
        })().catch(error => { console.error(error.message); process.exitCode = 1; });
      `, resolveDir: root, loader: 'ts' },
      outfile: source, bundle: true, platform: 'node', format: 'cjs', target: 'node20', minify: true
    });
    const env = { ...process.env, ELECTRON_RUN_AS_NODE: '1' };
    execFileSync(runtime, ['-e', `require(${JSON.stringify(bytenode)}).compileFile({ filename: ${JSON.stringify(source)}, output: ${JSON.stringify(compiled)} });`], { env, timeout: 30000 });
    const output = execFileSync(runtime, ['-e', `require(${JSON.stringify(bytenode)}); require(${JSON.stringify(compiled)});`], { env, encoding: 'utf8', timeout: 30000 });
    assert.deepEqual(JSON.parse(output.trim()), { pdf: true, image: true, missingFilesRejected: true, calls: 4 });
  } finally { fs.rmSync(tmp, { recursive: true, force: true }); }
});
