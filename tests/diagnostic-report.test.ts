import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const writer = fileURLToPath(new URL('../scripts/lib/diagnostic-report.mjs', import.meta.url));
test('diagnostic JSON survives an unwritable destination in the console log', t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-report-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const blocker = path.join(root, 'ordinary-file'); fs.writeFileSync(blocker, 'keep');
  // Use dynamic import so the helper path remains an argument rather than code.
  const run = (destination: string) => spawnSync(process.execPath, ['--input-type=module', '-e', 'const {pathToFileURL}=await import("node:url"); const {writeDiagnosticReport}=await import(pathToFileURL(process.argv[1]).href); writeDiagnosticReport(process.argv[2], {scope:"fixture",helperResult:{phase:"failed",message:"synthetic failure"},authorizationGranted:false});', writer, destination], { encoding: 'utf8' });
  const failed = run(path.join(blocker, 'report.json'));
  assert.equal(failed.status, 0, failed.stderr);
  const body = failed.stdout.split('CHARFLOAT_DIAGNOSTIC_REPORT_BEGIN\n')[1].split('CHARFLOAT_DIAGNOSTIC_REPORT_END')[0];
  assert.equal(JSON.parse(body).helperResult.message, 'synthetic failure');
  assert.equal(JSON.parse(body).authorizationGranted, false);
  assert.equal(JSON.parse(failed.stderr).reportPreservedInConsole, true);
  assert.equal(fs.readFileSync(blocker, 'utf8'), 'keep');
  const file = path.join(root, 'report.json'), succeeded = run(file);
  assert.equal(succeeded.status, 0); assert.equal(succeeded.stderr, '');
  assert.deepEqual(JSON.parse(fs.readFileSync(file, 'utf8')), JSON.parse(body));
});
