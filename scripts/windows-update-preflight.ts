// Offline security checks only. Runs with the installed Electron Node runtime,
// writes temporary fixtures and a report; never updates an installed application.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { fork, execFileSync } from 'node:child_process';
import { createHash, generateKeyPairSync, sign } from 'node:crypto';
import { canonicalDirectory, createInstallTicket, grantInstallAuthorization, consumeInstallAuthorization, revokeInstallAuthorization, writeInstallState } from '../src/main/update-install.js';
import { prepareWindowsUpdate } from '../windows-bootstrap/update.mjs';
import { PRODUCT } from '../windows-bootstrap/core.mjs';
import { writeDiagnosticReport } from './lib/diagnostic-report.mjs';

if (process.argv[2] === '--claim') {
  process.once('message', async (value: { file: string; nonce: string }) => {
    try { const guard = await consumeInstallAuthorization(value.file, value.nonce); await guard(); process.send?.({ claimed: true }, () => process.disconnect()); }
    catch { process.send?.({ claimed: false }, () => process.disconnect()); }
  });
  process.send?.({ ready: true });
} else {
  void main().catch((error: any) => { process.exitCode = 1; console.error(JSON.stringify({ diagnosticFailed: true, code: error.code, name: error.name, installs: false })); });
}

async function main() {
  const output = path.resolve(process.argv[2] || path.dirname(process.argv[1]));
  const report: any = { schema: 1, scope: 'offline update security checks; not GUI or full Windows CI acceptance', platform: process.platform, arch: process.arch, node: process.versions.node, electron: process.versions.electron, checks: [], passed: false };
  const file = path.join(output, `preflight-${Date.now()}.json`);
  if (process.platform !== 'win32' || process.arch !== 'x64') {
    report.checks.push({ name: 'Windows x64 runtime', passed: false });
    writeDiagnosticReport(file, report); process.exitCode = 2; return;
  }
  const config = JSON.parse(fs.readFileSync(path.join(path.dirname(process.argv[1]), 'preflight-config.json'), 'utf8'));
  report.runtimeMatches = createHash('sha256').update(fs.readFileSync(process.execPath)).digest('hex') === config.runtimeSha256;
  if (!report.runtimeMatches) { report.checks.push({ name: 'expected test runtime', passed: false }); writeDiagnosticReport(file, report); process.exitCode = 2; return; }
  report.sourceInputs = config.sourceInputs;
  const rawRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-win-preflight-'));
  const root = fs.realpathSync(rawRoot);
  const job = async () => createInstallTicket(root, { parentPid: process.ppid, platform: 'win32', arch: 'x64', currentVersion: '2.2.906', target: path.join(root, 'old'), archive: path.join(root, 'package.zip'), envelope: '{}' });
  async function check(name: string, action: () => Promise<unknown>) {
    try { await action(); report.checks.push({ name, passed: true }); }
    catch (error: any) { report.checks.push({ name, passed: false, error: error.code || error.name || 'Error' }); }
  }
  try {
    await check('temporary path aliases and linked-parent rejection', async () => {
      await canonicalDirectory(rawRoot); await canonicalDirectory(root);
      const short = execFileSync('cmd.exe', ['/d', '/u', '/c', 'for %I in ("%CHARFLOAT_PROBE_PATH%") do @echo %~sI'], { encoding: 'utf16le', windowsHide: true, env: { ...process.env, CHARFLOAT_PROBE_PATH: rawRoot } }).trim();
      assert.ok(path.isAbsolute(short)); await canonicalDirectory(short);
      report.shortPathDiffers = short !== rawRoot;
      const target = path.join(root, 'target'), link = path.join(root, 'junction');
      fs.mkdirSync(path.join(target, 'cache'), { recursive: true }); fs.symlinkSync(target, link, 'junction');
      await assert.rejects(canonicalDirectory(link)); await assert.rejects(canonicalDirectory(path.join(link, 'cache')));
      report.temporaryPathAliasExpandedByNativeRealpath = root !== await fs.promises.realpath(rawRoot);
    });
    await check('twelve separate processes consume one grant once', async () => {
      const ticket = await job(); await grantInstallAuthorization(ticket.file, ticket.ticket.nonce);
      const children: ReturnType<typeof fork>[] = [];
      let timer: ReturnType<typeof setTimeout> | undefined;
      try {
        const pending = Array.from({ length: 12 }, () => {
          const child = fork(process.argv[1], ['--claim'], { execArgv: [], windowsHide: true, stdio: ['ignore', 'ignore', 'ignore', 'ipc'] }); children.push(child);
          let ready = false, answer: boolean | undefined;
          const result = new Promise<boolean>((resolve, reject) => {
            child.once('error', reject);
            child.on('message', (value: any) => { if (typeof value.claimed === 'boolean') answer = value.claimed; });
            child.once('exit', code => code === 0 && answer !== undefined ? resolve(answer) : reject(new Error('worker result missing')));
          });
          // Attach a handler immediately so an early worker failure is never unhandled.
          void result.catch(() => {});
          return { result, ready: new Promise<void>((resolve, reject) => { child.once('error', reject); child.once('exit', () => { if (!ready) reject(new Error('worker stopped')); }); child.on('message', (value: any) => { if (value.ready) { ready = true; resolve(); } }); }) };
        });
        const run = async () => {
          await Promise.all(pending.map(value => value.ready));
          for (const child of children) child.send({ file: ticket.file, nonce: ticket.ticket.nonce });
          const claims = await Promise.all(pending.map(value => value.result)); assert.equal(claims.filter(Boolean).length, 1);
          await assert.rejects(consumeInstallAuthorization(ticket.file, ticket.ticket.nonce));
          await assert.rejects(grantInstallAuthorization(ticket.file, ticket.ticket.nonce));
        };
        await Promise.race([run(), new Promise((_, reject) => { timer = setTimeout(() => reject(new Error('worker timeout')), 30000); })]);
      } finally { clearTimeout(timer); for (const child of children) if (child.exitCode === null) child.kill(); }
    });
    await check('revoked handoff refuses late ready, grant and consume; fresh retry works', async () => {
      const ticket = await job(); await grantInstallAuthorization(ticket.file, ticket.ticket.nonce); await revokeInstallAuthorization(ticket.file, ticket.ticket.nonce);
      await assert.rejects(writeInstallState(ticket.directory, { phase: 'ready', nonce: ticket.ticket.nonce }));
      await assert.rejects(grantInstallAuthorization(ticket.file, ticket.ticket.nonce)); await assert.rejects(consumeInstallAuthorization(ticket.file, ticket.ticket.nonce));
      const retry = await job(); await grantInstallAuthorization(retry.file, retry.ticket.nonce);
      const guard = await consumeInstallAuthorization(retry.file, retry.ticket.nonce); await guard(); await revokeInstallAuthorization(retry.file, retry.ticket.nonce); await assert.rejects(guard());
    });
    await check('signed Windows handoff validates target and revocation before commit', async () => {
      const local = path.join(root, 'local'), installRoot = path.join(local, 'Programs', 'CharFloat');
      const oldHash = 'a'.repeat(64), target = path.join(installRoot, `2.2.906-${oldHash.slice(0, 12)}`);
      fs.mkdirSync(target, { recursive: true }); fs.writeFileSync(path.join(installRoot, '.charfloat-install.json'), JSON.stringify({ product: PRODUCT }));
      fs.writeFileSync(path.join(target, '.charfloat-install.json'), JSON.stringify({ product: PRODUCT, version: '2.2.906', payloadSha256: oldHash }));
      fs.writeFileSync(path.join(target, '字浮 CharFloat.exe'), 'synthetic target; never launched');
      const archive = path.join(root, 'package.zip'), bytes = Buffer.from('synthetic signed bytes; never executed'); fs.writeFileSync(archive, bytes);
      const sha256 = createHash('sha256').update(bytes).digest('hex'), now = Date.now();
      // Ephemeral fixture keys only; this tool never reads any production key.
      const keys = generateKeyPairSync('ed25519');
      const trust = { feeds: ['https://example.test/stable.json'], assetOrigins: ['https://example.test'], publicKeys: { fixture: keys.publicKey.export({ format: 'pem', type: 'spki' }).toString() } };
      const payload = Buffer.from(JSON.stringify({ schema: 1, product: 'charfloat', channel: 'stable', releaseId: 1, version: '2.2.907', publishedAt: new Date(now).toISOString(), expiresAt: new Date(now + 86400000).toISOString(), notes: '', assets: [{ platform: 'win32', arch: 'x64', kind: 'windows-zip', size: bytes.length, sha256, urls: ['https://example.test/package.zip'] }] }));
      const envelope = JSON.stringify({ keyId: 'fixture', payload: payload.toString('base64'), signature: sign(null, payload, keys.privateKey).toString('base64') });
      const ticket = await createInstallTicket(root, { parentPid: process.ppid, platform: 'win32', arch: 'x64', currentVersion: '2.2.906', target, archive, envelope });
      const prepared = await prepareWindowsUpdate(ticket.file, { version: '2.2.907', payloadSha256: sha256 }, local, { trust, platform: 'win32', arch: 'x64', wait: async () => {} });
      await assert.rejects(prepared.update.claimAuthorization()); await grantInstallAuthorization(ticket.file, ticket.ticket.nonce);
      const guard = await prepared.update.claimAuthorization(); await guard(); await revokeInstallAuthorization(ticket.file, ticket.ticket.nonce); await assert.rejects(guard());
    });
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
  report.passed = report.checks.length === 4 && report.checks.every((value: any) => value.passed);
  writeDiagnosticReport(file, report);
  console.log(report.passed ? 'PASS: four offline security checks passed.' : 'FAIL: stop before attempting an update.');
  process.exitCode = report.passed ? 0 : 1;
}
