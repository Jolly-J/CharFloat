// Diagnose an installed acceptance helper without granting installation permission.
// Only the known test application's cache and a new diagnostic ticket are accessed.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { createInstallTicket, readInstallTicket, revokeInstallAuthorization, plainFile, canonicalDirectory } from '../src/main/update-install.js';
import { UPDATE_TRUST } from '../src/main/update-trust.js';
import { writeDiagnosticReport } from './lib/diagnostic-report.mjs';
import { checkedProbeArchive } from './lib/update-probe-paths.js';

const clean = (value: unknown) => String(value || '').replace(/[A-Za-z]:[\\/]Users[\\/][^\\/\r\n]+/gi, '<user>').split(os.homedir()).join('<home>').slice(-10000);

async function verifyInChild(file: string) {
  return await new Promise<any>((resolve, reject) => {
    const verifier = spawn(process.execPath, [process.argv[1], '--verify-ticket', file], { env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '', stderr = '', settled = false;
    const finish = (error?: Error, result?: object) => { if (settled) return; settled = true; clearTimeout(timer); if (error) reject(error); else resolve(result); };
    const timer = setTimeout(() => { verifier.kill('SIGTERM'); finish(new Error('诊断票据验证子进程超时，未启动更新助手。')); }, 30000);
    verifier.stdout!.on('data', bytes => { stdout = (stdout + bytes).slice(-16384); });
    verifier.stderr!.on('data', bytes => { stderr = (stderr + bytes).slice(-4096); });
    verifier.once('error', error => finish(error));
    verifier.once('close', code => {
      try {
        const result = JSON.parse(stdout);
        if (code !== 0 || !result.ok) { const error: any = new Error(result.error?.message || '诊断票据验证失败。'); error.code = result.error?.code; finish(error); }
        else finish(undefined, result.packageVerified);
      } catch { finish(new Error('诊断票据验证子进程未返回有效 JSON：' + clean(stderr))); }
    });
  });
}

async function verifyTicketMode() {
  try {
    if (!process.argv[3] || process.argv[4]) throw new Error('诊断验证参数无效。');
    // Run in a different PID from the ticket creator, just like the real helper.
    // The normal parentPid guard remains intact.
    const value = await readInstallTicket(process.argv[3], UPDATE_TRUST, 'win32', 'x64');
    console.log(JSON.stringify({ ok: true, packageVerified: { version: value.selected.manifest.version, size: value.selected.asset.size, sha256: value.selected.asset.sha256 } }));
  } catch (error: any) { console.log(JSON.stringify({ ok: false, error: { code: error.code, message: clean(error.message) } })); process.exitCode = 1; }
}

async function main() {
  const output = path.resolve(process.argv[2] || path.dirname(process.argv[1]));
  const reportFile = path.join(output, `handoff-report-${Date.now()}.json`);
  const report: any = { schema: 1, diagnosticRevision: 4, scope: 'Windows acceptance helper preparation only; never grants or installs', platform: process.platform, arch: process.arch, electron: process.versions.electron, node: process.versions.node, authorizationGranted: false, checks: [] };
  let job: Awaited<ReturnType<typeof createInstallTicket>> | undefined;
  let child: ReturnType<typeof spawn> | undefined;
  let exit: { code: number | null; signal: string | null } | undefined;
  let stderr = '', stdout = '';
  try {
    report.stage = 'runtime-identity';
    if (process.platform !== 'win32' || process.arch !== 'x64') throw new Error('此工具仅用于 Windows x64。');
    const config = JSON.parse(fs.readFileSync(path.join(path.dirname(process.argv[1]), 'handoff-config.json'), 'utf8'));
    const runtimeHash = createHash('sha256').update(fs.readFileSync(process.execPath)).digest('hex');
    if (runtimeHash !== config.runtimeSha256) throw new Error('运行时与本轮 A 不一致，未启动助手。');
    report.runtimeMatches = true;
    report.stage = 'cache-locate';
    const cache = path.join(fs.realpathSync(os.tmpdir()), config.runtimeDirectoryName, 'desktop-state', 'updates');
    await canonicalDirectory(cache);
    const entries: { file: string; time: number }[] = [];
    for (const name of fs.readdirSync(cache)) {
      if (!/^install-[A-Za-z0-9]+$/.test(name)) continue;
      const file = path.join(cache, name, 'ticket.json');
      try { const stat = await plainFile(file, 512 * 1024); entries.push({ file, time: stat.mtimeMs }); } catch {}
    }
    entries.sort((a, b) => b.time - a.time);
    if (!entries.length) throw new Error('没有找到本轮更新交接记录；请保留 A 及下载好的 B。');
    const old = JSON.parse(fs.readFileSync(entries[0].file, 'utf8'));
    report.stage = 'target-identity';
    const target = path.dirname(process.execPath);
    await canonicalDirectory(target); await canonicalDirectory(old.target);
    if (old.currentVersion !== config.version || path.relative(await fs.promises.realpath(old.target), await fs.promises.realpath(target)) !== '') throw new Error('旧交接记录不属于当前 A，未重放。');
    report.stage = 'archive-containment';
    const paths = await checkedProbeArchive(cache, old.archive);
    report.pathForms = { cache: clean(cache), cachePhysical: clean(paths.cachePhysical), archive: clean(old.archive), archivePhysical: clean(paths.archivePhysical), relativeAfterNormalization: clean(paths.relative), cacheSpellingChanged: paths.cacheSpellingChanged, archiveSpellingChanged: paths.archiveSpellingChanged, replayPreservesArchiveSpelling: true };
    try {
      const previousFile = path.join(path.dirname(entries[0].file), 'result.json'); await plainFile(previousFile, 16384);
      const previous = JSON.parse(fs.readFileSync(previousFile, 'utf8'));
      report.previousResult = { phase: previous.phase, message: clean(previous.message) };
    } catch {}
    const preflight = fs.readdirSync(output).filter(name => /^preflight-\d+\.json$/.test(name)).sort().at(-1);
    if (preflight) {
      const file = path.join(output, preflight); await plainFile(file, 128 * 1024); const value = JSON.parse(fs.readFileSync(file, 'utf8'));
      report.preflight = { passed: value.passed === true, runtimeMatches: value.runtimeMatches === true, checks: Array.isArray(value.checks) ? value.checks.slice(0, 8).map((check: any) => ({ name: clean(check.name), passed: check.passed === true, error: clean(check.error) })) : [] };
    } else report.preflight = { executionReportFound: false };
    // Reuse only signed context; old authorization/age is never revived. The fresh
    // diagnostic ticket has its own nonce, parent PID and TTL, and gets no grant.
    report.stage = 'create-diagnostic-ticket';
    job = await createInstallTicket(paths.replayCache, { parentPid: process.pid, platform: 'win32', arch: 'x64', currentVersion: config.version, target: old.target, archive: old.archive, envelope: old.envelope });
    report.oldTicketAgeMs = Date.now() - old.createdAt;
    report.freshDiagnosticTicket = true;
    report.stage = 'verify-ticket-child';
    report.packageVerified = await verifyInChild(job.file);
    report.verificationRanInSeparateProcess = true;
    const env = { ...process.env };
    for (const key of Object.keys(env)) if (/^(ELECTRON_RUN_AS_NODE|WPS_BRIDGE_RESOURCES|PORTABLE_EXECUTABLE.*|CHARFLOAT_DIAGNOSTIC_.*)$/i.test(key)) delete env[key];
    env.ELECTRON_ENABLE_LOGGING = '1';
    report.stage = 'launch-installed-helper';
    child = spawn(process.execPath, ['--charfloat-update-helper', job.file], { env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
    child.stderr!.on('data', data => { stderr = (stderr + data).slice(-12000); });
    child.stdout!.on('data', data => { stdout = (stdout + data).slice(-4000); });
    child.on('exit', (code, signal) => { exit = { code, signal }; });
    let launchError: Error | undefined; child.on('error', error => { launchError = error; });
    const started = Date.now();
    while (Date.now() - started < 45000) {
      try {
        const file = path.join(job.directory, 'result.json'); await plainFile(file, 16384);
        const value = JSON.parse(fs.readFileSync(file, 'utf8'));
        if (value.nonce === job.ticket.nonce) report.helperResult = { phase: value.phase, message: clean(value.message) };
      } catch (error: any) { if (error.code !== 'ENOENT') throw error; }
      if (launchError) throw launchError;
      if (exit || ['ready', 'failed'].includes(report.helperResult?.phase)) break;
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    report.elapsedMs = Date.now() - started;
    report.stage = 'helper-observed';
    report.exitBeforeDiagnosticStop = exit;
    report.timedOut = !exit && !report.helperResult;
  } catch (error: any) { report.error = { code: error.code, message: clean(error.message) }; }
  finally {
    if (job) {
      try { await revokeInstallAuthorization(job.file, job.ticket.nonce); report.authorizationRevoked = true; }
      catch (error: any) { report.revocationError = clean(error.message); }
      report.grantFilePresent = fs.existsSync(path.join(job.directory, 'granted.json'));
      report.consumedFilePresent = fs.existsSync(path.join(job.directory, 'consumed.json'));
    }
    if (child && !exit) {
      report.stoppedOwnWaitingHelper = child.kill('SIGTERM');
      await new Promise<void>(resolve => { const timer = setTimeout(resolve, 2000); child!.once('exit', () => { clearTimeout(timer); resolve(); }); });
    }
    report.stderr = clean(stderr); report.stdout = clean(stdout);
    writeDiagnosticReport(reportFile, report);
    console.log('取证完成，未授予安装授权。请发回 handoff-console.log。');
  }
}
const execution = process.argv[2] === '--verify-ticket' ? verifyTicketMode() : main();
void execution.catch((error: any) => {
  console.error(JSON.stringify({ diagnosticFailed: true, code: error.code, message: clean(error.message), installs: false }));
  process.exitCode = 1;
});
