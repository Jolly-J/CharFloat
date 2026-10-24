import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { generateKeyPairSync, createHash, sign } from 'node:crypto';
import { deflateRawSync } from 'node:zlib';
import { extractMacArchive, stageMacUpdate, commitMacUpdate, validateMacBundle, recoverMacUpdate } from '../src/main/update-mac.js';
import { createInstallTicket, readInstallTicket, waitForParentExit, grantInstallAuthorization, revokeInstallAuthorization, consumeInstallAuthorization, writeInstallState } from '../src/main/update-install.js';
import { handoffUpdate } from '../src/main/update-handoff.js';
import { EventEmitter } from 'node:events';
import { retargetUpdateConfig } from '../src/main/update-config.js';
import { renderCodexBlock } from '../src/main/mcp-adapters.js';

type Entry = [string, Buffer, number?, number?];
function zip(entries: Entry[]) {
  const locals: Buffer[] = [], centers: Buffer[] = []; let offset = 0;
  for (const [name, data, mode = 0o100644, method = 8] of entries) {
    const n = Buffer.from(name), compressed = method === 0 ? data : deflateRawSync(data), l = Buffer.alloc(30), h = Buffer.alloc(46);
    l.writeUInt32LE(0x04034b50); l.writeUInt16LE(20, 4); l.writeUInt16LE(0x800, 6); l.writeUInt16LE(8, 8); l.writeUInt32LE(compressed.length, 18); l.writeUInt32LE(data.length, 22); l.writeUInt16LE(n.length, 26);
    h.writeUInt32LE(0x02014b50); h.writeUInt16LE(0x0314, 4); h.writeUInt16LE(20, 6); h.writeUInt16LE(0x800, 8); h.writeUInt16LE(8, 10); h.writeUInt32LE(compressed.length, 20); h.writeUInt32LE(data.length, 24); h.writeUInt16LE(n.length, 28); h.writeUInt32LE((mode << 16) >>> 0, 38); h.writeUInt32LE(offset, 42);
    l.writeUInt16LE(method, 8); h.writeUInt16LE(method, 10);
    locals.push(l, n, compressed); centers.push(h, n); offset += l.length + n.length + compressed.length;
  }
  const central = Buffer.concat(centers), end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50); end.writeUInt16LE(entries.length, 8); end.writeUInt16LE(entries.length, 10); end.writeUInt32LE(central.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, central, end]);
}
function appEntries(version = '2.2.4', arch = 'arm64', id = 'com.antigravity.charfloat'): Entry[] {
  const header = Buffer.alloc(64); header.writeUInt32BE(0xcffaedfe); header.writeUInt32LE(arch === 'arm64' ? 0x0100000c : 0x01000007, 4);
  return [
    ['字浮 CharFloat.app/Contents/Info.plist', Buffer.from(`<?xml version="1.0"?><plist version="1.0"><dict><key>CFBundleIdentifier</key><string>${id}</string><key>CFBundleShortVersionString</key><string>${version}</string><key>CFBundleExecutable</key><string>CharFloat</string></dict></plist>`)],
    ['字浮 CharFloat.app/Contents/MacOS/CharFloat', header, 0o100755],
    ['字浮 CharFloat.app/Contents/Frameworks/Mock.framework/Versions/A/file', Buffer.from('synthetic framework')],
    ['字浮 CharFloat.app/Contents/Frameworks/Mock.framework/Versions/Current', Buffer.from('A'), 0o120777],
    ['字浮 CharFloat.app/Contents/Frameworks/Mock.framework/file', Buffer.from('Versions/Current/file'), 0o120777],
  ];
}
async function fixture(t: any) {
  const dir = await fs.promises.realpath(fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-install-')));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const oldStage = path.join(dir, 'old'); fs.mkdirSync(oldStage);
  const oldZip = path.join(dir, 'old.zip'); fs.writeFileSync(oldZip, zip(appEntries('2.2.3')));
  const old = await extractMacArchive(oldZip, oldStage), target = path.join(dir, '字浮 CharFloat.app'); fs.renameSync(old, target);
  const archive = path.join(dir, 'new.zip'); fs.writeFileSync(archive, zip(appEntries()));
  return { dir, archive, target };
}
test('Mac same-volume stage preserves framework links, checks identity/version/architecture and commits with backup', async t => {
  const f = await fixture(t), data = path.join(f.dir, 'user-data'); fs.writeFileSync(data, 'keep');
  const job = await stageMacUpdate(f.archive, f.target, '2.2.3', '2.2.4', 'arm64');
  await validateMacBundle(f.target, '2.2.3', 'arm64');
  await commitMacUpdate(job.journal, job.tx);
  await validateMacBundle(f.target, '2.2.4', 'arm64'); await validateMacBundle(job.tx.backup, '2.2.3', 'arm64');
  assert.equal(fs.readFileSync(path.join(f.target, 'Contents/Frameworks/Mock.framework/file'), 'utf8'), 'synthetic framework');
  assert.equal(fs.readFileSync(data, 'utf8'), 'keep');
  assert.equal(JSON.parse(fs.readFileSync(job.journal, 'utf8')).phase, 'committed');
});
test('Mac stages STORE and DEFLATE entries, including empty files and stored framework links', { timeout: 5000 }, async t => {
  const f = await fixture(t);
  const entries = appEntries().map(([name, bytes, mode]) => [name, bytes, mode, 0] as Entry);
  entries.push(['字浮 CharFloat.app/Contents/empty', Buffer.alloc(0), 0o100644, 0], ['字浮 CharFloat.app/Contents/compressed', Buffer.from('deflated')]);
  fs.writeFileSync(f.archive, zip(entries));
  const job = await stageMacUpdate(f.archive, f.target, '2.2.3', '2.2.4', 'arm64');
  assert.equal(job.tx.phase, 'prepared');
  assert.equal(fs.statSync(path.join(job.tx.staged, 'Contents/empty')).size, 0);
  assert.equal(fs.readFileSync(path.join(job.tx.staged, 'Contents/Frameworks/Mock.framework/file'), 'utf8'), 'synthetic framework');
});
test('Mac commit failure after moving old app rolls back; retry succeeds without deleting backup', async t => {
  const f = await fixture(t), job = await stageMacUpdate(f.archive, f.target, '2.2.3', '2.2.4', 'arm64');
  await assert.rejects(commitMacUpdate(job.journal, job.tx, async () => { throw new Error('synthetic file lock'); }), /file lock/);
  await validateMacBundle(f.target, '2.2.3', 'arm64'); assert.equal(job.tx.phase, 'rolled-back');
  await commitMacUpdate(job.journal, job.tx); await validateMacBundle(f.target, '2.2.4', 'arm64');
});
test('Mac interrupted rename is recovered by the recorded transaction before any new install', async t => {
  const f = await fixture(t), job = await stageMacUpdate(f.archive, f.target, '2.2.3', '2.2.4', 'arm64');
  fs.renameSync(f.target, job.tx.backup); job.tx.phase = 'moving';
  await assert.rejects(commitMacUpdate(job.journal, job.tx), /旧应用已恢复/);
  await validateMacBundle(f.target, '2.2.3', 'arm64');
});
test('offline journal recovery restores only a missing application and refuses a present target', async t => {
  const f = await fixture(t), job = await stageMacUpdate(f.archive, f.target, '2.2.3', '2.2.4', 'arm64');
  fs.renameSync(f.target, job.tx.backup);
  assert.equal(await recoverMacUpdate(job.journal), f.target); await validateMacBundle(f.target, '2.2.3', 'arm64');
  await assert.rejects(recoverMacUpdate(job.journal), /仍存在/);
});
test('Mac rejects wrong identity, version and architecture before changing installed app', async t => {
  const f = await fixture(t);
  for (const entries of [appEntries('2.2.5'), appEntries('2.2.4', 'x64'), appEntries('2.2.4', 'arm64', 'other.app')]) {
    fs.writeFileSync(f.archive, zip(entries));
    await assert.rejects(stageMacUpdate(f.archive, f.target, '2.2.3', '2.2.4', 'arm64'));
    await validateMacBundle(f.target, '2.2.3', 'arm64');
  }
});
test('Mac archive traversal, escaping links, link children and duplicate/case paths never write outside staging', async t => {
  const f = await fixture(t), outside = path.join(f.dir, 'outside'); fs.writeFileSync(outside, 'keep');
  for (const extra of [
    [['../outside', Buffer.from('bad')]],
    [['字浮 CharFloat.app/escape', Buffer.from('../../outside'), 0o120777]],
    [['字浮 CharFloat.app/link', Buffer.from('Contents'), 0o120777], ['字浮 CharFloat.app/link/evil', Buffer.from('bad')]],
    [['字浮 CharFloat.app/contents/Info.plist', Buffer.from('duplicate')]],
    [['字浮 CharFloat.app/dangling', Buffer.from('missing'), 0o120777]],
  ] as Entry[][]) {
    fs.writeFileSync(f.archive, zip([...appEntries(), ...extra]));
    await assert.rejects(stageMacUpdate(f.archive, f.target, '2.2.3', '2.2.4', 'arm64'));
    assert.equal(fs.readFileSync(outside, 'utf8'), 'keep'); await validateMacBundle(f.target, '2.2.3', 'arm64');
  }
});
test('signed install ticket is verified again and refuses tamper, wrong system, downgrade and unsafe cache', async t => {
  const f = await fixture(t), cache = path.join(f.dir, 'cache'); fs.mkdirSync(cache, { mode: 0o700 });
  const keys = generateKeyPairSync('ed25519'), trust = { feeds: ['https://example.test/stable.json'], assetOrigins: ['https://example.test'], publicKeys: { fixture: keys.publicKey.export({ format: 'pem', type: 'spki' }).toString() } };
  const bytes = fs.readFileSync(f.archive), now = Date.now();
  const payload = Buffer.from(JSON.stringify({ schema: 1, product: 'charfloat', channel: 'stable', version: '2.2.4', releaseId: 1, publishedAt: new Date(now).toISOString(), expiresAt: new Date(now + 86400000).toISOString(), notes: '', assets: [{ platform: 'darwin', arch: 'arm64', kind: 'mac-zip', size: bytes.length, sha256: createHash('sha256').update(bytes).digest('hex'), urls: ['https://example.test/new.zip'] }] }));
  const envelope = JSON.stringify({ keyId: 'fixture', payload: payload.toString('base64'), signature: sign(null, payload, keys.privateKey).toString('base64') });
  const job = await createInstallTicket(cache, { parentPid: process.ppid, platform: 'darwin', arch: 'arm64', target: f.target, archive: f.archive, currentVersion: '2.2.3', envelope });
  assert.equal((await readInstallTicket(job.file, trust, 'darwin', 'arm64')).selected.manifest.version, '2.2.4');
  await assert.rejects(readInstallTicket(job.file, trust, 'win32', 'x64'), /不匹配/);
  fs.appendFileSync(f.archive, 'tamper'); await assert.rejects(readInstallTicket(job.file, trust, 'darwin', 'arm64'), /已改变/);
  await assert.rejects(waitForParentExit(process.pid, 20), /超时/);
});
test('Windows JSON/TOML/opencode retarget only the owned entry and preserve custom options and unrelated entries', async t => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-config-')); t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const previous = 'C:\\Users\\用户\\Programs\\CharFloat\\2.2.3-aaaaaaaaaaaa', next = 'C:\\Users\\用户\\Programs\\CharFloat\\2.2.4-bbbbbbbbbbbb';
  const entry = { command: path.win32.join(previous, '字浮 CharFloat.exe'), args: [path.win32.join(previous, 'resources/app.asar.unpacked/dist/bridge/cli.cjs'), '--custom'], env: { WPS_BRIDGE_RESOURCES: path.win32.join(previous, 'resources/app.asar'), KEEP: 'yes' } };
  for (const kind of ['json-mcp', 'json-opencode', 'toml-codex'] as const) {
    const file = path.join(dir, kind);
    const raw = kind === 'toml-codex' ? ['# keep comment', ...renderCodexBlock(entry), '', '[mcp_servers.other]', `command = ${JSON.stringify(entry.command)}`, ''].join('\n') : JSON.stringify(kind === 'json-mcp' ? { mcpServers: { charfloat: { ...entry, disabled: true }, other: entry }, outside: previous } : { mcp: { charfloat: { command: [entry.command, ...entry.args], environment: entry.env, enabled: false }, other: entry }, outside: previous });
    fs.writeFileSync(file, raw);
    assert.equal(retargetUpdateConfig(file, kind, previous, next), true, kind);
    const updated = fs.readFileSync(file, 'utf8'); assert.ok(updated.includes('2.2.4-bbbbbbbbbbbb')); assert.ok(updated.includes('2.2.3-aaaaaaaaaaaa')); assert.ok(updated.includes('--custom')); assert.ok(updated.includes('yes'));
    assert.equal(retargetUpdateConfig(file, kind, previous, next), false);
    assert.ok(fs.readdirSync(dir).some(name => name.startsWith(kind + '.backup-')));
  }
  const malformed = path.join(dir, 'broken'); fs.writeFileSync(malformed, '{bad'); assert.throws(() => retargetUpdateConfig(malformed, 'json-mcp', previous, next)); assert.equal(fs.readFileSync(malformed, 'utf8'), '{bad');
});
test('handoff timeout permanently revokes late installer; retry uses a fresh one-use authorization', async t => {
  const dir = await fs.promises.realpath(fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-revoke-'))); t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  let ticketFile = '';
  const spawnProcess: any = (executable: string, args: string[], options: any) => {
    assert.equal(executable, process.execPath, 'Windows online update must use the installed runtime, never execute a downloaded installer');
    assert.equal(args[0], '--charfloat-update-helper'); assert.equal(options.windowsHide, true);
    ticketFile = args[1]; const child: any = new EventEmitter(); child.pid = process.pid; child.unref = () => {}; child.kill = () => true; return child;
  };
  await assert.rejects(handoffUpdate({ cache: dir, artifact: { file: path.join(dir, 'synthetic.exe'), envelope: '{}' }, currentVersion: '2.2.3', target: path.join(dir, 'old'), helper: '', platform: 'win32', arch: 'x64', spawnProcess, timeoutMs: 10 }), /未能就绪/);
  const ticket = JSON.parse(fs.readFileSync(ticketFile, 'utf8'));
  // Late ready / later ordinary parent exit cannot create or consume a grant for this attempt.
  await assert.rejects(writeInstallState(path.dirname(ticketFile), { phase: 'ready', nonce: ticket.nonce }), /撤销/);
  await assert.rejects(consumeInstallAuthorization(ticketFile, ticket.nonce), /撤销/);
  await assert.rejects(grantInstallAuthorization(ticketFile, ticket.nonce), /撤销/);
  const retry = await createInstallTicket(dir, { parentPid: process.ppid, platform: 'win32', arch: 'x64', currentVersion: '2.2.3', target: path.join(dir, 'old'), archive: path.join(dir, 'synthetic.exe'), envelope: '{}' });
  await grantInstallAuthorization(retry.file, retry.ticket.nonce);
  const guard = await consumeInstallAuthorization(retry.file, retry.ticket.nonce); await guard();
  await assert.rejects(consumeInstallAuthorization(retry.file, retry.ticket.nonce), /已使用/);
  await revokeInstallAuthorization(retry.file, retry.ticket.nonce); await assert.rejects(guard(), /撤销/);
  const successSpawn: any = (_exe: string, args: string[]) => {
    const ticket = JSON.parse(fs.readFileSync(args[1], 'utf8'));
    fs.writeFileSync(path.join(path.dirname(args[1]), 'result.json'), JSON.stringify({ phase: 'ready', nonce: ticket.nonce }));
    const child: any = new EventEmitter(); child.pid = process.pid; child.unref = () => {}; child.kill = () => true; return child;
  };
  const success = await handoffUpdate({ cache: dir, artifact: { file: path.join(dir, 'synthetic.exe'), envelope: '{}' }, currentVersion: '2.2.3', target: path.join(dir, 'old'), helper: '', platform: 'win32', arch: 'x64', spawnProcess: successSpawn, timeoutMs: 1000 });
  const consumers = await Promise.allSettled([consumeInstallAuthorization(success.file, success.ticket.nonce), consumeInstallAuthorization(success.file, success.ticket.nonce)]);
  assert.equal(consumers.filter(r => r.status === 'fulfilled').length, 1);
  assert.equal(consumers.filter(r => r.status === 'rejected').length, 1);
});
