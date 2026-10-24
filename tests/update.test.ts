import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createHash, generateKeyPairSync, sign } from 'node:crypto';
import { UpdateService } from '../src/main/update-service.js';
import { readSignedManifest, compareVersion, type UpdateManifest, type UpdateTrust } from '../src/main/update-protocol.js';
import { prepareUpdate } from '../scripts/prepare-update.js';

// Ephemeral fixture key, never persisted or bundled in a release.
const keys = generateKeyPairSync('ed25519');
const publicKey = keys.publicKey.export({ format: 'pem', type: 'spki' }).toString();
const now = Date.now(), bytes = Buffer.from('MZ-synthetic-update-fixture');
const trust: UpdateTrust = { feeds: ['https://updates.example.test/stable.json'], assetOrigins: ['https://updates.example.test'], publicKeys: { fixture: publicKey } };
const context = { version: '2.2.3', platform: 'win32', arch: 'x64' };
const manifest = (): UpdateManifest => ({ schema: 1, product: 'charfloat', channel: 'stable', releaseId: 1, version: '2.2.4', publishedAt: new Date(now).toISOString(), expiresAt: new Date(now + 86400000).toISOString(), notes: '合成测试版本', assets: [{ platform: 'win32', arch: 'x64', kind: 'windows-zip', urls: ['https://updates.example.test/packages/2.2.4/fixture.zip'], size: bytes.length, sha256: createHash('sha256').update(bytes).digest('hex') }] });
function envelope(value: unknown = manifest()) { const payload = Buffer.from(JSON.stringify(value)); return JSON.stringify({ keyId: 'fixture', payload: payload.toString('base64'), signature: sign(null, payload, keys.privateKey).toString('base64') }); }
function fixture(options: { value?: UpdateManifest; request?: typeof fetch; trust?: UpdateTrust } = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-update-')); const config = path.join(dir, 'config.json'); fs.writeFileSync(config, '{"keep":"用户配置"}');
  const calls: { url: string; init?: RequestInit }[] = []; const states: any[] = [];
  const request: typeof fetch = options.request || (async (url, init) => { calls.push({ url: String(url), init }); return String(url).endsWith('stable.json') ? new Response(envelope(options.value)) : new Response(bytes); });
  const service = new UpdateService({ trust: options.trust || trust, context, cacheDirectory: path.join(dir, 'cache'), now: () => now, fetch: request, onState: state => states.push(state) });
  return { service, dir, calls, states, config, close: () => fs.rmSync(dir, { recursive: true, force: true }) };
}

test('unconfigured/invalid trust never contacts a source or claims latest', async () => {
  for (const config of [{ feeds: [], assetOrigins: [], publicKeys: {} }, { ...trust, feeds: ['http://updates.example.test/stable.json'] }]) {
    const f = fixture({ trust: config }); try { const state = await f.service.check(); assert.ok(['unconfigured', 'failed'].includes(state.phase)); assert.equal(f.calls.length, 0); assert.equal(fs.existsSync(path.join(f.dir, 'cache')), false); } finally { f.close(); }
  }
});
test('version comparison is numeric and rejects ambiguous versions', () => {
  assert.equal(compareVersion('2.10.0', '2.9.9'), 1);
  for (const v of ['2.02.3', '2.2', '2.2.4-beta', '../2.2.4']) assert.throws(() => compareVersion(v, '2.2.3'));
});
test('signature, key, expiry, product, platform, architecture, kind and URL are mandatory', () => {
  const valid = envelope(); assert.equal(readSignedManifest(valid, trust, context, undefined, now).manifest.version, '2.2.4');
  const modified = JSON.parse(valid); modified.signature = Buffer.alloc(64).toString('base64'); assert.throws(() => readSignedManifest(JSON.stringify(modified), trust, context, undefined, now), /签名无效/);
  modified.keyId = 'unknown'; assert.throws(() => readSignedManifest(JSON.stringify(modified), trust, context, undefined, now), /未信任/);
  for (const change of [
    (m: any) => { m.expiresAt = new Date(now - 1000).toISOString(); },
    (m: any) => { m.product = 'another-app'; },
    (m: any) => { m.version = '2.2.2'; },
    (m: any) => { m.assets[0].arch = 'arm64'; },
    (m: any) => { m.assets[0].kind = 'mac-zip'; },
    (m: any) => { m.assets[0].kind = 'windows-bootstrap'; },
    (m: any) => { m.assets[0].urls = ['http://updates.example.test/file']; },
    (m: any) => { m.assets[0].urls = ['https://attacker.example.test/file']; },
    (m: any) => { m.assets[0].urls = ['https://user:password@updates.example.test/file']; },
    (m: any) => { m.assets.push(m.assets[0]); },
  ]) { const m = manifest(); change(m); assert.throws(() => readSignedManifest(envelope(m), trust, context, undefined, now)); }
});
test('a known update remains visible through failed checks, cancellation and retry; checking a ready package keeps it ready', async () => {
  let offline = false;
  const f = fixture({ request: async url => { if (offline) throw new Error('offline'); return String(url).endsWith('stable.json') ? new Response(envelope()) : new Response(bytes); } });
  try {
    await f.service.check(); assert.equal(f.service.getState().availableVersion, '2.2.4');
    offline = true; assert.equal((await f.service.check()).phase, 'failed'); assert.equal(f.service.getState().availableVersion, '2.2.4');
    offline = false; await f.service.download();
    const file = (await f.service.verifiedDownload()).file; assert.equal(path.extname(file), '.zip');
    assert.equal((await f.service.check()).phase, 'ready'); assert.equal((await f.service.verifiedDownload()).file, file);
    f.service.preparingInstall(); assert.equal(f.service.getState().phase, 'installing');
    f.service.failedInstall('测试交接失败'); assert.equal(f.service.getState().phase, 'ready'); assert.match(f.service.getState().message, /交接失败/);
  } finally { f.close(); }
});
test('check never downloads; explicit download produces a reverified file and preserves user data', async () => {
  const f = fixture(); try {
    assert.equal((await f.service.check()).phase, 'available'); assert.equal(f.calls.length, 1);
    assert.equal((await f.service.download()).phase, 'ready'); assert.equal(f.calls.length, 2);
    for (const call of f.calls) { assert.equal(call.init?.credentials, 'omit'); assert.equal(call.init?.redirect, 'manual'); assert.equal(call.init?.body, undefined); }
    const ready = await f.service.verifiedDownload(); assert.deepEqual(fs.readFileSync(ready.file), bytes);
    assert.equal(fs.readFileSync(f.config, 'utf8'), '{"keep":"用户配置"}');
    assert.match(f.service.getState().message, /重启并更新/);
  } finally { f.close(); }
});
test('wrong hash is discarded, pending directory removed, and a later explicit retry can succeed', async () => {
  let bad = true; const f = fixture({ request: async url => String(url).endsWith('stable.json') ? new Response(envelope()) : new Response(bad ? Buffer.alloc(bytes.length) : bytes) });
  try {
    await f.service.check(); assert.equal((await f.service.download()).phase, 'failed');
    assert.match(f.service.getState().message, /校验失败/); assert.equal(fs.readdirSync(path.join(f.dir, 'cache')).some(n => n.startsWith('.pending-') || n.endsWith('.zip')), false);
    bad = false; assert.equal((await f.service.download()).phase, 'ready');
  } finally { f.close(); }
});
test('size mismatch is rejected before a package can be used', async () => {
  const f = fixture({ request: async url => String(url).endsWith('stable.json') ? new Response(envelope()) : new Response(Buffer.alloc(bytes.length + 1), { headers: { 'content-length': String(bytes.length + 1) } }) });
  try { await f.service.check(); assert.equal((await f.service.download()).phase, 'failed'); await assert.rejects(f.service.verifiedDownload()); } finally { f.close(); }
});
test('only configured HTTPS fallback is used after a disconnected source', async () => {
  const m = manifest(); m.assets[0].urls.push('https://updates.example.test/packages/backup.zip'); let requests = 0;
  const f = fixture({ value: m, request: async url => { if (String(url).endsWith('stable.json')) return new Response(envelope(m)); requests++; if (requests === 1) throw new TypeError('offline'); return new Response(bytes); } });
  try { await f.service.check(); assert.equal((await f.service.download()).phase, 'ready'); assert.equal(requests, 2); } finally { f.close(); }
});
test('unapproved redirect is blocked without following it', async () => {
  let calls = 0; const f = fixture({ request: async () => { calls++; return new Response(null, { status: 302, headers: { location: 'https://attacker.example.test/file' } }); } });
  try { assert.equal((await f.service.check()).phase, 'failed'); assert.equal(calls, 1); } finally { f.close(); }
});
test('cancel a partial download, discard only its staging data, retry without touching config', async () => {
  let slow = true, service!: UpdateService;
  const f = fixture({ request: async url => String(url).endsWith('stable.json') ? new Response(envelope()) : slow ? new Response(new ReadableStream({ start(controller) { controller.enqueue(bytes.subarray(0, 3)); queueMicrotask(() => service.cancel()); } })) : new Response(bytes) }); service = f.service;
  try {
    await service.check(); assert.equal((await service.download()).phase, 'cancelled');
    assert.equal(fs.readdirSync(path.join(f.dir, 'cache')).some(n => n.startsWith('.pending-')), false);
    assert.equal(fs.readFileSync(f.config, 'utf8'), '{"keep":"用户配置"}');
    slow = false; assert.equal((await service.download()).phase, 'ready');
  } finally { f.close(); }
});
test('shutdown cancels an unfinished manifest response and refuses new work', async () => {
  const f = fixture({ request: async () => new Response(new ReadableStream()) });
  try {
    const checking = f.service.check(); await new Promise(r => setTimeout(r, 20)); await f.service.shutdown();
    assert.equal((await checking).phase, 'cancelled'); await assert.rejects(f.service.check(), /正在退出/);
  } finally { f.close(); }
});
test('trusted highest release survives restart and rejects replay / version rollback', async () => {
  let m = manifest(); const f = fixture({ request: async () => new Response(envelope(m)) });
  try {
    assert.equal((await f.service.check()).phase, 'available');
    m = manifest(); m.releaseId = 2; m.version = '2.2.3';
    const restarted = new UpdateService({ trust, context, cacheDirectory: path.join(f.dir, 'cache'), now: () => now, fetch: async () => new Response(envelope(m)) });
    assert.equal((await restarted.check()).phase, 'failed'); assert.match(restarted.getState().message, /回退/);
    m = manifest(); m.notes = 'same number changed'; assert.equal((await f.service.check()).phase, 'failed');
  } finally { f.close(); }
});
test('tampering after download is detected again before any reveal/install handoff', async () => {
  const f = fixture(); try {
    await f.service.check(); await f.service.download(); const ready = await f.service.verifiedDownload(); fs.writeFileSync(ready.file, Buffer.alloc(bytes.length));
    await assert.rejects(f.service.verifiedDownload(), /已改变/);
  } finally { f.close(); }
});
test('cache symlink is not followed', { skip: process.platform === 'win32' }, async () => {
  const f = fixture(); try {
    fs.symlinkSync(f.dir, path.join(f.dir, 'cache')); assert.equal((await f.service.check()).phase, 'failed'); assert.equal(fs.existsSync(path.join(f.dir, 'seen.json')), false);
  } finally { f.close(); }
});
test('offline publisher signs copied immutable packages, exports only public key and refuses overwrite', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-update-publish-')); const file = path.join(dir, 'synthetic.zip'), output = path.join(dir, 'site'); fs.writeFileSync(file, bytes);
  const spec = { version: '2.2.4', releaseId: 1, keyId: 'fixture', expiresAt: new Date(now + 86400000).toISOString(), notes: '合成发布', packageBases: ['https://updates.example.test/updates/'], assets: [{ platform: 'win32' as const, arch: 'x64' as const, kind: 'windows-zip' as const, file }] };
  const privatePem = keys.privateKey.export({ format: 'pem', type: 'pkcs8' }).toString();
  try {
    await prepareUpdate(spec, privatePem, output, now);
    const signed = fs.readFileSync(path.join(output, 'stable.json'), 'utf8'); const m = readSignedManifest(signed, trust, context, undefined, now); assert.equal(m.asset.sha256, createHash('sha256').update(bytes).digest('hex'));
    assert.equal(fs.readFileSync(path.join(output, 'public-key.txt'), 'utf8').includes('PRIVATE KEY'), false);
    const packageFile = path.join(output, 'packages', '2.2.4', fs.readdirSync(path.join(output, 'packages', '2.2.4'))[0]); assert.deepEqual(fs.readFileSync(packageFile), bytes);
    await assert.rejects(prepareUpdate(spec, privatePem, output, now), /未覆盖/); assert.equal(fs.readFileSync(path.join(output, 'stable.json'), 'utf8'), signed);
    assert.equal(fs.readdirSync(output).some(n => /private|secret/i.test(n)), false);
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});

test('an unpublished feed reports HTTP 404 and never claims the client is current', async () => {
  const f = fixture({ request: async () => new Response('', { status: 404 }) });
  try { const result = await f.service.check(); assert.equal(result.phase, 'failed'); assert.match(result.message, /HTTP 404/); assert.equal(result.availableVersion, undefined); } finally { f.close(); }
});
