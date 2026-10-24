import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { initializeUpdateSigningKey } from '../scripts/init-update-signing-key.mjs';
import { checkKeyBoundaries, KEY_EXCLUSIONS, KEY_DIRECTORY, PRIVATE_KEY, assertPublicPaths } from '../scripts/lib/update-key-safety.mjs';
import { scanUpdateArtifacts } from '../scripts/check-update-secrets.mjs';
import { allowed, checkFiles } from '../scripts/lib/release-protection.mjs';
import { prepareUpdate } from '../scripts/prepare-update.js';

function fixture(t: any) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-key-safety-')); t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  execFileSync('git', ['init', '-q', root], { stdio: 'ignore' });
  fs.writeFileSync(path.join(root, '.gitignore'), '/.charfloat-local-secrets/\n');
  fs.writeFileSync(path.join(root, '.gitattributes'), '.charfloat-local-secrets export-ignore\n.charfloat-local-secrets/** export-ignore\n');
  fs.mkdirSync(path.join(root, 'scripts'));
  fs.writeFileSync(path.join(root, 'scripts/release-files.json'), JSON.stringify({ files: ['dist/main/index.js'] }));
  fs.writeFileSync(path.join(root, 'package.json'), JSON.stringify({ build: { files: [{ from: 'dist/release-app', filter: ['**/*', ...KEY_EXCLUSIONS] }, ...KEY_EXCLUSIONS] } }));
  return root;
}
test('local signing initialization requires all exclusions and refuses to rotate or overwrite', t => {
  const root = fixture(t);
  fs.writeFileSync(path.join(root, '.gitignore'), '');
  assert.throws(() => initializeUpdateSigningKey(root)); assert.equal(fs.existsSync(path.join(root, KEY_DIRECTORY)), false);
  fs.writeFileSync(path.join(root, '.gitignore'), '/.charfloat-local-secrets/\n');
  assert.equal(checkKeyBoundaries(root).gitIgnored, true);
  const first = initializeUpdateSigningKey(root), file = path.join(root, first.privateKeyPath);
  const before = createHash('sha256').update(fs.readFileSync(file)).digest('hex');
  const second = initializeUpdateSigningKey(root);
  assert.equal(first.created, true); assert.equal(second.created, false); assert.equal(first.publicKeySha256, second.publicKeySha256);
  assert.equal(createHash('sha256').update(fs.readFileSync(file)).digest('hex'), before);
  if (process.platform !== 'win32') { assert.equal(fs.statSync(file).mode & 0o777, 0o600); assert.equal(fs.statSync(path.dirname(file)).mode & 0o777, 0o700); }
  assert.equal(execFileSync('git', ['ls-files', '--others', '--exclude-standard'], { cwd: root, encoding: 'utf8' }).includes('.charfloat-local-secrets'), false);
});
test('missing private key with existing metadata and unsafe directory links do not generate a replacement', t => {
  const root = fixture(t), dir = path.join(root, KEY_DIRECTORY); fs.mkdirSync(dir, { recursive: true, mode: 0o700 }); fs.chmodSync(path.dirname(dir), 0o700);
  fs.writeFileSync(path.join(dir, 'key-info.json'), '{}');
  assert.throws(() => initializeUpdateSigningKey(root)); assert.equal(fs.existsSync(path.join(dir, PRIVATE_KEY)), false);
  if (process.platform !== 'win32') {
    const linked = fixture(t), outside = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-key-link-')); t.after(() => fs.rmSync(outside, { recursive: true, force: true }));
    fs.symlinkSync(outside, path.join(linked, '.charfloat-local-secrets'));
    assert.throws(() => initializeUpdateSigningKey(linked)); assert.equal(fs.readdirSync(outside).length, 0);
  }
});
test('artifact path gate covers dependencies, nested bundles and source archive names', () => {
  for (const name of ['.charfloat-local-secrets/update-signing/x', 'node_modules/x/.charfloat-local-secrets/key', 'resources/production-ed25519-private.pem', 'nested\\.charfloat-local-secrets\\key']) {
    assert.throws(() => assertPublicPaths([name])); assert.equal(allowed(name), false);
    assert.ok(checkFiles([name], () => Buffer.alloc(0)).some(error => error.includes('签名材料')));
  }
});
test('artifact scanner accepts public-only bytes and rejects private bytes without echoing them', async t => {
  const root = fixture(t), key = initializeUpdateSigningKey(root), file = path.join(root, key.privateKeyPath), output = path.join(root, 'output'); fs.mkdirSync(output);
  fs.copyFileSync(path.join(root, KEY_DIRECTORY, 'public-key.pem'), path.join(output, 'public.txt'));
  assert.equal((await scanUpdateArtifacts([output], file)).passed, true);
  // Only ephemeral test material is copied into an adversarial fixture, never the production private key.
  fs.copyFileSync(file, path.join(output, 'disguised.bin'));
  await assert.rejects(scanUpdateArtifacts([output], file), /包含更新签名私钥材料/);
});
test('offline publisher rejects local-secret paths and disguised private-key input before copying', async t => {
  const root = fixture(t), key = initializeUpdateSigningKey(root), file = path.join(root, key.privateKeyPath);
  const privatePem = fs.readFileSync(file, 'utf8'), disguised = path.join(root, 'wrong-installer.exe'); fs.writeFileSync(disguised, privatePem);
  const spec = { version: '2.2.4', releaseId: 1, keyId: key.keyId, expiresAt: new Date(Date.now() + 3600000).toISOString(), notes: 'Synthetic safety test', packageBases: ['https://example.test/'], assets: [{ platform: 'win32' as const, arch: 'x64' as const, kind: 'windows-bootstrap' as const, file }] };
  await assert.rejects(prepareUpdate(spec, privatePem, path.join(root, 'public-output')), /签名材料/);
  await assert.rejects(prepareUpdate({ ...spec, assets: [{ ...spec.assets[0], file: disguised }] }, privatePem, path.join(root, 'public-output')), /包含签名私钥材料/);
  assert.equal(fs.existsSync(path.join(root, 'public-output')), false);
});
test('ZIP scans finish with empty and framework-link entries and reject private bytes in ordinary names', async t => {
  const root = fixture(t), key = initializeUpdateSigningKey(root), privateFile = path.join(root, key.privateKeyPath);
  function zip(entries: [string, Buffer, number][]) {
    const locals: Buffer[] = [], central: Buffer[] = []; let offset = 0;
    for (const [name, bytes, mode] of entries) {
      let crc = 0xffffffff; for (const byte of bytes) { crc ^= byte; for (let i = 0; i < 8; i++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0); } crc = (crc ^ 0xffffffff) >>> 0;
      const text = Buffer.from(name), local = Buffer.alloc(30), header = Buffer.alloc(46);
      local.writeUInt32LE(0x04034b50); local.writeUInt16LE(20,4); local.writeUInt16LE(0x800,6); local.writeUInt32LE(crc,14); local.writeUInt32LE(bytes.length,18); local.writeUInt32LE(bytes.length,22); local.writeUInt16LE(text.length,26);
      header.writeUInt32LE(0x02014b50); header.writeUInt16LE(0x0314,4); header.writeUInt16LE(20,6); header.writeUInt16LE(0x800,8); header.writeUInt32LE(crc,16); header.writeUInt32LE(bytes.length,20); header.writeUInt32LE(bytes.length,24); header.writeUInt16LE(text.length,28); header.writeUInt32LE((mode<<16)>>>0,38); header.writeUInt32LE(offset,42);
      locals.push(local,text,bytes); central.push(header,text); offset += local.length+text.length+bytes.length;
    }
    const directory=Buffer.concat(central), end=Buffer.alloc(22); end.writeUInt32LE(0x06054b50); end.writeUInt16LE(entries.length,8); end.writeUInt16LE(entries.length,10); end.writeUInt32LE(directory.length,12); end.writeUInt32LE(offset,16);
    return Buffer.concat([...locals,directory,end]);
  }
  const entries: [string,Buffer,number][] = [['app/empty',Buffer.alloc(0),0o100644],['app/Current',Buffer.from('Versions/A'),0o120777],['app/public.pem',fs.readFileSync(path.join(root,KEY_DIRECTORY,'public-key.pem')),0o100644]];
  const good=path.join(root,'public.zip'); fs.writeFileSync(good,zip(entries)); assert.equal((await scanUpdateArtifacts([good],privateFile)).passed,true);
  const bad=path.join(root,'bad.zip'); fs.writeFileSync(bad,zip([...entries,['app/notes.txt',fs.readFileSync(privateFile),0o100644]]));
  await assert.rejects(scanUpdateArtifacts([bad],privateFile),/包含更新签名私钥材料/);
});
