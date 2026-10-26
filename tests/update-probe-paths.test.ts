import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { execFileSync } from 'node:child_process';
import { checkedProbeArchive } from '../scripts/lib/update-probe-paths.js';

function fixture(t: any) {
  const root = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat probe paths ')));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const cache = path.join(root, 'cache'); fs.mkdirSync(cache);
  const archive = path.join(cache, 'inside.zip'); fs.writeFileSync(archive, 'synthetic');
  return { root, cache, archive };
}
test('diagnostic containment accepts owned files but rejects outside paths, junctions and hard links', async t => {
  const f = fixture(t), result = await checkedProbeArchive(f.cache, f.archive);
  assert.equal(result.relative, 'inside.zip'); assert.equal(result.replayCache, f.cache);
  const outside = path.join(f.root, 'outside'); fs.mkdirSync(outside); const file = path.join(outside, 'other.zip'); fs.writeFileSync(file, 'keep');
  await assert.rejects(checkedProbeArchive(f.cache, file), /不在本轮缓存/);
  const link = path.join(f.cache, 'linked'); fs.symlinkSync(outside, link, process.platform === 'win32' ? 'junction' : 'dir');
  await assert.rejects(checkedProbeArchive(f.cache, path.join(link, 'other.zip')), /链接/);
  fs.linkSync(f.archive, path.join(f.cache, 'alias.zip'));
  await assert.rejects(checkedProbeArchive(f.cache, f.archive), /独立的普通文件/);
  assert.equal(fs.readFileSync(file, 'utf8'), 'keep');
});
test('Windows diagnostic compares physical names while replaying the original 8.3 spelling', { skip: process.platform !== 'win32' }, async t => {
  const f = fixture(t), physicalCache = await fs.promises.realpath(f.cache);
  const short = execFileSync('cmd.exe', ['/d', '/u', '/c', 'for %I in ("%CHARFLOAT_PROBE_PATH%") do @echo %~sI'], { env: { ...process.env, CHARFLOAT_PROBE_PATH: f.cache }, windowsHide: true, encoding: 'utf16le' }).trim();
  const archive = path.join(short, 'inside.zip'), result = await checkedProbeArchive(physicalCache, archive);
  assert.equal(result.relative, 'inside.zip'); assert.equal(result.replayCache, short);
  assert.equal(result.archivePhysical, await fs.promises.realpath(f.archive));
  t.diagnostic(short === physicalCache ? 'volume returned no distinct short spelling' : 'actual short/long alias pair checked');
});
