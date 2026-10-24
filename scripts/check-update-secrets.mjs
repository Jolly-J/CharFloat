// No key generation or key-content output. Paths are checked before reading any artifact.
import fs from 'node:fs';
import path from 'node:path';
import { createPrivateKey } from 'node:crypto';
import { createGunzip } from 'node:zlib';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import yauzl from 'yauzl';
import asar from '@electron/asar';
import { ROOT, KEY_DIRECTORY, PRIVATE_KEY, assertPublicPaths, checkKeyBoundaries } from './lib/update-key-safety.mjs';

async function scanStream(stream, patterns) {
  let tail = Buffer.alloc(0), total = 0; const retained = Math.max(1, ...patterns.map(p => p.length)) - 1;
  // yauzl's legacy entry stream needs explicit completion/error handling, including destroy on rejection.
  await new Promise((resolve, reject) => {
    let finished = false;
    const finish = error => {
      if (finished) return; finished = true;
      stream.removeListener('data', read);
      if (error) { stream.destroy(); reject(error); } else resolve();
    };
    const read = bytes => {
      try {
        total += bytes.length;
        if (total > 2 * 1024 ** 3) throw new Error('产物超过本地签名材料扫描上限。');
        const joined = Buffer.concat([tail, bytes]);
        if (patterns.some(pattern => joined.includes(pattern))) throw new Error('产物包含更新签名私钥材料，已拒绝交付。');
        tail = joined.subarray(Math.max(0, joined.length - retained));
      } catch (error) { finish(error); }
    };
    stream.on('data', read); stream.on('error', finish); stream.once('end', () => finish());
    stream.once('close', () => { if (!finished) finish(new Error('产物读取提前结束，未完成扫描。')); });
  });
}
export async function scanUpdateArtifacts(targets, privateFile) {
  let patterns = [];
  if (privateFile) {
    const st = fs.lstatSync(privateFile);
    if (!st.isFile() || st.isSymbolicLink() || st.nlink !== 1 || st.size > 16384) throw new Error('本地签名文件类型异常，未扫描。');
    const pem = fs.readFileSync(privateFile), key = createPrivateKey(pem);
    if (key.asymmetricKeyType !== 'ed25519') throw new Error('本地更新密钥类型不匹配。');
    const der = key.export({ format: 'der', type: 'pkcs8' });
    patterns = [pem, der, Buffer.from(der.toString('base64'))];
  }
  let files = 0;
  async function scan(file) {
    assertPublicPaths([file]);
    const st = fs.lstatSync(file);
    if (st.isSymbolicLink()) throw new Error('产物扫描不跟随目录或文件链接。');
    if (st.isDirectory()) { for (const name of fs.readdirSync(file)) await scan(path.join(file, name)); return; }
    if (!st.isFile()) throw new Error('产物不是普通文件。');
    files++;
    if (/\.(?:tgz|tar\.gz|tar)$/.test(file)) {
      const names = execFileSync('tar', ['-tf', file], { encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 }).trim().split('\n');
      assertPublicPaths(names);
      const input = fs.createReadStream(file); await scanStream(/\.gz$|\.tgz$/.test(file) ? input.pipe(createGunzip()) : input, patterns); return;
    }
    if (file.endsWith('.zip')) {
      const zip = await new Promise((resolve, reject) => yauzl.open(file, { lazyEntries: true }, (error, value) => error ? reject(error) : resolve(value)));
      await new Promise((resolve, reject) => {
        const fail = error => { zip.close(); reject(error); };
        zip.on('error', fail); zip.on('end', resolve);
        zip.on('entry', entry => { void (async () => {
          assertPublicPaths([entry.fileName]);
          // Empty entries have no private bytes; their paths are still checked above.
          if (!entry.fileName.endsWith('/') && entry.uncompressedSize > 0) {
            const input = await new Promise((res, rej) => zip.openReadStream(entry, (error, value) => error ? rej(error) : res(value)));
            await scanStream(input, patterns);
          }
          zip.readEntry();
        })().catch(fail); }); zip.readEntry();
      }); return;
    }
    if (file.endsWith('.asar')) assertPublicPaths(asar.listPackage(file));
    await scanStream(fs.createReadStream(file), patterns);
  }
  for (const file of targets) await scan(path.resolve(file));
  return { passed: true, files, privateMaterialChecked: patterns.length > 0 };
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    const boundaries = checkKeyBoundaries(); const targets = process.argv.slice(2);
    const file = path.join(ROOT, KEY_DIRECTORY, PRIVATE_KEY);
    const result = targets.length ? await scanUpdateArtifacts(targets, fs.existsSync(file) ? file : undefined) : undefined;
    console.log(JSON.stringify({ boundaries, ...(result ? { artifacts: result } : {}) }));
  } catch { console.error('更新签名隔离检查失败；未输出密钥内容，请检查排除规则与本地产物。'); process.exitCode = 1; }
}
