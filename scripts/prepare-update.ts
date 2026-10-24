import fs from 'node:fs';
import path from 'node:path';
import { createHash, createPrivateKey, createPublicKey, sign } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { httpsUrl, readSignedManifest, type UpdateManifest } from '../src/main/update-protocol.js';
import { assertPublicPaths } from './lib/update-key-safety.mjs';

type Spec = { version: string; releaseId: number; keyId: string; publishedAt?: string; expiresAt: string; notes: string; packageBases: string[]; assets: { platform: 'darwin' | 'win32'; arch: 'arm64' | 'x64'; kind: 'mac-zip' | 'windows-zip'; file: string }[] };
/** Offline only. Caller supplies an existing signing key; never creates keys, deploys or overwrites output. */
export async function prepareUpdate(spec: Spec, privateKeyPem: string, outputDirectory: string, now = Date.now()) {
  if (!/^[A-Za-z0-9_-]{1,64}$/.test(spec.keyId) || !Array.isArray(spec.packageBases) || !spec.packageBases.length || spec.packageBases.length > 4 || !Array.isArray(spec.assets) || !spec.assets.length || spec.assets.length > 8) throw new Error('发布规格无效。');
  const bases = spec.packageBases.map(base => { const u = httpsUrl(base); if (u.search || !u.pathname.endsWith('/')) throw new Error('包目录必须是以斜杠结尾且不带查询参数的 HTTPS 地址。'); return u; });
  const key = createPrivateKey(privateKeyPem); if (key.asymmetricKeyType !== 'ed25519') throw new Error('需要已有的 Ed25519 离线发布签名密钥。');
  assertPublicPaths([outputDirectory, ...spec.assets.map(asset => asset.file)]);
  const privateDer = key.export({ format: 'der', type: 'pkcs8' });
  const privatePatterns = [Buffer.from(privateKeyPem), privateDer, Buffer.from(privateDer.toString('base64'))];
  const retained = Math.max(...privatePatterns.map(value => value.length)) - 1;
  const publicKey = createPublicKey(key).export({ format: 'pem', type: 'spki' }).toString();
  const output = path.resolve(outputDirectory); if (fs.existsSync(output)) throw new Error('输出目录已存在，未覆盖；请使用新的本地目录。');
  await fs.promises.mkdir(path.dirname(output), { recursive: true });
  const staging = await fs.promises.mkdtemp(path.join(path.dirname(output), '.update-publish-'));
  try {
    const assets: UpdateManifest['assets'] = [];
    for (const asset of spec.assets) {
      const stat = await fs.promises.lstat(asset.file);
      if (!stat.isFile() || stat.isSymbolicLink() || stat.size <= 0 || stat.size > 2 * 1024 ** 3) throw new Error('发布包必须是普通文件且不超过 2 GiB。');
      const hash = createHash('sha256'); let tail = Buffer.alloc(0);
      for await (const bytes of fs.createReadStream(asset.file)) {
        const joined = Buffer.concat([tail, bytes]);
        if (privatePatterns.some(pattern => joined.includes(pattern))) throw new Error('更新包输入包含签名私钥材料，未复制或签名发布。');
        tail = joined.subarray(Math.max(0, joined.length - retained)); hash.update(bytes);
      }
      const sha256 = hash.digest('hex');
      const filename = `CharFloat-${spec.version}-${asset.platform}-${asset.arch}-${sha256}.zip`;
      if (!/^[A-Za-z0-9.-]+$/.test(filename)) throw new Error('版本或平台不能构成安全的发布文件名。');
      const relative = `packages/${spec.version}/${filename}`; const file = path.join(staging, relative); await fs.promises.mkdir(path.dirname(file), { recursive: true });
      await fs.promises.copyFile(asset.file, file, fs.constants.COPYFILE_EXCL);
      // Hash the copied bytes as well, so a concurrently changed source can never be signed accidentally.
      const copiedHash = createHash('sha256'); for await (const bytes of fs.createReadStream(file)) copiedHash.update(bytes);
      if ((await fs.promises.stat(file)).size !== stat.size || copiedHash.digest('hex') !== sha256) throw new Error('复制时发布包发生变化，未生成清单。');
      assets.push({ platform: asset.platform, arch: asset.arch, kind: asset.kind, size: stat.size, sha256, urls: bases.map(base => new URL(relative, base).href) });
    }
    const manifest: UpdateManifest = { schema: 1, product: 'charfloat', channel: 'stable', releaseId: spec.releaseId, version: spec.version, publishedAt: spec.publishedAt || new Date(now).toISOString(), expiresAt: spec.expiresAt, notes: spec.notes, assets };
    const payload = Buffer.from(JSON.stringify(manifest));
    const envelope = JSON.stringify({ keyId: spec.keyId, payload: payload.toString('base64'), signature: sign(null, payload, key).toString('base64') }, null, 2)+'\n';
    for (const asset of assets) readSignedManifest(envelope, { feeds: [], assetOrigins: [...new Set(bases.map(u => u.origin))], publicKeys: { [spec.keyId]: publicKey } }, { version: spec.version, platform: asset.platform, arch: asset.arch }, undefined, now);
    await fs.promises.writeFile(path.join(staging, 'stable.json'), envelope, { flag: 'wx' });
    await fs.promises.writeFile(path.join(staging, 'public-key.txt'), publicKey, { flag: 'wx' });
    await fs.promises.rename(staging, output);
    return { version: spec.version, releaseId: spec.releaseId, files: assets.length, keyId: spec.keyId };
  } catch (error) { await fs.promises.rm(staging, { recursive: true, force: true }); throw error; }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [specFile, keyFile, output] = process.argv.slice(2);
  if (!specFile || !keyFile || !output) { console.error('用法：node --import tsx scripts/prepare-update.ts <spec.json> <已有离线私钥文件> <新的本地输出目录>'); process.exitCode = 1; }
  else {
    try { console.log(JSON.stringify(await prepareUpdate(JSON.parse(fs.readFileSync(specFile, 'utf8')), fs.readFileSync(keyFile, 'utf8'), output))); }
    catch { console.error('发布目录生成失败；未部署、未创建密钥、未覆盖已有输出。请检查规格、文件和签名密钥。'); process.exitCode = 1; }
  }
}
