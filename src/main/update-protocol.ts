import { createHash, createPublicKey, verify } from 'node:crypto';

export type UpdateTrust = { feeds: string[]; assetOrigins: string[]; publicKeys: Record<string, string> };
export type UpdateAsset = { platform: 'darwin' | 'win32'; arch: 'arm64' | 'x64'; kind: 'mac-zip' | 'windows-zip'; urls: string[]; size: number; sha256: string };
export type UpdateManifest = { schema: 1; product: 'charfloat'; channel: 'stable'; releaseId: number; version: string; publishedAt: string; expiresAt: string; notes: string; assets: UpdateAsset[] };
export type UpdateContext = { version: string; platform: string; arch: string };
export type SeenRelease = { releaseId: number; version: string; payloadSha256: string };
export class UpdateSecurityError extends Error {}
const security = (message: string): never => { throw new UpdateSecurityError(message); };

export function compareVersion(a: string, b: string) {
  const parts = (s: string) => {
    if (!/^(0|[1-9]\d{0,5})\.(0|[1-9]\d{0,5})\.(0|[1-9]\d{0,5})$/.test(s)) return security('更新版本格式无效。');
    return s.split('.').map(Number);
  };
  const x = parts(a), y = parts(b);
  for (let n = 0; n < 3; n++) if (x[n] !== y[n]) return x[n] > y[n] ? 1 : -1;
  return 0;
}
export function httpsUrl(value: string, origins?: readonly string[]) {
  let url: URL;
  try { url = new URL(value); } catch { return security('更新地址无效。'); }
  if (url.protocol !== 'https:' || url.username || url.password || url.hash || (origins && !origins.includes(url.origin))) return security('更新地址不在已批准的 HTTPS 发布源内。');
  return url;
}
export function configuredTrust(trust: UpdateTrust) {
  if (!trust.feeds.length || !trust.assetOrigins.length || !Object.keys(trust.publicKeys).length) return false;
  if (trust.feeds.length > 4 || trust.assetOrigins.length > 8) return security('更新发布源配置超出限制。');
  for (const source of trust.feeds) httpsUrl(source);
  for (const source of trust.assetOrigins) if (httpsUrl(source).origin !== source) return security('下载源必须填写 HTTPS origin。');
  for (const [id, pem] of Object.entries(trust.publicKeys)) {
    if (!/^[a-zA-Z0-9_-]{1,64}$/.test(id) || createPublicKey(pem).asymmetricKeyType !== 'ed25519') return security('更新公钥配置无效。');
  }
  return true;
}
function base64(value: unknown, limit: number) {
  if (typeof value !== 'string' || value.length > Math.ceil(limit / 3) * 4 || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(value)) return security('更新签名编码无效。');
  const bytes = Buffer.from(value, 'base64');
  if (bytes.length > limit || bytes.toString('base64') !== value) return security('更新签名编码无效。');
  return bytes;
}
export function readSignedManifest(envelopeText: string, trust: UpdateTrust, context: UpdateContext, seen?: SeenRelease, now = Date.now()) {
  if (Buffer.byteLength(envelopeText) > 256 * 1024) return security('更新清单过大。');
  let envelope: any, raw: any;
  try { envelope = JSON.parse(envelopeText); } catch { return security('更新清单不是有效 JSON。'); }
  if (!envelope || typeof envelope.keyId !== 'string' || !Object.hasOwn(trust.publicKeys, envelope.keyId)) return security('更新清单使用了未信任的签名公钥。');
  const payload = base64(envelope.payload, 128 * 1024), signature = base64(envelope.signature, 64);
  if (signature.length !== 64 || !verify(null, payload, createPublicKey(trust.publicKeys[envelope.keyId]), signature)) return security('更新清单签名无效，已停止更新。');
  try { raw = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(payload)); } catch { return security('已签名更新清单的格式无效。'); }
  const payloadSha256 = createHash('sha256').update(payload).digest('hex');
  if (raw?.schema !== 1 || raw.product !== 'charfloat' || raw.channel !== 'stable' || !Number.isSafeInteger(raw.releaseId) || raw.releaseId < 1 || typeof raw.version !== 'string') return security('更新清单产品、频道或编号不匹配。');
  if (compareVersion(raw.version, context.version) < 0) return security('更新清单版本低于当前版本，已拒绝降级。');
  const published = Date.parse(raw.publishedAt), expires = Date.parse(raw.expiresAt);
  if (!Number.isFinite(published) || !Number.isFinite(expires) || published > now + 300000 || expires <= now || expires <= published || expires - published > 31 * 86400000) return security('更新清单已过期或发布时间无效。');
  if (seen && (raw.releaseId < seen.releaseId || compareVersion(raw.version, seen.version) < 0 || (raw.releaseId === seen.releaseId && payloadSha256 !== seen.payloadSha256))) return security('更新清单发生回退或同编号内容冲突，已停止更新。');
  if (!Array.isArray(raw.assets) || !raw.assets.length || raw.assets.length > 8 || typeof raw.notes !== 'string' || raw.notes.length > 4000) return security('更新包清单无效。');
  const identities = new Set<string>();
  const assets: UpdateAsset[] = raw.assets.map((asset: any) => {
    if (!asset || !['darwin', 'win32'].includes(asset.platform) || !['arm64', 'x64'].includes(asset.arch) || asset.kind !== (asset.platform === 'darwin' ? 'mac-zip' : 'windows-zip') || !Number.isSafeInteger(asset.size) || asset.size <= 0 || asset.size > 2 * 1024 ** 3 || !/^[a-f0-9]{64}$/.test(asset.sha256)) return security('更新包的平台、格式、大小或哈希无效。');
    const id = `${asset.platform}/${asset.arch}`; if (identities.has(id)) return security('更新包平台条目重复。'); identities.add(id);
    if (!Array.isArray(asset.urls) || !asset.urls.length || asset.urls.length > 4) return security('更新包下载源无效。');
    const urls = asset.urls.map((url: unknown) => { if (typeof url !== 'string') return security('更新包地址无效。'); return httpsUrl(url, trust.assetOrigins).href; });
    return { platform: asset.platform, arch: asset.arch, kind: asset.kind, size: asset.size, sha256: asset.sha256, urls };
  });
  const asset = assets.find(a => a.platform === context.platform && a.arch === context.arch);
  if (!asset) return security('此版本没有适用于当前系统和架构的更新包。');
  const manifest: UpdateManifest = { schema: 1, product: 'charfloat', channel: 'stable', releaseId: raw.releaseId, version: raw.version, publishedAt: raw.publishedAt, expiresAt: raw.expiresAt, notes: raw.notes, assets };
  return { manifest, asset, payloadSha256 };
}
