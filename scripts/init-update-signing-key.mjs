// Explicit local initialization only. Never called by builds, uploads or runtime code.
import fs from 'node:fs';
import path from 'node:path';
import { createHash, createPrivateKey, createPublicKey, generateKeyPairSync } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { ROOT, LOCAL_SECRETS, KEY_DIRECTORY, PRIVATE_KEY, checkKeyBoundaries } from './lib/update-key-safety.mjs';

function secureDirectory(dir) {
  try { fs.mkdirSync(dir, { mode: 0o700 }); } catch (error) { if (error.code !== 'EEXIST') throw error; }
  const st = fs.lstatSync(dir);
  if (!st.isDirectory() || st.isSymbolicLink() || (process.platform !== 'win32' && ((st.mode & 0o077) || st.uid !== process.getuid?.()))) throw new Error('本地签名目录权限或类型不安全，未生成密钥。');
}
function writeNew(file, bytes) {
  const fd = fs.openSync(file, 'wx', 0o600);
  try { fs.writeFileSync(fd, bytes); fs.fsyncSync(fd); } finally { fs.closeSync(fd); }
}
function existing(file) {
  try { return fs.lstatSync(file); } catch (error) { if (error.code === 'ENOENT') return undefined; throw error; }
}
function secureRead(file) {
  const before = fs.lstatSync(file);
  if (!before.isFile() || before.isSymbolicLink()) throw new Error('本地签名文件类型异常，未读取密钥。');
  const fd = fs.openSync(file, fs.constants.O_RDONLY | (fs.constants.O_NOFOLLOW || 0));
  try {
    const st = fs.fstatSync(fd);
    if (!st.isFile() || st.nlink !== 1 || st.size > 16384 || st.dev !== before.dev || st.ino !== before.ino || (process.platform !== 'win32' && ((st.mode & 0o077) || st.uid !== process.getuid?.()))) throw new Error('本地签名文件权限或类型异常，未读取密钥。');
    return fs.readFileSync(fd);
  } finally { fs.closeSync(fd); }
}
export function initializeUpdateSigningKey(root = ROOT) {
  checkKeyBoundaries(root);
  root = fs.realpathSync(root);
  secureDirectory(path.join(root, LOCAL_SECRETS));
  const dir = path.join(root, KEY_DIRECTORY); secureDirectory(dir);
  const lock = path.join(dir, '.initialize.lock');
  const fd = fs.openSync(lock, 'wx', 0o600);
  try {
    const known = new Set([PRIVATE_KEY, 'public-key.pem', 'key-info.json', '.initialize.lock']);
    if (fs.readdirSync(dir).some(name => !known.has(name))) throw new Error('签名目录中存在未知文件，未创建或覆盖密钥。');
    const privateFile = path.join(dir, PRIVATE_KEY), publicFile = path.join(dir, 'public-key.pem'), infoFile = path.join(dir, 'key-info.json');
    const privateStat = existing(privateFile);
    if (!privateStat && (existing(publicFile) || existing(infoFile))) throw new Error('已有签名元数据但私钥缺失，未生成替代密钥。');
    let key;
    if (privateStat) key = createPrivateKey(secureRead(privateFile));
    else {
      key = generateKeyPairSync('ed25519').privateKey;
      // Exclusive creation and sync: a failure never triggers deletion, replacement or another key generation.
      const pem = key.export({ format: 'pem', type: 'pkcs8' });
      writeNew(privateFile, pem);
    }
    if (key.asymmetricKeyType !== 'ed25519') throw new Error('已有私钥不是 Ed25519，未替换。');
    const publicKey = createPublicKey(key), der = publicKey.export({ format: 'der', type: 'spki' });
    const fingerprint = createHash('sha256').update(der).digest('hex');
    const keyId = `stable-${fingerprint.slice(0, 16)}`;
    const publicPem = publicKey.export({ format: 'pem', type: 'spki' });
    if (existing(publicFile)) {
      const saved = createPublicKey(secureRead(publicFile)).export({ format: 'der', type: 'spki' });
      if (!saved.equals(der)) throw new Error('已有公钥与私钥不匹配，未覆盖。');
    } else writeNew(publicFile, publicPem);
    if (existing(infoFile)) {
      const info = JSON.parse(secureRead(infoFile).toString('utf8'));
      if (info.algorithm !== 'Ed25519' || info.keyId !== keyId || info.publicKeySha256 !== fingerprint) throw new Error('已有签名元数据不匹配，未覆盖。');
    } else writeNew(infoFile, JSON.stringify({ algorithm: 'Ed25519', keyId, publicKeySha256: fingerprint, fingerprintEncoding: 'SHA-256 of DER SPKI', createdAt: new Date().toISOString() }, null, 2) + '\n');
    return { created: !privateStat, privateKeyPath: `${KEY_DIRECTORY}/${PRIVATE_KEY}`, keyId, publicKeySha256: fingerprint };
  } finally { fs.closeSync(fd); fs.unlinkSync(lock); }
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    if (process.argv.length !== 2) throw new Error('初始化工具不接受自定义输出路径。');
    console.log(JSON.stringify(initializeUpdateSigningKey()));
  } catch { console.error('本地更新签名初始化未完成；未输出私钥、未覆盖已有密钥。请检查排除规则、目录状态与权限。'); process.exitCode = 1; }
}
