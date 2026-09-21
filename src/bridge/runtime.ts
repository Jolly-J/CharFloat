import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import crypto from 'node:crypto';
import { execSync } from 'node:child_process';

export const VERSION = '2.1.0';
export const PROTOCOL = 2;
export function runtimeHome() {
  return process.env.WPS_BRIDGE_HOME || (process.platform === 'win32'
    ? path.join(process.env.LOCALAPPDATA || os.homedir(), 'WPSBridge')
    : path.join(os.homedir(), '.wps-bridge'));
}
export function runtimePort() {
  const port = Number(process.env.WPS_BRIDGE_PORT || 19890);
  if (!Number.isInteger(port) || port < 1024 || port > 65535) throw new Error('WPS_BRIDGE_PORT 必须为 1024–65535');
  return port;
}
export function resourcePath(relative: string) {
  const entryDir = path.dirname(path.resolve(process.argv[1] || '.'));
  const roots = [process.env.WPS_BRIDGE_RESOURCES, path.resolve(entryDir, '../..'), process.cwd()].filter(Boolean) as string[];
  const found = roots.flatMap(root => { const p = path.join(root, relative); return [p.replace(/app\.asar([\\/])/, 'app.asar.unpacked$1'), p]; }).find(p => fs.existsSync(p));
  if (!found) throw new Error(`缺少运行资源 ${relative}。请重新构建或修复 Bridge 安装。`);
  return found;
}
export function atomicWrite(file: string, data: string, backup = false) {
  fs.mkdirSync(path.dirname(file), { recursive: true, mode: 0o700 });
  if (fs.existsSync(file) && fs.readFileSync(file, 'utf8') === data) return;
  if (backup && fs.existsSync(file)) fs.copyFileSync(file, `${file}.backup-${Date.now()}-${crypto.randomBytes(3).toString('hex')}`);
  const tmp = `${file}.${crypto.randomBytes(6).toString('hex')}.tmp`;
  try {
    fs.writeFileSync(tmp, data, { mode: 0o600, flag: 'wx' });
    fs.renameSync(tmp, file);
  } finally { if (fs.existsSync(tmp)) fs.unlinkSync(tmp); }
}
export function getToken(create = false) {
  const file = path.join(runtimeHome(), 'token');
  if (create) {
    fs.mkdirSync(runtimeHome(), { recursive: true, mode: 0o700 });
    try { fs.writeFileSync(file, crypto.randomBytes(32).toString('hex'), { flag: 'wx', mode: 0o600 }); }
    catch (error: any) { if (error.code !== 'EEXIST') throw error; }
  }
  return fs.readFileSync(file, 'utf8').trim();
}
export function validToken(candidate?: string) {
  if (!candidate) return false;
  const expected = Buffer.from(getToken());
  const actual = Buffer.from(candidate);
  return actual.length === expected.length && crypto.timingSafeEqual(actual, expected);
}
export function previewPath(extension: string) {
  const dir = path.join(runtimeHome(), 'previews');
  fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
  return path.join(dir, `${crypto.randomUUID()}.${extension}`);
}

export function runtimeHttpsPort() {
  const port = Number(process.env.WPS_BRIDGE_HTTPS_PORT || 19891);
  return port;
}

export function getOrGenerateCerts() {
  const certDir = path.join(runtimeHome(), 'certs');
  fs.mkdirSync(certDir, { recursive: true, mode: 0o700 });
  const keyPath = path.join(certDir, 'localhost.key');
  const certPath = path.join(certDir, 'localhost.crt');

  if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
    return {
      key: fs.readFileSync(keyPath, 'utf8'),
      cert: fs.readFileSync(certPath, 'utf8')
    };
  }

  // 1. 优先从预置资源提取 (解决 Windows 缺少 OpenSSL 命令行导致 HTTPS 无法启动的问题)
  try {
    const prebuiltKey = resourcePath('resources/certs/localhost.key');
    const prebuiltCert = resourcePath('resources/certs/localhost.crt');
    if (fs.existsSync(prebuiltKey) && fs.existsSync(prebuiltCert)) {
      fs.copyFileSync(prebuiltKey, keyPath);
      fs.copyFileSync(prebuiltCert, certPath);
      return {
        key: fs.readFileSync(keyPath, 'utf8'),
        cert: fs.readFileSync(certPath, 'utf8')
      };
    }
  } catch {}

  // 2. 备用尝试调用 openssl
  try {
    execSync(`openssl req -x509 -newkey rsa:2048 -keyout "${keyPath}" -out "${certPath}" -days 3650 -nodes -subj "/CN=localhost" -addext "subjectAltName=DNS:localhost,IP:127.0.0.1"`, { stdio: 'ignore' });
    if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
      return {
        key: fs.readFileSync(keyPath, 'utf8'),
        cert: fs.readFileSync(certPath, 'utf8')
      };
    }
  } catch {}

  return null;
}

export function appendServiceLog(tag: string, message: string) {
  try {
    const home = runtimeHome();
    fs.mkdirSync(home, { recursive: true, mode: 0o700 });
    const logPath = path.join(home, 'service.log');
    const time = new Date().toISOString();
    fs.appendFileSync(logPath, `[${time}] [${tag}] ${message}\n`, 'utf8');
  } catch {}
}

