import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createReadStream } from 'node:fs';
import { createHash } from 'node:crypto';
import { pipeline } from 'node:stream/promises';
import asar from '@electron/asar';
import yauzl from 'yauzl';
import { checkFiles, walk, MANIFEST, PROFILE } from './release-protection.mjs';

async function hashFile(file) {
  const hash = createHash('sha256');
  for await (const chunk of createReadStream(file)) hash.update(chunk);
  return hash.digest('hex');
}

export function checkAsar(file, { allowPlainCli = false } = {}) {
  const names = asar.listPackage(file).map(n => n.replaceAll('\\', '/').replace(/^\//, ''))
    .filter(n => !asar.statFile(file, n).files);
  const errors = checkFiles(names, n => asar.extractFile(file, n), { allowPlainCli });
  const unpacked = file + '.unpacked';
  const physical = fs.existsSync(unpacked) ? walk(unpacked) : [];
  for (const name of physical) if (!names.includes(name)) errors.push(`未登记的解包文件: ${name}`);
  for (const name of ['dist/bridge/cli.cjs', 'dist/bridge/cli.jsc', 'node_modules/bytenode/lib/index.js']) {
    if (!physical.includes(name)) errors.push(`安装路径缺少运行文件: ${name}`);
  }
  for (const name of names) {
    if (asar.statFile(file, name).unpacked && !physical.includes(name)) errors.push(`解包文件缺失: ${name}`);
  }
  if (errors.length) throw new Error(errors.join('\n'));
  return { profile: PROFILE, files: names.length,
    protectedFiles: Object.keys(JSON.parse(asar.extractFile(file, MANIFEST)).files).length,
    runtimeVerified: false, hostVerified: false };
}

// 只提取审计所需的 app.asar 与 unpacked；不执行包内代码。
// 验证 ZIP 路径、重复项和尺寸；失败时关闭读取并清理临时副本。
async function extractApp(zipFile, destination) {
  const zip = await new Promise((resolve, reject) => yauzl.open(zipFile, { lazyEntries: true }, (err, z) => err ? reject(err) : resolve(z)));
  const seen = new Set();
  let found = null, bytes = 0;
  await new Promise((resolve, reject) => {
    const fail = error => { zip.close(); reject(error); };
    zip.on('error', fail);
    zip.on('end', resolve);
    zip.on('entry', entry => {
      (async () => {
        const name = entry.fileName;
        if (name.includes('\\') || name.startsWith('/') || name.split('/').includes('..') || seen.has(name)) throw new Error(`不安全或重复的 ZIP 路径: ${name}`);
        seen.add(name);
        if (/\.map$/.test(name)) throw new Error(`ZIP 内包含 sourcemap: ${name}`);
        const marker = name.match(/^(.*\/)?app\.asar(\.unpacked\/.*)?$/);
        if (!marker || name.endsWith('/')) { zip.readEntry(); return; }
        const base = marker[1] || '';
        if (found !== null && found !== base) throw new Error('ZIP 包含多个 app.asar 根目录');
        found = base;
        bytes += entry.uncompressedSize;
        if (bytes > 512 * 1024 * 1024) throw new Error('审计资源超过 512 MiB 上限');
        const relative = name.slice(base.length);
        const dest = path.join(destination, relative);
        fs.mkdirSync(path.dirname(dest), { recursive: true });
        const stream = await new Promise((res, rej) => zip.openReadStream(entry, (err, s) => err ? rej(err) : res(s)));
        await pipeline(stream, fs.createWriteStream(dest, { flags: 'wx' }));
        zip.readEntry();
      })().catch(fail);
    });
    zip.readEntry();
  });
  if (found === null || !fs.existsSync(path.join(destination, 'app.asar'))) throw new Error('ZIP 缺少 app.asar');
}

export async function checkRelease(target, options = {}) {
  let file = path.resolve(target);
  if (file.endsWith('.app')) file = path.join(file, 'Contents/Resources/app.asar');
  if (!fs.existsSync(file)) throw new Error(`发布包不存在: ${file}`);
  if (!file.endsWith('.zip')) return { artifact: path.basename(file), sha256: await hashFile(file), ...checkAsar(file, options) };
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'oab-release-check-'));
  try {
    await extractApp(file, tmp);
    return { artifact: path.basename(file), sha256: await hashFile(file), ...checkAsar(path.join(tmp, 'app.asar'), options) };
  } finally { fs.rmSync(tmp, { recursive: true, force: true }); }
}
