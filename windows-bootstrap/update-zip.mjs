// Online updates share the user-install transaction, but never enter the branded installer UI.
import fs from './physical-fs.mjs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import asar from '@electron/asar';
import { openZipStream } from './zip-stream.mjs';
import { openZip, safeName, regular, fileHash, validateManifest, PRODUCT, EXE } from './core.mjs';

/** The whole ZIP must already be authorized by the signed manifest; derive a bounded file inventory. */
export async function inspectUpdateZip(archive, version, expectedHash) {
  await regular(archive);
  if (await fileHash(archive) !== expectedHash) throw new Error('更新 ZIP 已改变，未准备安装。');
  const zip = await openZip(archive), files = [], names = new Map();
  let totalBytes = 0, count = 0;
  await new Promise((resolve, reject) => {
    let finished = false;
    const fail = error => { if (!finished) { finished = true; zip.close(); reject(error); } };
    zip.on('error', fail);
    zip.on('end', () => { if (!finished) { finished = true; resolve(); } });
    zip.on('entry', entry => { void (async () => {
      if (++count > 20000 || (entry.generalPurposeBitFlag & 1)) throw new Error('更新 ZIP 加密或条目过多。');
      const name = safeName(entry.fileName), key = name.normalize('NFC').toLowerCase();
      const directory = entry.fileName.endsWith('/'), type = (entry.externalFileAttributes >>> 16) & 0xf000;
      if (names.has(key) || (type && type !== (directory ? 0x4000 : 0x8000))) throw new Error('更新 ZIP 含重复路径、链接或特殊文件。');
      const parts = key.split('/');
      for (let i = 1; i < parts.length; i++) if (names.get(parts.slice(0, i).join('/')) === false) throw new Error('更新 ZIP 文件与目录冲突。');
      if (!directory && [...names.keys()].some(n => n.startsWith(key + '/'))) throw new Error('更新 ZIP 文件与目录冲突。');
      names.set(key, directory);
      if (directory) {
        if (entry.uncompressedSize) throw new Error('更新 ZIP 目录条目无效。');
      } else {
        totalBytes += entry.uncompressedSize;
        if (!Number.isSafeInteger(totalBytes) || totalBytes > 4 * 1024 ** 3) throw new Error('更新 ZIP 展开大小超出限制。');
        const input = await openZipStream(zip, entry);
        const hash = createHash('sha256'); let size = 0;
        for await (const chunk of input) {
          size += chunk.length;
          if (size > entry.uncompressedSize) throw new Error('更新 ZIP 文件大小超出声明。');
          hash.update(chunk);
        }
        if (size !== entry.uncompressedSize) throw new Error('更新 ZIP 文件不完整。');
        files.push({ name, size, sha256: hash.digest('hex') });
      }
      zip.readEntry();
    })().catch(fail); });
    zip.readEntry();
  });
  return validateManifest({ product: PRODUCT, version, payloadSha256: expectedHash, files, totalBytes });
}

export async function validateWindowsApplication(directory, version, arch, manifest) {
  if (arch !== 'x64') throw new Error('Windows 更新仅支持 x64。');
  const expected = new Set(manifest.files.map(f => f.name));
  async function walk(dir, prefix = '') {
    for (const name of await fs.promises.readdir(dir)) {
      const relative = prefix + name, file = path.join(dir, name), st = await fs.promises.lstat(file);
      if (st.isSymbolicLink()) throw new Error('更新应用含有链接，未切换启动入口。');
      if (st.isDirectory()) await walk(file, relative + '/');
      else if (!st.isFile() || st.nlink !== 1 || (!expected.has(relative) && relative !== '.charfloat-install.json')) throw new Error('更新应用含未知文件，未切换启动入口。');
    }
  }
  await walk(directory);
  const exe = path.join(directory, EXE); await regular(exe);
  const handle = await fs.promises.open(exe, 'r');
  try {
    const dos = Buffer.alloc(64);
    if ((await handle.read(dos, 0, dos.length, 0)).bytesRead !== 64 || dos.toString('ascii', 0, 2) !== 'MZ') throw new Error('更新程序不是 Windows PE。');
    const offset = dos.readUInt32LE(60), pe = Buffer.alloc(24);
    if (offset < 64 || offset > 1024 * 1024 || (await handle.read(pe, 0, pe.length, offset)).bytesRead !== 24 || pe.readUInt32LE(0) !== 0x4550 || pe.readUInt16LE(4) !== 0x8664) throw new Error('更新主程序平台或架构不匹配。');
  } finally { await handle.close(); }
  const archive = path.join(directory, 'resources/app.asar'), st = await regular(archive);
  // Bound the ASAR header before the library allocates/decodes it.
  const input = await fs.promises.open(archive, 'r');
  try {
    const header = Buffer.alloc(8);
    if ((await input.read(header, 0, 8, 0)).bytesRead !== 8 || header.readUInt32LE(0) !== 4 || header.readUInt32LE(4) > 16 * 1024 * 1024 || header.readUInt32LE(4) + 8 > st.size) throw new Error('更新应用 ASAR 标识无效。');
  } finally { await input.close(); }
  asar.uncache(archive);
  try {
    const entry = asar.statFile(archive, 'package.json', false);
    if (entry.link || entry.unpacked || !Number.isSafeInteger(entry.size) || entry.size < 1 || entry.size > 1024 * 1024) throw new Error('更新应用包标识无效。');
    const pkg = JSON.parse(asar.extractFile(archive, 'package.json', false).toString('utf8'));
    if (pkg.name !== 'charfloat' || pkg.version !== version || pkg.main !== 'dist/main/index.js') throw new Error('更新应用身份或版本与签名清单不符。');
  } finally { asar.uncache(archive); }
}
