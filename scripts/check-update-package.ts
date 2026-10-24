// Exercise the actual updater against a built ZIP, without installing or launching the app.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { extractMacArchive, validateMacBundle } from '../src/main/update-mac.js';
import { inspectUpdateZip, validateWindowsApplication } from '../windows-bootstrap/update-zip.mjs';
import { extractPayload, fileHash } from '../windows-bootstrap/core.mjs';

export async function checkUpdatePackage(platform: string, version: string, file: string) {
  if (!['darwin', 'win32'].includes(platform) || !/^\d+\.\d+\.\d+$/.test(version)) throw new Error('需要平台 darwin/win32、准确版本和完整 ZIP 路径。');
  const root = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'charfloat-package-check-'));
  try {
    const directory = await fs.promises.realpath(root), archive = path.resolve(file);
    if (platform === 'darwin') {
      const application = await extractMacArchive(archive, directory);
      await validateMacBundle(application, version, 'arm64');
    } else {
      const manifest = await inspectUpdateZip(archive, version, await fileHash(archive));
      await extractPayload(archive, directory, manifest);
      await validateWindowsApplication(directory, version, 'x64', manifest);
    }
    return { platform, version, actualUpdaterExtraction: true, applicationIdentity: true, applicationLaunched: false, systemAcceptance: false };
  } finally { await fs.promises.rm(root, { recursive: true, force: true }); }
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    const [platform, version, file] = process.argv.slice(2);
    if (!file) throw new Error('用法：node --import tsx scripts/check-update-package.ts <darwin|win32> <版本> <ZIP>');
    console.log(JSON.stringify(await checkUpdatePackage(platform, version, file)));
  } catch (error: any) { console.error(error.message || '真实更新包校验未完成'); process.exitCode = 1; }
}
