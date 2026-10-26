import fs from 'node:fs';
import path from 'node:path';
import { canonicalDirectory, plainFile } from '../../src/main/update-install.js';

/** Resolve spelling aliases only after rejecting links; never changes the product validator. */
export async function checkedProbeArchive(cache: string, archive: string) {
  await canonicalDirectory(cache);
  await canonicalDirectory(path.dirname(archive));
  await plainFile(archive);
  const cachePhysical = await fs.promises.realpath(cache), archivePhysical = await fs.promises.realpath(archive);
  const relative = path.relative(cachePhysical, archivePhysical);
  if (!relative || relative === '..' || relative.startsWith('..' + path.sep) || path.isAbsolute(relative)) throw new Error('更新 ZIP 不在本轮缓存内，未读取。');
  // Preserve the GUI's original short/long spelling when replaying its helper.
  const replayCache = path.relative(cachePhysical, path.dirname(archivePhysical)) === '' ? path.dirname(archive) : cache;
  return { cachePhysical, archivePhysical, replayCache, relative, cacheSpellingChanged: cache !== cachePhysical, archiveSpellingChanged: archive !== archivePhysical };
}
