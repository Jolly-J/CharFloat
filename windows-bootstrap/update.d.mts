import type { ZipManifest } from './update-zip.mjs';
export function prepareWindowsUpdate(file: string, manifest: ZipManifest, localAppData: string): Promise<{
  update: { target: string; currentVersion: string; claimAuthorization: () => Promise<() => Promise<void>> };
}>;
