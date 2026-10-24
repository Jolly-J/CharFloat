import type { ZipManifest } from './update-zip.mjs';
export function fileHash(file: string): Promise<string>;
export function extractPayload(file: string, destination: string, manifest: ZipManifest): Promise<void>;
export function validateUpdateTarget(localAppData: string, target: string, currentVersion: string): Promise<string>;
export function install(options: {
  localAppData: string; payload: string; manifest: ZipManifest;
  shortcut: (args: { executable: string; root: string; directory: string }) => Promise<void>;
  update: { target: string; currentVersion: string; verifyApplication: (directory: string) => Promise<void>; claimAuthorization: () => Promise<() => Promise<void>> };
}): Promise<{ executable: string; directory: string; payloadSha256: string }>;
