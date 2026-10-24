export type ZipManifest = { product: string; version: string; payloadSha256: string; totalBytes: number; files: { name: string; size: number; sha256: string }[] };
export function inspectUpdateZip(archive: string, version: string, expectedHash: string): Promise<ZipManifest>;
export function validateWindowsApplication(directory: string, version: string, arch: string, manifest: ZipManifest): Promise<void>;
