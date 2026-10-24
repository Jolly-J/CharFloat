import { readInstallTicket, waitForParentExit, writeInstallState, consumeInstallAuthorization } from '../src/main/update-install.ts';
import { UPDATE_TRUST } from '../src/main/update-trust.ts';
import { validateUpdateTarget } from './core.mjs';

export function updateTicketArgument(argv) {
  const index = argv.indexOf('--charfloat-update');
  if (index < 0) return undefined;
  if (!argv[index + 1] || argv[index + 2]) throw new Error('更新启动参数无效');
  return argv[index + 1];
}
export async function prepareWindowsUpdate(file, manifest, localAppData, options = {}) {
  const trust = options.trust || UPDATE_TRUST;
  const value = await readInstallTicket(file, trust, options.platform || process.platform, options.arch || process.arch);
  const { ticket, selected, directory } = value;
  try {
    if (ticket.platform !== 'win32' || ticket.arch !== 'x64' || selected.asset.kind !== 'windows-zip' || selected.manifest.version !== manifest.version || selected.asset.sha256 !== manifest.payloadSha256) throw new Error('更新 ZIP 版本或架构、哈希与签名清单不符');
    await validateUpdateTarget(localAppData, ticket.target, ticket.currentVersion);
    await writeInstallState(directory, { phase: 'ready', nonce: ticket.nonce });
    await (options.wait || waitForParentExit)(ticket.parentPid);
    await readInstallTicket(file, trust, options.platform || process.platform, options.arch || process.arch);
    return { ...value, update: { target: ticket.target, currentVersion: ticket.currentVersion, claimAuthorization: async () => {
      await readInstallTicket(file, trust, options.platform || process.platform, options.arch || process.arch);
      return consumeInstallAuthorization(file, ticket.nonce);
    } } };
  } catch (error) { await writeInstallState(directory, { phase: 'failed', nonce: ticket.nonce, message: error.message }); throw error; }
}
export async function reportWindowsUpdate(value, phase, message) {
  await writeInstallState(value.directory, { phase, nonce: value.ticket.nonce, message });
}
