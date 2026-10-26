// Standalone Node-mode helper. Bundled into one CJS file and copied outside the app before launch.
import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { createRequire } from 'node:module';
import { readInstallTicket, waitForParentExit, writeInstallState, consumeInstallAuthorization, updateFailure } from './update-install.js';
import { stageMacUpdate, commitMacUpdate, recoverMacUpdate } from './update-mac.js';
import { UPDATE_TRUST } from './update-trust.js';
import { inspectUpdateZip, validateWindowsApplication } from '../../windows-bootstrap/update-zip.mjs';
import { install } from '../../windows-bootstrap/core.mjs';
import { createShortcut } from '../../windows-bootstrap/shortcut.mjs';
import { prepareWindowsUpdate } from '../../windows-bootstrap/update.mjs';

const windows = process.platform === 'win32' ? createRequire(process.execPath)('electron') as typeof import('electron') : undefined;
// The GUI closes its pipe when it exits. A later logging error must never abort
// the updater's failure receipt or old-version recovery.
process.stderr.on('error', () => {});
let verified: Awaited<ReturnType<typeof readInstallTicket>> | undefined;
async function launch(target: string) {
  const env = { ...process.env }; delete env.ELECTRON_RUN_AS_NODE; delete env.WPS_BRIDGE_RESOURCES;
  // This is the visible GUI (including recovery), not the background helper.
  // windowsHide suppresses its first ShowWindow on Windows.
  const child = spawn(windows ? path.join(target, '字浮 CharFloat.exe') : '/usr/bin/open', windows ? [] : [target], { env, cwd: path.dirname(target), stdio: 'ignore', detached: Boolean(windows) });
  await new Promise<void>((resolve, reject) => {
    child.once('error', reject);
    if (windows) child.once('spawn', () => resolve());
    else child.once('exit', code => code === 0 ? resolve() : reject(new Error('系统未能重新打开应用，请从 Finder 打开；旧包备份已保留。')));
  });
  child.unref();
}

async function main() {
  if (process.argv[2] === '--recover' && process.argv.length === 4 && process.platform === 'darwin') {
    await recoverMacUpdate(process.argv[3]); return;
  }
  const index = process.argv.indexOf('--charfloat-update-helper');
  const file = windows ? process.argv[index + 1] : process.argv[2];
  if (windows && (index < 0 || !file || process.argv[index + 2])) throw new Error('更新助手启动参数无效。');
  verified = await readInstallTicket(file, UPDATE_TRUST, process.platform, process.arch);
  const { ticket, selected, directory } = verified;
  if (windows) {
    windows.app.setPath('userData', path.join(directory, 'helper-profile'));
    await windows.app.whenReady();
    const localAppData = process.env.LOCALAPPDATA || '';
    const manifest = await inspectUpdateZip(ticket.archive, selected.manifest.version, selected.asset.sha256);
    const prepared = await prepareWindowsUpdate(file, manifest, localAppData);
    const result = await install({ localAppData, payload: ticket.archive, manifest,
      shortcut: args => createShortcut({ ...args, desktop: windows.app.getPath('desktop'), shell: windows.shell }),
      update: { ...prepared.update,
        verifyApplication: target => validateWindowsApplication(target, selected.manifest.version, ticket.arch, manifest),
      },
    });
    await writeInstallState(directory, { phase: 'installed', nonce: ticket.nonce, version: selected.manifest.version, message: '应用文件已更新，旧版本和用户配置已保留。' });
    try { await launch(result.directory); }
    catch (error) {
      await createShortcut({ desktop: windows.app.getPath('desktop'), executable: path.join(ticket.target, '字浮 CharFloat.exe'), directory: ticket.target, root: path.dirname(ticket.target), shell: windows.shell });
      throw error;
    }
    return;
  }
  if (ticket.platform !== 'darwin') throw new Error('当前系统不支持此更新。');
  const staged = await stageMacUpdate(ticket.archive, ticket.target, ticket.currentVersion, selected.manifest.version, ticket.arch);
  await writeInstallState(directory, { phase: 'ready', nonce: ticket.nonce, journal: staged.journal });
  try {
    await waitForParentExit(ticket.parentPid);
    // Reverify the signed ticket/package after waiting; no network or untrusted URL is used here.
    await readInstallTicket(file, UPDATE_TRUST, process.platform, process.arch);
    const guard = await consumeInstallAuthorization(file, ticket.nonce);
    await commitMacUpdate(staged.journal, staged.tx, undefined, guard);
  } catch (e) {
    if (staged.tx.phase === 'prepared') await fs.promises.rm(staged.directory, { recursive: true, force: true });
    throw e;
  }
  await writeInstallState(directory, { phase: 'installed', nonce: ticket.nonce, version: selected.manifest.version, journal: staged.journal, message: '应用文件已更新，旧包备份已保留。' });
  // LaunchServices keeps Gatekeeper in the normal launch path. Never strip quarantine or change policy.
  await launch(ticket.target);
}
main().catch(async error => {
  // Before verification no argv-supplied directory is trusted for writing a
  // receipt. Report through our parent's pipe instead of touching that path.
  const failure = updateFailure(error);
  console.error(JSON.stringify({ charfloatUpdateFailure: true, ...failure }));
  const file = process.argv[2];
  if (file === '--recover') { console.error(error?.message || '恢复未完成，已保留应用与备份。'); process.exitCode = 1; return; }
  // Retain validated identity even when authorization is later revoked or the feed expires.
  // A failure must remain visible after the old GUI exits.
  if (verified) {
    const { ticket, directory } = verified;
    const message = failure.message;
    try { await writeInstallState(directory, { phase: 'failed', nonce: ticket.nonce, ...failure }); } catch {}
    let parentExited = false;
    try { await waitForParentExit(ticket.parentPid, 0); parentExited = true; } catch {}
    if (parentExited) {
      try { await launch(ticket.target); }
      catch {
        if (windows) await windows.dialog.showMessageBox({ type: 'error', title: '字浮更新未完成', message, detail: '自动重新打开失败，请从保留的旧版本启动字浮，在版本更新中重试。' });
        else spawn('/usr/bin/osascript', ['-e', 'on run argv\ndisplay alert "字浮更新未完成" message (item 1 of argv)\nend run', message], { stdio: 'ignore' }).unref();
      }
    }
  }
  process.exitCode = 1;
}).finally(() => { if (windows) windows.app.exit(Number(process.exitCode) || 0); });
