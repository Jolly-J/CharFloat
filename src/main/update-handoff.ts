import fs from 'node:fs';
import path from 'node:path';
import { spawn, type ChildProcess } from 'node:child_process';
import { createInstallTicket, plainFile, grantInstallAuthorization, revokeInstallAuthorization } from './update-install.js';

export async function handoffUpdate(options: { cache: string; artifact: { file: string; envelope: string }; currentVersion: string; target: string; helper: string; platform: 'darwin' | 'win32'; arch: 'arm64' | 'x64'; timeoutMs?: number; spawnProcess?: typeof spawn }) {
  const job = await createInstallTicket(options.cache, { parentPid: process.pid, platform: options.platform, arch: options.arch, currentVersion: options.currentVersion, target: options.target, archive: options.artifact.file, envelope: options.artifact.envelope });
  const env = { ...process.env };
  for (const key of Object.keys(env)) if (/^(ELECTRON_RUN_AS_NODE|WPS_BRIDGE_RESOURCES|PORTABLE_EXECUTABLE.*|CHARFLOAT_DIAGNOSTIC_.*)$/i.test(key)) delete env[key];
  // Windows runs the bundled helper in a separate Electron process, never the downloaded ZIP/installer.
  let executable = process.execPath, args = ['--charfloat-update-helper', job.file];
  let child: ChildProcess | undefined, exited = false;
  try {
  if (options.platform === 'darwin') {
    const copy = path.join(job.directory, 'helper.cjs');
    await fs.promises.copyFile(options.helper, copy, fs.constants.COPYFILE_EXCL);
    executable = process.execPath; args = [copy, job.file]; env.ELECTRON_RUN_AS_NODE = '1';
  }
  child = (options.spawnProcess || spawn)(executable, args, { env, stdio: 'ignore', detached: true, windowsHide: true });
  let launchError: Error | undefined;
  child.on('error', error => { launchError = error; }); child.on('exit', () => { exited = true; });
  const deadline = Date.now() + (options.timeoutMs ?? 120000);
    while (Date.now() < deadline) {
      if (launchError || exited) throw launchError || new Error('更新安装器已退出，当前应用未替换。');
      try {
        const file = path.join(job.directory, 'result.json'); await plainFile(file, 16384);
        const result = JSON.parse(await fs.promises.readFile(file, 'utf8'));
        if (result.nonce !== job.ticket.nonce) throw new Error('更新交接回执不匹配。');
        if (result.phase === 'failed') throw new Error(result.message || '更新准备失败。');
        if (result.phase === 'ready') { await grantInstallAuthorization(job.file, job.ticket.nonce); child.unref(); return job; }
      } catch (error: any) { if (error.code !== 'ENOENT') throw error; }
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    throw new Error('更新安装器未能就绪；请检查磁盘空间或系统安全提示，当前应用未替换。');
  } catch (error) {
    await revokeInstallAuthorization(job.file, job.ticket.nonce);
    // The old parent is still alive, so helpers cannot have entered their commit phase.
    // Terminate only this directly spawned waiting helper, never processes found by name.
    if (!exited && child?.pid) child.kill('SIGTERM');
    throw error;
  }
}
