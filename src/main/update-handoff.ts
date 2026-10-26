import fs from 'node:fs';
import path from 'node:path';
import { spawn, type ChildProcess } from 'node:child_process';
import { createInstallTicket, plainFile, grantInstallAuthorization, revokeInstallAuthorization, updateFailure } from './update-install.js';

export async function handoffUpdate(options: { cache: string; artifact: { file: string; envelope: string }; currentVersion: string; target: string; helper: string; platform: 'darwin' | 'win32'; arch: 'arm64' | 'x64'; timeoutMs?: number; spawnProcess?: typeof spawn }) {
  const job = await createInstallTicket(options.cache, { parentPid: process.pid, platform: options.platform, arch: options.arch, currentVersion: options.currentVersion, target: options.target, archive: options.artifact.file, envelope: options.artifact.envelope });
  const env = { ...process.env };
  for (const key of Object.keys(env)) if (/^(ELECTRON_RUN_AS_NODE|WPS_BRIDGE_RESOURCES|PORTABLE_EXECUTABLE.*|CHARFLOAT_DIAGNOSTIC_.*)$/i.test(key)) delete env[key];
  // Windows runs the bundled helper in a separate Electron process, never the downloaded ZIP/installer.
  let executable = process.execPath, args = ['--charfloat-update-helper', job.file];
  let child: ChildProcess | undefined, exited = false, exitCode: number | null = null, signal: string | null = null, stderr = '';
  try {
  if (options.platform === 'darwin') {
    const copy = path.join(job.directory, 'helper.cjs');
    await fs.promises.copyFile(options.helper, copy, fs.constants.COPYFILE_EXCL);
    executable = process.execPath; args = [copy, job.file]; env.ELECTRON_RUN_AS_NODE = '1';
  }
  child = (options.spawnProcess || spawn)(executable, args, { env, stdio: ['ignore', 'ignore', 'pipe'], detached: true, windowsHide: true });
  child.stderr?.on('data', bytes => { stderr = (stderr + bytes).slice(-8192); });
  let launchError: Error | undefined;
  child.on('error', error => { launchError = error; }); child.on('exit', (code, stoppedBy) => { exited = true; exitCode = code; signal = stoppedBy; });
  const deadline = Date.now() + (options.timeoutMs ?? 120000);
    while (Date.now() < deadline) {
      let ready = false;
      try {
        const file = path.join(job.directory, 'result.json'); await plainFile(file, 16384);
        const result = JSON.parse(await fs.promises.readFile(file, 'utf8'));
        if (result.nonce !== job.ticket.nonce) throw new Error('更新交接回执不匹配。');
        if (result.phase === 'failed') throw Object.assign(new Error(result.message || '更新准备失败。'), { code: result.code });
        ready = result.phase === 'ready';
      } catch (error: any) { if (error.code !== 'ENOENT') throw error; }
      if (launchError) throw launchError;
      if (exited) {
        let message = `更新助手已退出（退出码 ${exitCode ?? signal ?? '未知'}），当前应用未替换。`, code = 'UPDATE_HELPER_EXIT';
        for (const line of stderr.trim().split(/\r?\n/).reverse()) {
          try { const value = JSON.parse(line); if (value.charfloatUpdateFailure === true && typeof value.message === 'string') { const failure = updateFailure(value); message = failure.message; code = failure.code || code; break; } } catch {}
        }
        throw Object.assign(new Error(message), { code, exitCode, signal, stderr });
      }
      if (ready) { await grantInstallAuthorization(job.file, job.ticket.nonce); child.unref(); return job; }
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    throw new Error('更新安装器未能就绪；请检查磁盘空间或系统安全提示，当前应用未替换。');
  } catch (error) {
    await revokeInstallAuthorization(job.file, job.ticket.nonce, updateFailure(error));
    // The old parent is still alive, so helpers cannot have entered their commit phase.
    // Terminate only this directly spawned waiting helper, never processes found by name.
    if (!exited && child?.pid) child.kill('SIGTERM');
    throw error;
  }
}
