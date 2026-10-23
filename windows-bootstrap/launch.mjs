import { spawn } from 'node:child_process';
import { verifyInstallation } from './core.mjs';

export async function launchApplication({ executable, directory, manifest, spawnProcess = spawn, environment = process.env }) {
  await verifyInstallation(directory, manifest);
  const env = { ...environment };
  for (const key of Object.keys(env)) if (key === 'ELECTRON_RUN_AS_NODE' || key.startsWith('PORTABLE_EXECUTABLE')) delete env[key];
  const child = spawnProcess(executable, [], { cwd: directory, detached: true, stdio: 'ignore', env });
  await new Promise((resolve, reject) => { child.once('spawn', resolve); child.once('error', reject); });
  child.unref();
}
