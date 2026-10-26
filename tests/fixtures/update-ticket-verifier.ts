import fs from 'node:fs';
import { readInstallTicket } from '../../src/main/update-install.js';

try {
  const value = await readInstallTicket(process.argv[2], JSON.parse(fs.readFileSync(process.argv[3], 'utf8')), 'win32', 'x64');
  console.log(JSON.stringify({ verified: true, version: value.selected.manifest.version }));
} catch (error: any) { console.error(error.message); process.exitCode = 1; }
