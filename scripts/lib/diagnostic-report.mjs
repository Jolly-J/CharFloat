import fs from 'node:fs';

// The caller supplies an already redacted, bounded report. Always keep the JSON
// in stdout as well: a report-directory failure must not erase the evidence.
export function writeDiagnosticReport(file, report) {
  const body = JSON.stringify(report, null, 2) + '\n';
  process.stdout.write(`CHARFLOAT_DIAGNOSTIC_REPORT_BEGIN\n${body}CHARFLOAT_DIAGNOSTIC_REPORT_END\n`);
  try { fs.writeFileSync(file, body, { flag: 'wx' }); return true; }
  catch (error) {
    process.stderr.write(JSON.stringify({ reportWriteFailed: true, code: error.code || 'UNKNOWN', reportPreservedInConsole: true }) + '\n');
    return false;
  }
}
