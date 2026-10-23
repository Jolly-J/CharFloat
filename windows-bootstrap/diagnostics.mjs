// 显式离线诊断 opt-in；只记录白名单状态，不记录文稿、环境字典、堆栈或凭据。
import fs from './physical-fs.mjs';
import path from 'node:path';

const keys = new Set(['stage','phase','activity','file','bytes','verifiedBytes','totalBytes','fileBytes','fileTotalBytes','enabled','message','code','reason','exitCode','action','idleMs','rendererSilentMs','electron','node','arch','platform']);
export function createDiagnostics({ env = process.env, now = Date.now, maximumBytes = 2 * 1024 * 1024 } = {}) {
  const disabled = { enabled: false, record() {}, flush() {}, close() {} };
  const runId = env.CHARFLOAT_DIAGNOSTIC_RUN_ID, dir = env.CHARFLOAT_DIAGNOSTIC_DIR;
  function unavailable(reason) { if (runId || dir) console.error(`CharFloat diagnostics unavailable: ${reason}.`); return { ...disabled, reason }; }
  if (!/^[a-f0-9]{8}-(?:[a-f0-9]{4}-){3}[a-f0-9]{12}$/i.test(runId || '') || !env.LOCALAPPDATA || !dir) return unavailable('invalid-context');
  let fd, physicalDir;
  try {
    let current = path.resolve(dir);
    for (let depth = 0; ; depth++) {
      if (depth > 64) throw new Error('diagnostic path depth');
      const stat = fs.lstatSync(current); if (!stat.isDirectory() || stat.isSymbolicLink()) throw new Error('diagnostic path link');
      const parent = path.dirname(current); if (parent === current) break; current = parent;
    }
    // 先解析同一物理目录，再比较；不能让 Win32 大小写/8.3 表示差异关闭日志。
    const realpath = fs.realpathSync.native || fs.realpathSync;
    physicalDir = realpath(dir);
    const expected = realpath(path.join(env.LOCALAPPDATA, 'CharFloatStartupDiagnostics', runId));
    const normalize = value => process.platform === 'win32' ? value.toLowerCase() : value;
    if (normalize(physicalDir) !== normalize(expected)) return unavailable('directory-mismatch');
    fd = fs.openSync(path.join(physicalDir, 'bootstrap.jsonl'), 'wx', 0o600);
  } catch { return unavailable('unsafe-or-unavailable-directory'); }
  const start = now(); let sequence = 0, size = 0, closed = false, last = {}, failed = false, capped = false;
  const prefixes = [env.LOCALAPPDATA, env.APPDATA, env.USERPROFILE, env.TEMP, env.TMP].filter(Boolean).sort((a,b) => b.length-a.length);
  function text(value) {
    let result = String(value).slice(0, 2048);
    for (const prefix of prefixes) result = result.replace(new RegExp(prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), '<user-path>');
    return result.replace(/\b(Bearer\s+)[a-z0-9._~+/-]+/gi, '$1<redacted>').replace(/\b(token|password|cookie|authorization|api[_-]?key)\s*[:=]\s*[^\s,;]+/gi, '$1=<redacted>').slice(0, 512);
  }
  function entry(event, fields) {
    const value = { runId, sequence: ++sequence, at: new Date(now()).toISOString(), elapsedMs: now()-start, event };
    for (const [key,item] of Object.entries(fields || {})) if (keys.has(key)) {
      if (typeof item === 'string') value[key] = text(item);
      else if (typeof item === 'boolean' || typeof item === 'number' && Number.isFinite(item)) value[key] = item;
    }
    return value;
  }
  function warn() { if (!failed) { failed = true; console.error('CharFloat diagnostic log could not be written.'); } }
  return {
    enabled: true,
    record(event, fields = {}) {
      if (closed || !/^[a-z][a-z0-9-]{0,48}$/.test(event)) return;
      const value = entry(event, fields); last = value;
      const line = JSON.stringify(value)+'\n', length = Buffer.byteLength(line);
      if (size + length > maximumBytes) { capped = true; return; }
      try { fs.writeSync(fd, line); size += length; } catch { warn(); }
    },
    flush() { if (!closed) try { fs.fsyncSync(fd); } catch { warn(); } },
    close(fields = {}) {
      if (closed) return;
      const value = { ...last, ...entry('process-exit', fields), logBytes: size, logCapped: capped, writeFailure: failed };
      try { fs.writeFileSync(path.join(physicalDir, 'bootstrap-final.json'), JSON.stringify(value), { flag: 'wx', mode: 0o600 }); fs.fsyncSync(fd); } catch { warn(); }
      try { fs.closeSync(fd); } catch {} closed = true;
    }
  };
}
