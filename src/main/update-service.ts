import fs from 'node:fs';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import { Readable, Transform } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { configuredTrust, httpsUrl, readSignedManifest, compareVersion, UpdateSecurityError, type UpdateTrust, type UpdateContext, type SeenRelease } from './update-protocol.js';
import { lastInstallReport } from './update-install.js';

export type UpdateState = { phase: 'unconfigured' | 'idle' | 'checking' | 'up-to-date' | 'available' | 'downloading' | 'ready' | 'installing' | 'cancelled' | 'failed'; message: string; version?: string; availableVersion?: string; notes?: string; bytes?: number; totalBytes?: number; installation?: { phase: string; message: string } };
type Selected = ReturnType<typeof readSignedManifest> & { envelope: string };
type Options = { trust: UpdateTrust; context: UpdateContext; cacheDirectory: string; fetch?: typeof fetch; now?: () => number; onState?: (state: UpdateState) => void };
class UpdateSourceError extends Error {}

export class UpdateService {
  private state: UpdateState;
  private selected?: Selected;
  private artifact?: { file: string; selected: Selected };
  private active?: AbortController;
  private pending?: Promise<UpdateState>;
  private closed = false;
  private request: typeof fetch;
  private configured = false;
  private installation?: { phase: string; message: string };
  constructor(private options: Options) {
    this.options = { ...options, trust: { feeds: [...options.trust.feeds], assetOrigins: [...options.trust.assetOrigins], publicKeys: { ...options.trust.publicKeys } } };
    this.request = options.fetch || fetch;
    try {
      this.configured = configuredTrust(this.options.trust);
      this.state = this.configured ? { phase: 'idle', message: '尚未检查更新。' } : { phase: 'unconfigured', message: '官方更新源尚未配置完成，请等待后续正式版本。' };
    } catch { this.state = { phase: 'failed', message: '官方更新配置无效，已停止更新，请联系维护者。' }; }
  }
  getState() {
    const available = this.selected && compareVersion(this.selected.manifest.version, this.options.context.version) > 0;
    return { ...this.state, ...(available ? { availableVersion: this.selected!.manifest.version, notes: this.selected!.manifest.notes } : {}), ...(this.installation ? { installation: this.installation } : {}) };
  }
  get isWorking() { return Boolean(this.active) || this.state.phase === 'installing'; }
  preparingInstall() { return this.publish({ ...this.state, phase: 'installing', message: '正在准备更新并关闭字浮后台，请稍候…' }); }
  failedInstall(message: string) { return this.publish({ ...this.state, phase: this.artifact ? 'ready' : 'failed', message }); }
  async restoreInstallReport() { this.installation = await lastInstallReport(this.options.cacheDirectory); return this.getState(); }
  private publish(state: UpdateState) { this.state = state; try { this.options.onState?.(this.getState()); } catch { /* renderer loss cannot change verification */ } return this.getState(); }
  private async cache() {
    const dir = this.options.cacheDirectory;
    await fs.promises.mkdir(dir, { recursive: true, mode: 0o700 });
    const stat = await fs.promises.lstat(dir);
    if (!stat.isDirectory() || stat.isSymbolicLink() || (process.platform !== 'win32' && ((stat.mode & 0o022) !== 0 || stat.uid !== process.getuid?.()))) throw new UpdateSecurityError('更新缓存目录不安全，未写入任何安装文件。');
    return dir;
  }
  private async seen(): Promise<SeenRelease | undefined> {
    const file = path.join(await this.cache(), 'seen.json');
    try {
      const stat = await fs.promises.lstat(file);
      if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 4096) throw new Error();
      const value = JSON.parse(await fs.promises.readFile(file, 'utf8'));
      if (!Number.isSafeInteger(value.releaseId) || value.releaseId < 1 || !/^[a-f0-9]{64}$/.test(value.payloadSha256)) throw new Error();
      compareVersion(value.version, this.options.context.version); return value;
    } catch (error: any) { if (error.code === 'ENOENT') return undefined; throw new UpdateSecurityError('更新版本记录损坏，已停止更新，请保留缓存并检查日志。'); }
  }
  private async remember(selected: Selected) {
    const dir = await this.cache(), target = path.join(dir, 'seen.json'), temporary = path.join(dir, `.seen-${randomUUID()}.tmp`);
    try {
      await fs.promises.writeFile(temporary, JSON.stringify({ releaseId: selected.manifest.releaseId, version: selected.manifest.version, payloadSha256: selected.payloadSha256 }), { flag: 'wx', mode: 0o600 });
      await fs.promises.rename(temporary, target);
    } finally { await fs.promises.unlink(temporary).catch(() => {}); }
  }
  private async response(url: string, origins: string[], signal: AbortSignal) {
    let current = httpsUrl(url, origins);
    for (let redirects = 0; redirects < 4; redirects++) {
      const response = await this.request(current.href, { signal, redirect: 'manual', credentials: 'omit', cache: 'no-store', referrerPolicy: 'no-referrer' });
      if ([301, 302, 303, 307, 308].includes(response.status)) {
        const location = response.headers.get('location'); await response.body?.cancel();
        if (!location) throw new UpdateSecurityError('更新重定向缺少目标。');
        current = httpsUrl(new URL(location, current).href, origins); continue;
      }
      if (!response.ok) { await response.body?.cancel(); throw new UpdateSourceError(`官方更新源返回 HTTP ${response.status}，请稍后检查或重试。`); }
      return response;
    }
    throw new UpdateSecurityError('更新重定向次数超出限制。');
  }
  private async text(response: Response, signal: AbortSignal) {
    if (!response.body) throw new Error('更新清单为空。');
    const reader = response.body.getReader(); const parts: Uint8Array[] = []; let bytes = 0;
    const abort = () => { void reader.cancel().catch(() => {}); };
    signal.addEventListener('abort', abort, { once: true });
    try {
      while (true) {
        signal.throwIfAborted(); const chunk = await reader.read(); signal.throwIfAborted(); if (chunk.done) break;
        bytes += chunk.value.length; if (bytes > 256 * 1024) throw new UpdateSecurityError('更新清单过大。'); parts.push(chunk.value);
      }
      return new TextDecoder('utf-8', { fatal: true }).decode(Buffer.concat(parts));
    } finally { signal.removeEventListener('abort', abort); await reader.cancel().catch(() => {}); }
  }
  private run(action: (signal: AbortSignal) => Promise<UpdateState>) {
    if (this.closed) return Promise.reject(new Error('字浮正在退出，已停止更新任务。'));
    if (this.active) return Promise.reject(new Error('更新任务正在进行，请先取消或等待完成。'));
    const controller = new AbortController(); this.active = controller;
    const pending = action(controller.signal).catch(error => this.publish({ phase: controller.signal.aborted ? 'cancelled' : 'failed', message: controller.signal.aborted ? '更新已取消，现有应用和配置未改变。' : error instanceof UpdateSecurityError || error instanceof UpdateSourceError ? error.message : '更新未完成，请检查网络或磁盘空间后重试。', ...(this.selected ? { version: this.selected.manifest.version, totalBytes: this.selected.asset.size } : {}) })).finally(() => { if (this.active === controller) { this.active = undefined; this.pending = undefined; } });
    this.pending = pending; return pending;
  }
  check() {
    if (!this.configured) return Promise.resolve(this.getState());
    return this.run(async signal => {
      // Keep the last verified discovery visible while checking, or if a later request fails.
      this.publish({ phase: 'checking', message: '正在检查官方更新…' });
      const seen = await this.seen(); const origins = this.options.trust.feeds.map(url => httpsUrl(url).origin); let selected: Selected | undefined;
      let sourceFailure: unknown;
      for (const source of this.options.trust.feeds) {
        try {
          const deadline = AbortSignal.any([signal, AbortSignal.timeout(30000)]);
          const envelope = await this.text(await this.response(source, origins, deadline), deadline);
          selected = { ...readSignedManifest(envelope, this.options.trust, this.options.context, seen, this.options.now?.() ?? Date.now()), envelope }; break;
        } catch (error) { if (signal.aborted || error instanceof UpdateSecurityError) throw error; sourceFailure = error; }
      }
      if (!selected) throw sourceFailure instanceof UpdateSourceError ? sourceFailure : new UpdateSourceError('暂时无法连接官方更新源，请检查网络后重试。');
      signal.throwIfAborted(); await this.remember(selected); signal.throwIfAborted(); this.selected = selected;
      if (this.artifact?.selected.asset.sha256 === selected.asset.sha256) this.artifact.selected = selected;
      return this.publish(compareVersion(selected.manifest.version, this.options.context.version) === 0
        ? { phase: 'up-to-date', message: '当前已是此更新源提供的最新版本。', version: selected.manifest.version }
        : this.artifact?.selected.asset.sha256 === selected.asset.sha256
          ? { phase: 'ready', message: '更新包已验证，可点击“重启并更新”。', version: selected.manifest.version, totalBytes: selected.asset.size }
          : { phase: 'available', message: `发现新版本 ${selected.manifest.version}，可在后台下载。`, version: selected.manifest.version, notes: selected.manifest.notes, totalBytes: selected.asset.size });
    });
  }
  download() {
    if (!this.selected || !['available', 'failed', 'cancelled', 'ready'].includes(this.state.phase)) return Promise.reject(new Error('请先检查并确认可用更新。'));
    const selected = this.selected;
    return this.run(async signal => {
      readSignedManifest(selected.envelope, this.options.trust, this.options.context, await this.seen(), this.options.now?.() ?? Date.now());
      for (const url of selected.asset.urls) {
        signal.throwIfAborted(); const staging = await fs.promises.mkdtemp(path.join(await this.cache(), '.pending-')); const partial = path.join(staging, 'package.part');
        try {
          let bytes = 0, last = 0; const hash = createHash('sha256');
          this.publish({ phase: 'downloading', message: '正在下载并校验更新包…', version: selected.manifest.version, bytes: 0, totalBytes: selected.asset.size });
          const deadline = AbortSignal.any([signal, AbortSignal.timeout(10 * 60 * 1000)]);
          const response = await this.response(url, this.options.trust.assetOrigins, deadline);
          const length = response.headers.get('content-length');
          if (length !== null && (!/^\d+$/.test(length) || Number(length) !== selected.asset.size)) { await response.body?.cancel(); throw new UpdateSecurityError('更新包大小与签名清单不符。'); }
          if (!response.body) throw new Error('更新包为空。');
          await pipeline(Readable.fromWeb(response.body as any), new Transform({ transform: (chunk, _encoding, done) => {
            bytes += chunk.length; if (bytes > selected.asset.size) return done(new UpdateSecurityError('更新包超过签名清单大小。'));
            hash.update(chunk);
            if (Date.now() - last >= 200) { last = Date.now(); this.publish({ phase: 'downloading', message: '正在下载并校验更新包…', version: selected.manifest.version, bytes, totalBytes: selected.asset.size }); }
            done(null, chunk);
          } }), fs.createWriteStream(partial, { flags: 'wx', mode: 0o600 }), { signal: deadline });
          signal.throwIfAborted();
          if (bytes !== selected.asset.size || hash.digest('hex') !== selected.asset.sha256) throw new UpdateSecurityError('更新包校验失败，已丢弃下载；原应用未改变。');
          const file = path.join(this.options.cacheDirectory, `${selected.manifest.releaseId}-${randomUUID()}.zip`);
          const handle = await fs.promises.open(partial, 'r+'); try { await handle.sync(); } finally { await handle.close(); }
          await fs.promises.rename(partial, file); this.artifact = { file, selected };
          return this.publish({ phase: 'ready', message: '更新包已验证。保存办公文档后，点击“重启并更新”。', version: selected.manifest.version, bytes, totalBytes: selected.asset.size });
        } catch (error) { if (signal.aborted || error instanceof UpdateSecurityError || url === selected.asset.urls.at(-1)) throw error; }
        finally { await fs.promises.rm(staging, { recursive: true, force: true }); }
      }
      throw new Error('更新下载未完成。');
    });
  }
  cancel() { this.active?.abort(); }
  async shutdown() { this.closed = true; this.cancel(); await this.pending; }
  resumeAfterFailedHandoff() { if (this.active) throw new Error('更新任务仍在进行。'); this.closed = false; }
  async verifiedDownload() {
    if (this.active || this.state.phase !== 'ready' || !this.artifact) throw new Error('没有可安装的已验证更新包。');
    const { file, selected } = this.artifact;
    readSignedManifest(selected.envelope, this.options.trust, this.options.context, await this.seen(), this.options.now?.() ?? Date.now());
    const handle = await fs.promises.open(file, fs.constants.O_RDONLY | (fs.constants.O_NOFOLLOW || 0));
    try {
      const before = await handle.stat(); if (!before.isFile() || before.nlink !== 1 || before.size !== selected.asset.size) throw new UpdateSecurityError('更新包被替换，已停止安装。');
      const hash = createHash('sha256'); for await (const chunk of handle.createReadStream({ autoClose: false })) hash.update(chunk);
      const after = await handle.stat();
      if (hash.digest('hex') !== selected.asset.sha256 || before.mtimeMs !== after.mtimeMs || before.size !== after.size) throw new UpdateSecurityError('更新包已改变，已停止安装，请重新下载。');
      return { file, kind: selected.asset.kind, version: selected.manifest.version, envelope: selected.envelope };
    } finally { await handle.close(); }
  }
}
