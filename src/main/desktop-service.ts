import type { ChildProcess } from 'node:child_process';
import { ensureService } from '../bridge/service-client.js';
import { stopOwnedChild } from '../bridge/owned-process.js';

/** Lifetime belongs to this GUI only. Attaching an existing Bridge never grants ownership. */
export class DesktopService {
  private abort = new AbortController();
  private children = new Set<ChildProcess>();
  private starts = new Set<Promise<void>>();
  private stopping?: Promise<void>;
  get closing() { return this.abort.signal.aborted; }
  ownsProcess(pid: number) { return [...this.children].some(child => child.pid === pid && child.exitCode === null && child.signalCode === null); }
  beginClosing() { this.abort.abort(); }
  constructor(private ensure = ensureService, private graceMs = 7000) {}
  start(entry: string) {
    if (this.closing) return Promise.reject(new Error('字浮正在退出，不能重新启动后台。'));
    const pending = this.ensure(entry, { signal: this.abort.signal, resumeAutomaticStart: true, onSpawn: child => {
      this.children.add(child);
      child.once('exit', () => this.children.delete(child));
      child.once('error', () => { if (!child.pid) this.children.delete(child); });
    } });
    this.starts.add(pending);
    void pending.then(() => this.starts.delete(pending), () => this.starts.delete(pending));
    return pending;
  }
  stop() {
    this.beginClosing();
    if (!this.stopping) this.stopping = (async () => {
      const results = await Promise.allSettled([...this.children].map(child => stopOwnedChild(child, {
        graceful: () => { if (child.connected) child.send({ type: 'charfloat:shutdown' }, () => {}); },
        graceMs: this.graceMs, force: false,
      })));
      const failure = results.find(r => r.status === 'rejected');
      if (failure?.status === 'rejected') throw failure.reason;
      await Promise.allSettled([...this.starts]);
    })().catch(error => { this.stopping = undefined; throw error; });
    return this.stopping;
  }
  /** Only used after a failed update handoff; it never starts a service automatically. */
  resumeAfterFailedHandoff() {
    if (this.children.size || this.starts.size) throw new Error('后台仍未退出，不能恢复启动。');
    this.abort = new AbortController(); this.stopping = undefined;
  }
}
