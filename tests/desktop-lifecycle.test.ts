import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn, type ChildProcess } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import net from 'node:net';
import { DesktopService } from '../src/main/desktop-service.js';
import { stopOwnedChild } from '../src/bridge/owned-process.js';
import { runProcess, stopNativeJobs } from '../src/bridge/process-runner.js';
import { pauseAutomaticServiceStart, ensureService, serviceRequest } from '../src/bridge/service-client.js';

const sleep = (ms: number) => new Promise(r => setTimeout(r, ms));
async function fixture(ignore = false) {
  const child = spawn(process.execPath, ['-e', `const http=require('http'); const s=http.createServer((q,r)=>r.end('fixture')); s.listen(0,'127.0.0.1',()=>process.send({port:s.address().port})); process.on('message',m=>{if(m.type==='charfloat:shutdown'&&!${ignore})s.close(()=>process.exit(0));}); process.on('SIGTERM',()=>{});`], { stdio: ['ignore', 'ignore', 'ignore', 'ipc'] });
  const ready = await new Promise<any>((resolve, reject) => {
    const cleanup = () => { clearTimeout(timer); child.removeListener('message', message); child.removeListener('exit', exit); child.removeListener('error', error); };
    const message = (value: any) => { cleanup(); resolve(value); };
    const exit = (code: number | null) => { cleanup(); reject(new Error(`Temporary listener exited before readiness (${code}); check local listen permissions.`)); };
    const error = (cause: Error) => { cleanup(); reject(cause); };
    const timer = setTimeout(() => { cleanup(); child.kill('SIGKILL'); reject(new Error('Temporary listener readiness timed out.')); }, 5000);
    child.once('message', message); child.once('exit', exit); child.once('error', error);
  });
  return { child, port: ready.port };
}
async function alive(port: number) { try { return (await fetch(`http://127.0.0.1:${port}/health`, { signal: AbortSignal.timeout(1000) })).ok; } catch { return false; } }

test('desktop quit releases only its owned listener, is idempotent and blocks restart', async () => {
  const foreign = await fixture(); let own: Awaited<ReturnType<typeof fixture>> | undefined;
  const desktop = new DesktopService(async (_entry, options) => { own = await fixture(); options?.onSpawn?.(own.child); });
  try {
    await desktop.start('fixture'); assert.equal(await alive(own!.port), true);
    await Promise.all([desktop.stop(), desktop.stop()]);
    assert.equal(await alive(own!.port), false); assert.equal(await alive(foreign.port), true);
    await assert.rejects(desktop.start('fixture'), /正在退出/);
  } finally { if (own?.child.exitCode === null) own.child.kill('SIGKILL'); foreign.child.kill('SIGKILL'); }
});

test('attaching a pre-existing service does not give desktop ownership', async () => {
  const foreign = await fixture(); const desktop = new DesktopService(async () => {});
  try { await desktop.start('fixture'); await desktop.stop(); assert.equal(await alive(foreign.port), true); }
  finally { foreign.child.kill('SIGKILL'); }
});

test('unresponsive daemon fails shutdown rather than claiming success or killing its unknown tree', async () => {
  const own = await fixture(true); const desktop = new DesktopService(async (_entry, o) => { o?.onSpawn?.(own.child); }, 60);
  try { await desktop.start('fixture'); await assert.rejects(desktop.stop(), /未确认完成退出/); assert.equal(await alive(own.port), true); }
  finally { own.child.kill('SIGKILL'); }
});

test('exit during an in-flight start cancels before spawn', async () => {
  let release!: () => void, spawned = 0;
  const pending = new Promise<void>(r => { release = r; });
  const desktop = new DesktopService(async (_entry, o) => { await pending; if (o?.signal?.aborted) throw new Error('cancelled'); spawned++; });
  const start = desktop.start('fixture'); const rejection = assert.rejects(start, /cancelled/);
  const stopping = desktop.stop(); release(); await stopping; await rejection; assert.equal(spawned, 0);
});

test('owned direct helper can be force-stopped after bounded graceful wait', async () => {
  const own = await fixture(true);
  await stopOwnedChild(own.child, { graceMs: 40, forceMs: 40 });
  assert.equal(await alive(own.port), false);
});

test('real CLI owned by desktop shuts down and frees its isolated port', async () => {
  const home = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-owned-cli-'));
  const listener = net.createServer(); await new Promise<void>(r => listener.listen(0, '127.0.0.1', r));
  const port = (listener.address() as net.AddressInfo).port; await new Promise<void>(r => listener.close(() => r()));
  const saved = { home: process.env.WPS_BRIDGE_HOME, port: process.env.WPS_BRIDGE_PORT, resources: process.env.WPS_BRIDGE_RESOURCES };
  process.env.WPS_BRIDGE_HOME = home; process.env.WPS_BRIDGE_PORT = String(port); process.env.WPS_BRIDGE_RESOURCES = process.cwd();
  const desktop = new DesktopService();
  try {
    await desktop.start(path.resolve('dist/bridge/cli.cjs'));
    const health: any = await (await fetch(`http://127.0.0.1:${port}/health`)).json(); assert.ok(health.pid > 0);
    await pauseAutomaticServiceStart(); await desktop.stop(); assert.equal(await alive(port), false);
    await assert.rejects(ensureService(path.resolve('dist/bridge/cli.cjs')), /已退出并暂停/);
    const resumed = new DesktopService();
    try { await resumed.start(path.resolve('dist/bridge/cli.cjs')); assert.equal(fs.existsSync(path.join(home, 'desktop-stopped.json')), false); }
    finally { await resumed.stop(); }
  } finally {
    await desktop.stop();
    for (const [key, value] of Object.entries({ WPS_BRIDGE_HOME: saved.home, WPS_BRIDGE_PORT: saved.port, WPS_BRIDGE_RESOURCES: saved.resources })) {
      if (value === undefined) delete process.env[key]; else process.env[key] = value;
    }
    fs.rmSync(home, { recursive: true, force: true });
  }
});
test('review R3: daemon spawned before GUI exit cannot commit a late automatic start after pause', async () => {
  const home = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-late-start-')), ready = path.join(home, 'waiting'), gate = path.join(home, 'continue');
  const listener = net.createServer(); await new Promise<void>(r => listener.listen(0, '127.0.0.1', r));
  const port = (listener.address() as net.AddressInfo).port; await new Promise<void>(r => listener.close(() => r()));
  const saved = { home: process.env.WPS_BRIDGE_HOME, port: process.env.WPS_BRIDGE_PORT, resources: process.env.WPS_BRIDGE_RESOURCES };
  process.env.WPS_BRIDGE_HOME = home; process.env.WPS_BRIDGE_PORT = String(port); process.env.WPS_BRIDGE_RESOURCES = process.cwd();
  const wrapper = path.join(home, 'delayed.cjs');
  fs.writeFileSync(wrapper, `const fs=require('fs');fs.writeFileSync(${JSON.stringify(ready)},'waiting');const gate=${JSON.stringify(gate)};const deadline=Date.now()+5000;while(!fs.existsSync(gate)&&Date.now()<deadline)Atomics.wait(new Int32Array(new SharedArrayBuffer(4)),0,0,10);require(${JSON.stringify(path.resolve('dist/bridge/cli.cjs'))});`);
  let child: ChildProcess | undefined;
  const starting = ensureService(wrapper, { onSpawn: value => { child = value; } });
  // Observe rejection immediately so a fast fenced child cannot trigger an unhandled rejection.
  const result = starting.then(() => ({ success: true, message: '' }), e => ({ success: false, message: e.message }));
  try {
    for (let n = 0; n < 200 && !fs.existsSync(ready); n++) await sleep(10);
    assert.equal(fs.existsSync(ready), true);
    await pauseAutomaticServiceStart(); // GUI exit is now committed in another process's start window.
    fs.writeFileSync(gate, 'continue');
    const outcome = await result;
    assert.equal(outcome.success, false, 'late daemon must not resurrect after GUI exit');
    assert.equal(await alive(port), false);
    const explicit = spawn(process.execPath, [path.resolve('dist/bridge/cli.cjs'), '--start'], { env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }, stdio: 'ignore' });
    const code = await new Promise<number | null>((resolve, reject) => {
      const timer = setTimeout(() => { explicit.kill('SIGKILL'); reject(new Error('Explicit fixture start timed out')); }, 10000);
      explicit.once('exit', code => { clearTimeout(timer); resolve(code); }); explicit.once('error', error => { clearTimeout(timer); reject(error); });
    });
    assert.equal(code, 0); assert.equal(await alive(port), true, 'fresh explicit --start must resume');
    await serviceRequest('/api/v1/service/stop', {});
    for (let n = 0; n < 100 && await alive(port); n++) await sleep(10);
    assert.equal(await alive(port), false);
  } finally {
    if (await alive(port)) { await serviceRequest('/api/v1/service/stop', {}).catch(() => {}); await sleep(150); }
    if (child?.exitCode === null) { child.kill('SIGKILL'); await new Promise<void>(r => child!.once('exit', () => r())); }
    for (const [key, value] of Object.entries({ WPS_BRIDGE_HOME: saved.home, WPS_BRIDGE_PORT: saved.port, WPS_BRIDGE_RESOURCES: saved.resources })) { if (value === undefined) delete process.env[key]; else process.env[key] = value; }
    fs.rmSync(home, { recursive: true, force: true });
  }
});

test('daemon shutdown stops only recorded native helper and never treats interruption as success', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-native-stop-')); const ready = path.join(dir, 'ready');
  const unrelated = await fixture();
  const job = runProcess(process.execPath, ['-e', `require('fs').writeFileSync(process.argv[1],'ready');setInterval(()=>{},1000);`, ready]);
  const rejected = assert.rejects(job, /中断|退出/);
  try {
    for (let n = 0; n < 100 && !fs.existsSync(ready); n++) await sleep(20);
    assert.ok(fs.existsSync(ready)); await stopNativeJobs(); await rejected;
    assert.equal(await alive(unrelated.port), true);
    await assert.rejects(runProcess(process.execPath, ['-e', '']), /正在退出/);
  } finally { unrelated.child.kill('SIGKILL'); fs.rmSync(dir, { recursive: true, force: true }); }
});
