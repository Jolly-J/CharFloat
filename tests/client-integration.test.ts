import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import * as TOML from '@iarna/toml';
import { parse as parseYaml } from 'yaml';
import { applyMcpEntry, inspectMcpConfig, inspectMcpEnablement, mergeCodexToml, mergeMcpConfig } from '../src/main/mcp-adapters.js';
import { readDshMcp } from '../src/main/dsh-mcp.js';
import { CLIENT_DEFINITIONS, clientPathContext } from '../src/main/client-registry.js';
import { syncManagedSkills, linkClientSkills, inspectSkillLinks } from '../src/main/skill-links.js';
import { retargetUpdateConfig } from '../src/main/update-config.js';
import { runDesktopDiagnostics } from '../src/main/desktop-diagnostics.js';
import { VERSION } from '../src/bridge/runtime.js';

function fixture(t: any) {
  const dir = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), 'charfloat-client-')));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const cli = path.join(dir, 'cli.cjs'); fs.writeFileSync(cli, '// fixture');
  return { dir, entry: { command: process.execPath, args: [cli], env: { ELECTRON_RUN_AS_NODE: '1', WPS_BRIDGE_HOME: dir } } };
}

test('Mac and Windows client paths use their actual roots, not host separator substitution', () => {
  const mac = clientPathContext({ home: '/Users/fixture', platform: 'darwin' });
  const win = clientPathContext({ home: 'C:\\Users\\测试', platform: 'win32', env: { APPDATA: 'D:\\Roaming' } });
  const paths = (id: string, ctx: typeof mac) => CLIENT_DEFINITIONS.find(d => d.id === id)!.paths(ctx);
  for (const [id, macPath, winPath] of [
    ['claude-desktop', '/Users/fixture/Library/Application Support/Claude/claude_desktop_config.json', 'D:\\Roaming\\Claude\\claude_desktop_config.json'],
    ['vscode', '/Users/fixture/Library/Application Support/Code/User/mcp.json', 'D:\\Roaming\\Code\\User\\mcp.json'],
    ['codex', '/Users/fixture/.codex/config.toml', 'C:\\Users\\测试\\.codex\\config.toml'],
    ['deepseek-harness', '/Users/fixture/.dsh/cordis.patch.yml', 'C:\\Users\\测试\\.dsh\\cordis.patch.yml'],
    ['opencode', '/Users/fixture/.config/opencode/opencode.json', 'C:\\Users\\测试\\.config\\opencode\\opencode.json'],
    ['windsurf', '/Users/fixture/.config/devin/mcp_config.json', 'D:\\Roaming\\devin\\mcp_config.json'],
  ]) { assert.equal(paths(id, mac).configPath, macPath, id); assert.equal(paths(id, win).configPath, winPath, id); }
  const custom = clientPathContext({ home: 'C:\\Users\\测试', platform: 'win32', env: { APPDATA: 'D:\\Roaming', DSH_HOME: 'E:\\Data\\DSH', CODEX_HOME: 'E:\\Data\\Codex', KIMI_CODE_HOME: 'E:\\Data\\Kimi', XDG_CONFIG_HOME: 'E:\\Config' } });
  assert.equal(paths('deepseek-harness', custom).configPath, 'E:\\Data\\DSH\\cordis.patch.yml');
  assert.equal(paths('codex', custom).configPath, 'E:\\Data\\Codex\\config.toml');
  assert.equal(paths('kimi', custom).skillsPath, 'E:\\Data\\Kimi\\skills');
  assert.equal(paths('opencode', custom).configPath, 'E:\\Config\\opencode\\opencode.json');
  assert.equal(paths('codex', mac).skillsPath, '/Users/fixture/.agents/skills');
  assert.equal(CLIENT_DEFINITIONS.find(d => d.id === 'qwen-work')!.configKind, 'manual');
  assert.equal(CLIENT_DEFINITIONS.find(d => d.id === 'qwen')!.name, 'Qwen Code');
});

test('JSON/JSONC readback distinguishes a named entry from a valid, current, enabled command', t => {
  const f = fixture(t), file = path.join(f.dir, 'mcp.json');
  assert.equal(inspectMcpConfig('json-mcp', file, f.entry).status, 'missing');
  fs.writeFileSync(file, '{"mcpServers":{"charfloat":{}}}');
  assert.equal(inspectMcpConfig('json-mcp', file, f.entry).status, 'invalid');
  applyMcpEntry('json-mcp', file, f.entry);
  assert.equal(inspectMcpConfig('json-mcp', file, f.entry).status, 'valid');
  let data = JSON.parse(fs.readFileSync(file, 'utf8')); data.mcpServers.charfloat.command = '/missing/old-app'; fs.writeFileSync(file, JSON.stringify(data));
  assert.equal(inspectMcpConfig('json-mcp', file, f.entry).status, 'outdated');
  data.mcpServers.charfloat.disabled = true; data.mcpServers.charfloat.custom = 'keep'; fs.writeFileSync(file, JSON.stringify(data));
  applyMcpEntry('json-mcp', file, f.entry);
  assert.equal(inspectMcpConfig('json-mcp', file, f.entry).status, 'disabled');
  assert.equal(JSON.parse(fs.readFileSync(file, 'utf8')).mcpServers.charfloat.custom, 'keep');
  const vscode = path.join(f.dir, 'vscode.json');
  fs.writeFileSync(vscode, '{\n// keep user comment\n"servers": {"other":{"command":"keep"} // other comment\n}, "inputs": []\n}');
  applyMcpEntry('json-vscode', vscode, f.entry); const text = fs.readFileSync(vscode, 'utf8');
  assert.match(text, /keep user comment/); assert.match(text, /other comment/);
  assert.equal(inspectMcpConfig('json-vscode', vscode, f.entry).status, 'valid');
  assert.doesNotMatch(text, /mcpServers/);
});

test('VS Code preserves both official formats and rejects relative custom homes', t => {
  const f = fixture(t), ctx = clientPathContext({ home: f.dir, platform: 'darwin' });
  const native = path.join(ctx.appData, 'Code/User/mcp.json'), portable = path.join(f.dir, '.copilot/mcp-config.json');
  for (const file of [native, portable]) { fs.mkdirSync(path.dirname(file), { recursive: true }); fs.writeFileSync(file, '{}'); }
  const targets = CLIENT_DEFINITIONS.find(d => d.id === 'vscode')!.paths(ctx).configTargets!;
  assert.deepEqual(targets, [{ file: native, kind: 'json-vscode' }, { file: portable, kind: 'json-mcp' }]);
  for (const target of targets) applyMcpEntry(target.kind, target.file, f.entry);
  assert.ok(JSON.parse(fs.readFileSync(native, 'utf8')).servers.charfloat);
  assert.ok(JSON.parse(fs.readFileSync(portable, 'utf8')).mcpServers.charfloat);
  for (const platform of ['darwin', 'win32'] as const) {
    const bad = clientPathContext({ home: platform === 'darwin' ? f.dir : 'C:\\Users\\fixture', platform, env: { DSH_HOME: 'relative' } });
    assert.throws(() => CLIENT_DEFINITIONS.find(d => d.id === 'deepseek-harness')!.paths(bad), /不是绝对路径/);
  }
});

test('Gemini persistent disable state is inspected without rewriting or enabling it', t => {
  const f = fixture(t), file = path.join(f.dir, 'enablement.json'), valid = { status: 'valid', detail: 'saved' } as const;
  fs.writeFileSync(file, '{"charfloat":{"enabled":false},"other":{"enabled":false}}');
  const before = fs.readFileSync(file, 'utf8');
  assert.equal(inspectMcpEnablement(file, valid).status, 'disabled');
  assert.equal(fs.readFileSync(file, 'utf8'), before);
  fs.writeFileSync(file, 'broken'); assert.equal(inspectMcpEnablement(file, valid).status, 'invalid');
});

test('malformed, duplicate-key and linked configs are preserved without exposing their values', t => {
  const f = fixture(t), file = path.join(f.dir, 'broken.json');
  for (const content of ['{"mcpServers":{},"mcpServers":{}}', '{"secret":"sensitive-test-value",broken']) {
    fs.writeFileSync(file, content);
    assert.throws(() => mergeMcpConfig(file, f.entry), error => !String(error).includes('sensitive-test-value'));
    assert.equal(fs.readFileSync(file, 'utf8'), content);
  }
  const target = path.join(f.dir, 'target.json'); fs.writeFileSync(target, '{}'); fs.unlinkSync(file); fs.symlinkSync(target, file);
  assert.throws(() => mergeMcpConfig(file, f.entry)); assert.equal(fs.readFileSync(target, 'utf8'), '{}');
});

test('Codex preserves unrelated sections, comments, custom options and disabled state', t => {
  const f = fixture(t), file = path.join(f.dir, 'config.toml');
  fs.writeFileSync(file, '# my settings\nmodel = "keep"\n[mcp_servers.other]\ncommand = "other"\n[mcp_servers.charfloat]\ncommand = "old"\nenabled = false\ntool_timeout_sec = 120\n[mcp_servers.charfloat.env]\nKEEP = "yes"\n');
  mergeCodexToml(file, f.entry); const text = fs.readFileSync(file, 'utf8'), parsed: any = TOML.parse(text);
  assert.equal(parsed.model, 'keep'); assert.equal(parsed.mcp_servers.other.command, 'other'); assert.equal(parsed.mcp_servers.charfloat.enabled, false);
  assert.equal(parsed.mcp_servers.charfloat.tool_timeout_sec, 120); assert.equal(parsed.mcp_servers.charfloat.env.KEEP, 'yes'); assert.match(text, /# my settings/);
  mergeCodexToml(file, f.entry); assert.equal(fs.readFileSync(file, 'utf8'), text);
  const multiline = 'prompt = """\n[mcp_servers.charfloat]\nexample in a string\n"""\n'; fs.writeFileSync(file, multiline);
  assert.throws(() => mergeCodexToml(file, f.entry)); assert.equal(fs.readFileSync(file, 'utf8'), multiline);
});

test('DSH merges a real insert patch, retains opaque !!js expressions, and never executes them', t => {
  const f = fixture(t), file = path.join(f.dir, 'cordis.patch.yml');
  fs.writeFileSync(file, '# keep DSH\n- insert:\n    - id: user-plugin\n      name: user-plugin\n      config:\n        value: !!js globalThis.__charfloatShouldNeverRun = true\n');
  applyMcpEntry('yaml-dsh', file, f.entry);
  assert.equal((globalThis as any).__charfloatShouldNeverRun, undefined);
  const text = fs.readFileSync(file, 'utf8'); assert.match(text, /!!js/); assert.match(text, /# keep DSH/); assert.match(text, /user-plugin/);
  assert.equal(readDshMcp(file)?.command, f.entry.command); assert.equal(inspectMcpConfig('yaml-dsh', file, f.entry).status, 'valid');
  applyMcpEntry('yaml-dsh', file, f.entry); assert.equal(fs.readFileSync(file, 'utf8'), text);
  const duplicate = text + '\n- insert:\n    - id: mcp-charfloat\n      name: other-plugin\n'; fs.writeFileSync(file, duplicate);
  assert.throws(() => applyMcpEntry('yaml-dsh', file, f.entry)); assert.equal(fs.readFileSync(file, 'utf8'), duplicate);
});

test('DSH preserves explicit disabling and refuses ambiguous dynamic or replacement patches before writing', t => {
  const f = fixture(t), file = path.join(f.dir, 'cordis.patch.yml');
  applyMcpEntry('yaml-dsh', file, f.entry); const base = fs.readFileSync(file, 'utf8');
  fs.writeFileSync(file, base + '\n- id: mcp-charfloat\n  disabled: true\n');
  applyMcpEntry('yaml-dsh', file, f.entry); assert.equal(inspectMcpConfig('yaml-dsh', file, f.entry).status, 'disabled');
  for (const override of ['disabled: !!js process.platform === "win32"', 'name: another-plugin', 'config: {command: another}']) {
    const before = base + '\n- id: mcp-charfloat\n  ' + override + '\n'; fs.writeFileSync(file, before);
    assert.throws(() => applyMcpEntry('yaml-dsh', file, f.entry)); assert.equal(fs.readFileSync(file, 'utf8'), before);
  }
});

function skillsFixture(t: any) {
  const f = fixture(t), source = path.join(f.dir, 'bundle-skills'), root = path.join(f.dir, 'managed');
  fs.mkdirSync(path.join(source, 'charfloat', 'scripts'), { recursive: true });
  fs.writeFileSync(path.join(source, 'charfloat', 'SKILL.md'), '---\nname: charfloat\n---\nfixture version one');
  const binary = Buffer.from([0, 255, 128, 4]); fs.writeFileSync(path.join(source, 'charfloat', 'scripts', 'fixture.bin'), binary);
  return { ...f, source, root, binary };
}
test('one stable skill source, verified directory links, backups, idempotency and refreshed content', t => {
  const f = skillsFixture(t), managed = syncManagedSkills(f.source, f.root), dest = path.join(f.dir, 'client', 'skills');
  fs.mkdirSync(path.join(dest, 'charfloat'), { recursive: true }); fs.writeFileSync(path.join(dest, 'charfloat', 'SKILL.md'), 'user custom copy');
  const linked = linkClientSkills('fixture-client', dest, managed); assert.equal(linked.backups, 1);
  assert.equal(fs.lstatSync(path.join(dest, 'charfloat')).isSymbolicLink(), true); assert.equal(inspectSkillLinks(dest, f.root).status, 'linked');
  assert.deepEqual(fs.readFileSync(path.join(dest, 'charfloat/scripts/fixture.bin')), f.binary);
  const backupRoot = path.join(f.dir, 'client/.charfloat-skill-backups');
  const backup = path.join(backupRoot, fs.readdirSync(backupRoot)[0], 'charfloat/SKILL.md'); assert.equal(fs.readFileSync(backup, 'utf8'), 'user custom copy');
  assert.equal(linkClientSkills('fixture-client', dest, managed).backups, 0);
  fs.writeFileSync(path.join(f.source, 'charfloat/SKILL.md'), 'updated version two'); syncManagedSkills(f.source, f.root);
  assert.equal(fs.readFileSync(path.join(dest, 'charfloat/SKILL.md'), 'utf8'), 'updated version two');
});
test('link failure leaves the existing skill intact; Windows requests a directory junction', t => {
  const f = skillsFixture(t), managed = syncManagedSkills(f.source, f.root), dest = path.join(f.dir, 'client/skills');
  fs.mkdirSync(path.join(dest, 'charfloat'), { recursive: true }); fs.writeFileSync(path.join(dest, 'charfloat/SKILL.md'), 'keep');
  const original = fs.symlinkSync; let requested: any;
  t.mock.method(fs, 'symlinkSync', (_target: any, _path: any, type: any) => { requested = type; throw new Error('simulated permission denial'); });
  assert.throws(() => linkClientSkills('fixture', dest, managed, 'win32'), /permission denial/); assert.equal(requested, 'junction');
  assert.equal(fs.readFileSync(path.join(dest, 'charfloat/SKILL.md'), 'utf8'), 'keep');
  t.mock.restoreAll(); assert.equal(fs.symlinkSync, original);
});
test('skill sync preflights custom modifications and never overwrites an unmanaged source', t => {
  const f = skillsFixture(t); syncManagedSkills(f.source, f.root);
  fs.writeFileSync(path.join(f.root, 'charfloat/scripts/fixture.bin'), 'user-edited');
  fs.writeFileSync(path.join(f.source, 'charfloat/SKILL.md'), 'new source');
  assert.throws(() => syncManagedSkills(f.source, f.root), /自定义修改/);
  assert.match(fs.readFileSync(path.join(f.root, 'charfloat/SKILL.md'), 'utf8'), /version one/);
  const unknown = path.join(f.dir, 'unknown'); fs.mkdirSync(unknown); fs.writeFileSync(path.join(unknown, 'notes'), 'keep');
  assert.throws(() => syncManagedSkills(f.source, unknown)); assert.equal(fs.readFileSync(path.join(unknown, 'notes'), 'utf8'), 'keep');
});

test('skill refresh rolls back written files when a later write fails', t => {
  const f = skillsFixture(t); syncManagedSkills(f.source, f.root);
  const first = path.join(f.root, 'charfloat/SKILL.md'), before = fs.readFileSync(first);
  fs.writeFileSync(path.join(f.source, 'charfloat/SKILL.md'), 'new version');
  fs.writeFileSync(path.join(f.source, 'charfloat/scripts/fixture.bin'), 'new binary');
  const rename = fs.renameSync;
  t.mock.method(fs, 'renameSync', (from: any, to: any) => {
    if (to === path.join(f.root, 'charfloat/scripts/fixture.bin')) throw new Error('simulated I/O failure');
    return rename(from, to);
  });
  assert.throws(() => syncManagedSkills(f.source, f.root), /simulated I\/O/);
  assert.deepEqual(fs.readFileSync(first), before);
  assert.deepEqual(fs.readFileSync(path.join(f.root, 'charfloat/scripts/fixture.bin')), f.binary);
});

test('Windows path migration also handles DSH YAML and VS Code JSONC without enabling disabled entries', t => {
  const f = fixture(t), previous = 'C:\\Programs\\CharFloat\\2.2.3-aaaaaaaaaaaa', next = 'C:\\Programs\\CharFloat\\2.2.4-bbbbbbbbbbbb';
  const entry = { command: path.win32.join(previous, '字浮 CharFloat.exe'), args: [path.win32.join(previous, 'resources/cli.cjs')], env: { WPS_BRIDGE_RESOURCES: path.win32.join(previous, 'resources') } };
  for (const kind of ['yaml-dsh', 'json-vscode'] as const) {
    const file = path.join(f.dir, kind); applyMcpEntry(kind, file, entry);
    if (kind === 'json-vscode') { const data = JSON.parse(fs.readFileSync(file, 'utf8')); data.servers.charfloat.disabled = true; fs.writeFileSync(file, JSON.stringify(data)); }
    assert.equal(retargetUpdateConfig(file, kind, previous, next), true);
    const text = fs.readFileSync(file, 'utf8'); assert.ok(text.includes('2.2.4-bbbbbbbbbbbb')); assert.equal(retargetUpdateConfig(file, kind, previous, next), false);
    if (kind === 'json-vscode') assert.equal(JSON.parse(text).servers.charfloat.disabled, true);
    else assert.equal(parseYaml(text)[0].insert[0].config.command, path.win32.join(next, '字浮 CharFloat.exe'));
  }
});

test('desktop checks still run offline and show unknown states rather than a fake healthy result', async () => {
  let statusCalls = 0;
  const result = await runDesktopDiagnostics({ probe: async () => ({ ready: false, occupied: false } as any), status: async () => { statusCalls++; return {}; }, wps: () => ({ installed: false }), office: async () => ({ installed: false }), clients: () => [] });
  assert.equal(statusCalls, 0); assert.ok(result.attention >= 1); assert.equal(result.checks.find(c => c.id === 'bridge')?.action, 'start');
  const online = await runDesktopDiagnostics({ probe: async () => ({ ready: true, info: { version: VERSION } } as any), status: async () => ({ isWpsConnected: true, components: {} }), wps: () => ({ installed: true, current: true }), office: async () => ({ installed: true, current: true }), clients: () => [] });
  assert.equal(online.checks.find(c => c.id === 'wps')?.status, 'ok'); assert.equal(online.checks.find(c => c.id === 'office')?.status, 'idle');
});
