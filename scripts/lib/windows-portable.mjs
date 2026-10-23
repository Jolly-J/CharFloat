// 仅适配锁定的 electron-builder 25.1.8；不修改 node_modules 或用户 TEMP。
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';

export const PORTABLE_TEMPLATE_SHA256 = '80fa75cf8cb68f4999eb92afc9f37f8e5b605cb1c3a077821bbc350a9b907a48';
export function assertPortableDefines(defines) {
  assert.equal(defines.REQUEST_EXECUTION_LEVEL, 'user', '外壳不得请求管理员权限');
  assert.ok(!('UNPACK_DIR_NAME' in defines), '必须使用每次启动独立的临时目录');
  assert.ok(!('SPLASH_IMAGE' in defines), '外壳不得显示独立开屏');
}
export function preparePortableScript(script, template, version) {
  assert.equal(version, '25.1.8', 'builder 升级后须重新核查 portable 初始化路径');
  assert.equal(createHash('sha256').update(template).digest('hex'), PORTABLE_TEMPLATE_SHA256, 'portable 模板已变化');
  assert.ok(script.endsWith(template), '实际编译脚本未包含已核查的 portable 模板');
  assert.equal((template.match(/^Function \.onInit$/gm) || []).length, 1);
  const patched = template.replace('Function .onInit\n', 'Function .onInit\n  ; Silent mode skips .onGUIInit: initialize unique temp before extraction.\n  InitPluginsDir\n')
    .replace('Function .onGUIInit\n  InitPluginsDir\n', 'Function .onGUIInit\n');
  assert.notEqual(patched, template);
  assert.equal((patched.match(/^  InitPluginsDir$/gm) || []).length, 1);
  return script.slice(0, -template.length) + patched;
}
export function installPortableCompileHook(Target, { template, version, onCompiled = () => {} }) {
  preparePortableScript(template, template, version);
  const original = Target.prototype.executeMakensis;
  assert.equal(typeof original, 'function', 'NSIS 编译接口已变化');
  async function compile(defines, commands, script) {
    assert.equal(this.isPortable, true);
    assertPortableDefines(defines);
    const prepared = preparePortableScript(script, template, version);
    const result = await original.call(this, defines, commands, prepared);
    onCompiled(prepared);
    return result;
  }
  Target.prototype.executeMakensis = compile;
  return () => { assert.equal(Target.prototype.executeMakensis, compile); Target.prototype.executeMakensis = original; };
}
