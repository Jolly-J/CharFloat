/**
 * macOS 磁盘权限识别测试（本次用户体验修复的核心逻辑）。
 *
 * 覆盖的风险：把**无关失败**误判成权限问题会让用户跑去改系统设置却解决不了问题；
 * 把**权限失败**漏判则用户继续看到裸 EPERM。两种都要防。
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { detectPermissionIssueCore, extractProtectedPath } from '../src/main/permission-detect.js';

test('真实报错被识别为权限问题并提取出目标路径', () => {
  const wps = new Error("EPERM: operation not permitted, open '/Users/x/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml'");
  const r = detectPermissionIssueCore(wps, 'darwin');
  assert.ok(r, 'WPS 侧 EPERM 必须被识别');
  assert.equal(r!.targetPath, '/Users/x/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml');

  const office = new Error('cp: /Users/x/Library/Containers/com.microsoft.Excel/Data/Documents/wef/wps-bridge-manifest.xml: Operation not permitted');
  const r2 = detectPermissionIssueCore(office, 'darwin');
  assert.ok(r2, 'Office 侧 Operation not permitted 必须被识别');
  assert.equal(r2!.targetPath, '/Users/x/Library/Containers/com.microsoft.Excel/Data/Documents/wef/wps-bridge-manifest.xml');
});

test('EACCES 与带 code 的错误同样识别', () => {
  const e: any = new Error('写入失败'); e.code = 'EACCES';
  assert.ok(detectPermissionIssueCore(e, 'darwin'));
  const e2: any = new Error('open failed'); e2.code = 'EPERM';
  assert.ok(detectPermissionIssueCore(e2, 'darwin'));
});

test('无关失败不得被误判为权限问题', () => {
  for (const e of [new Error('ENOENT: no such file or directory'), new Error('磁盘已满'), new Error('目标清单 XML 语法错误'), 'timeout']) {
    assert.equal(detectPermissionIssueCore(e, 'darwin'), null, `不应误判: ${String(e)}`);
  }
});

test('非 macOS 平台不产生引导（Windows 走各自路径）', () => {
  assert.equal(detectPermissionIssueCore(new Error('EPERM: operation not permitted'), 'win32'), null);
  assert.equal(detectPermissionIssueCore(new Error('EPERM: operation not permitted'), 'linux'), null);
});

test('无路径的权限错误也能识别，路径字段留空而不是编造', () => {
  const r = detectPermissionIssueCore(new Error('EPERM: operation not permitted'), 'darwin');
  assert.ok(r);
  assert.equal(r!.targetPath, '', '提取不到路径时应为空字符串，不得编造');
  assert.equal(extractProtectedPath('无关文本'), '');
});
