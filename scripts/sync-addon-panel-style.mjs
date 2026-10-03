// Office 侧栏样式为唯一来源，WPS 部署副本自包含，避免文件跨目录引用失效。
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export function syncAddonPanelStyle(root = ROOT, checkOnly = false) {
  const source = fs.readFileSync(path.join(root, 'office-addon/public/taskpane.css'), 'utf8');
  const output = '/* 由 scripts/sync-addon-panel-style.mjs 生成；统一样式请改 office-addon/public/taskpane.css */\n' + source;
  const target = path.join(root, 'wps-addon/panel.css');
  const matches = fs.existsSync(target) && fs.readFileSync(target, 'utf8') === output;
  if (checkOnly && !matches) throw new Error('WPS 侧栏样式与 Office 来源不一致，请重新构建加载项');
  if (!checkOnly && !matches) fs.writeFileSync(target, output);
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  syncAddonPanelStyle(ROOT, process.argv.includes('--check'));
  console.log('WPS / Office 侧栏样式已同步。');
}
