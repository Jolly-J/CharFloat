import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

export const ROOT = fileURLToPath(new URL('../../', import.meta.url));
export const LOCAL_SECRETS = '.charfloat-local-secrets';
export const PRIVATE_KEY = 'production-ed25519-private.pem';
export const KEY_DIRECTORY = `${LOCAL_SECRETS}/update-signing`;
export const KEY_EXCLUSIONS = [`!**/${LOCAL_SECRETS}`, `!**/${LOCAL_SECRETS}/**`];
export function privateUpdatePath(value) {
  const parts = String(value).replaceAll('\\', '/').toLowerCase().split('/');
  return parts.includes(LOCAL_SECRETS) || parts.includes(PRIVATE_KEY);
}
export function assertPublicPaths(names) {
  if (names.some(privateUpdatePath)) throw new Error('文件清单包含本地更新签名材料，已拒绝构建或归档。');
}
export function checkKeyBoundaries(root = ROOT) {
  const privatePath = `${KEY_DIRECTORY}/${PRIVATE_KEY}`;
  if (!fs.readFileSync(path.join(root, '.gitignore'), 'utf8').split(/\r?\n/).includes(`/${LOCAL_SECRETS}/`)) throw new Error('缺少本地签名目录 Git 排除规则。');
  const npmIgnore = path.join(root, '.npmignore');
  if (fs.existsSync(npmIgnore) && !fs.readFileSync(npmIgnore, 'utf8').split(/\r?\n/).includes(`/${LOCAL_SECRETS}/`)) throw new Error('npmignore 覆盖 Git 规则但未排除本地签名目录。');
  const attrs = fs.readFileSync(path.join(root, '.gitattributes'), 'utf8');
  if (!attrs.split(/\r?\n/).includes(`${LOCAL_SECRETS} export-ignore`) || !attrs.split(/\r?\n/).includes(`${LOCAL_SECRETS}/** export-ignore`)) throw new Error('缺少本地签名目录归档排除规则。');
  execFileSync('git', ['check-ignore', '-q', '--', privatePath], { cwd: root, stdio: 'ignore' });
  const tracked = execFileSync('git', ['ls-files', '-z', '--', LOCAL_SECRETS], { cwd: root, encoding: 'utf8' });
  if (tracked.length) throw new Error('本地签名目录中存在已跟踪文件，已停止。');
  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  if (!KEY_EXCLUSIONS.every(value => pkg.build?.files?.includes(value))) throw new Error('缺少应用包签名目录排除规则。');
  const stage = pkg.build.files.find(value => value?.from === 'dist/release-app');
  if (!stage || !KEY_EXCLUSIONS.every(value => stage.filter?.includes(value))) throw new Error('缺少发布副本复制排除规则。');
  const policy = JSON.parse(fs.readFileSync(path.join(root, 'scripts/release-files.json'), 'utf8'));
  assertPublicPaths(policy.files);
  return { gitIgnored: true, gitTrackedSecrets: 0, archiveExcluded: true, applicationExcluded: true, releaseAllowlistExcluded: true };
}
