/**
 * 构建指纹（ISS-59）。
 *
 * 问题：本项目有**两个独立的代码落点**——WPS/Office 加载项（部署 + 宿主重载）与桥接后台
 * （`build:main` + 后台重启）。磁盘上是新构建、进程里跑的却是旧构建时，双方版本号都写着
 * "2.1.0"，**没有任何指纹能判断"运行中的是哪一版"**。ISS-59 因此复现过两次。
 *
 * 本模块给出桥接侧能做到的最大范围：把**磁盘产物**的构建指纹（sha256 + mtime + size + 路径）
 * 如实暴露出来，并给出"部署副本 vs 随包资源"是否一致、"连接上报的版本 vs 当前桥接版本"是否一致。
 *
 * 诚实边界（不要把它当成"运行中版本的证明"）：
 * - 加载项目前**不上报自身构建指纹**（只上报版本号），因此"进程里加载的到底是哪一份字节"
 *   **无法证实**，只能靠"宿主重载/重启后重新 register"间接保证；
 * - 本模块只做只读读取：不部署、不重载、不改任何文件。
 */

import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import crypto from 'node:crypto';
import { VERSION, resourcePath } from './runtime.js';

export interface ArtifactFingerprint {
  /** 产物用途说明。 */
  role: string;
  path: string;
  exists: boolean;
  sha256?: string;
  sizeBytes?: number;
  mtime?: string;
  error?: string;
}

export interface BuildFingerprints {
  capturedAt: number;
  bridgeVersion: string;
  /** 运行中的桥接入口产物（dist/bridge/cli.cjs 等）。 */
  bridgeEntry: ArtifactFingerprint;
  /** WPS 加载项：随包资源 + 已部署副本。 */
  wpsAddon: {
    resource: ArtifactFingerprint;
    deployed: ArtifactFingerprint[];
    /** 部署副本与随包资源的 sha256 是否一致；任一不可读时为 null（判不了）。 */
    deployedMatchesResource: boolean | null;
  };
  /** Microsoft Office.js 任务窗格产物（由桥接 HTTP 静态托管）。 */
  officeAddon: { resource: ArtifactFingerprint };
  note: string;
}

function fingerprint(role: string, filePath: string | null): ArtifactFingerprint {
  if (!filePath) return { role, path: '(未找到)', exists: false, error: '路径不可解析' };
  try {
    const stat = fs.statSync(filePath);
    const sha256 = crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex');
    return { role, path: filePath, exists: true, sha256, sizeBytes: stat.size, mtime: new Date(stat.mtimeMs).toISOString() };
  } catch (error: any) {
    return { role, path: filePath, exists: false, error: error?.code || error?.message || String(error) };
  }
}

/** 解析随包资源；资源缺失时返回 null，而不是抛错（诊断工具不应因缺文件而失败）。 */
function tryResource(relative: string): string | null {
  try { return resourcePath(relative); } catch { return null; }
}

/**
 * WPS 用户加载项目录（只读镜像 `src/main/addon-installer.ts` 的候选路径）。
 *
 * 这里不 import `src/main/**`：桥接执行层不能反向依赖 Electron 主进程模块（分层约定）。
 * 路径列表与安装器保持一致，仅用于**读取**已部署副本做指纹比对。
 */
function wpsAddonDirectories(): string[] {
  if (process.env.WPS_BRIDGE_ADDON_DIR) return [process.env.WPS_BRIDGE_ADDON_DIR];
  const home = os.homedir();
  const candidates = process.platform === 'win32'
    ? [path.join(process.env.APPDATA || path.join(home, 'AppData/Roaming'), 'kingsoft/wps/jsaddons')]
    : process.platform === 'darwin'
      ? [
        'Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons',
        'Library/Containers/cn.wps.moffice_mac/Data/.kingsoft/wps/jsaddons',
        'Library/Containers/com.kingsoft.wpsoffice.mac/Data/.local/share/Kingsoft/wps/jsaddons',
        'Library/Containers/cn.wps.moffice_mac/Data/.local/share/Kingsoft/wps/jsaddons',
        'Library/Application Support/Kingsoft/wps/jsaddons'
      ].map(p => path.join(home, p))
      : [];
  return candidates.filter(p => fs.existsSync(p) || fs.existsSync(path.dirname(p)));
}

let cache: { at: number; value: BuildFingerprints } | undefined;
const CACHE_MS = 10_000;

/** 采集构建指纹（10 秒缓存：状态接口可能被频繁轮询，避免反复读取大文件）。 */
export function collectBuildFingerprints(force = false): BuildFingerprints {
  if (!force && cache && Date.now() - cache.at < CACHE_MS) return cache.value;

  const wpsResource = fingerprint('WPS 加载项随包产物', tryResource('wps-addon/addon-core.js'));
  const deployed = wpsAddonDirectories()
    .map(dir => path.join(dir, 'wps-bridge', 'addon-core.js'))
    .filter(file => fs.existsSync(file))
    .map(file => fingerprint('WPS 加载项已部署副本', file));

  const comparable = wpsResource.exists && deployed.length > 0 && deployed.every(d => d.exists && d.sha256);
  const deployedMatchesResource = comparable
    ? deployed.every(d => d.sha256 === wpsResource.sha256)
    : null;

  const value: BuildFingerprints = {
    capturedAt: Date.now(),
    bridgeVersion: VERSION,
    bridgeEntry: fingerprint('桥接入口产物', process.argv[1] ? path.resolve(process.argv[1]) : null),
    wpsAddon: { resource: wpsResource, deployed, deployedMatchesResource },
    officeAddon: { resource: fingerprint('Office.js 任务窗格产物（由桥接托管）', tryResource('office-addon/public/taskpane.js')) },
    note:
      '指纹只证明**磁盘/部署副本**的构建身份；加载项目前不上报自身指纹，' +
      '"进程里加载的是哪一份字节"无法由此证实（需宿主重载/重启后重新 register）。未实机验收。'
  };
  cache = { at: Date.now(), value };
  return value;
}

/** 判断连接上报的版本是否与当前桥接产物版本一致。 */
export function reportedVersionStatus(reported?: string) {
  if (!reported || reported === '未知') return { reportedVersion: reported ?? null, matchesBridgeVersion: null, staleSuspect: null };
  const matches = reported === VERSION || reported.startsWith(`${VERSION}.`);
  return { reportedVersion: reported, matchesBridgeVersion: matches, staleSuspect: !matches };
}
