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

/** 从加载项产物头部注释解析构建指纹（由 scripts/build-wps-addon.mjs 注入）。 */
export function readAddonFingerprint(filePath: string | null): string | null {
  if (!filePath) return null;
  try {
    // 只读头部若干字节即可，避免每次比对都读整个产物
    const fd = fs.openSync(filePath, 'r');
    try {
      const buf = Buffer.alloc(4096);
      const read = fs.readSync(fd, buf, 0, buf.length, 0);
      const m = buf.toString('utf8', 0, read).match(/ADDON_BUILD_FINGERPRINT:\s*([0-9a-f]{16,64})/);
      return m ? m[1] : null;
    } finally {
      fs.closeSync(fd);
    }
  } catch {
    return null;
  }
}

/**
 * ISS-59：判定**WPS 进程里正在运行的加载项**是不是最新的那一份构建。
 *
 * 加载项在注册报文里上报运行时构建指纹，这里与磁盘/已部署副本比对：
 *   - `reported` 来自**进程内**（真正在跑的字节）；
 *   - `resource` 是仓库里的构建产物；
 *   - `deployed` 是安装到 WPS 加载项目录的副本。
 * 典型故障：部署了新构建但没重载 WPS → `reported` 是旧的、`resource`/`deployed` 是新的。
 * 这正是此前"部署了但没生效"无法被识别的状态。
 */
export function reportedAddonBuildStatus(reported?: string | null) {
  const fingerprints = collectBuildFingerprints();
  const resource = readAddonFingerprint(fingerprints.wpsAddon.resource.exists ? fingerprints.wpsAddon.resource.path : null);
  const deployed = fingerprints.wpsAddon.deployed
    .map(d => readAddonFingerprint(d.exists ? d.path : null))
    .filter((f): f is string => Boolean(f));

  if (!reported) {
    return {
      reportedBuild: null,
      resourceBuild: resource,
      deployedBuilds: deployed,
      runningMatchesResource: null,
      runningMatchesDeployed: null,
      staleSuspect: null,
      note: '加载项未上报构建指纹（多半是旧构建；重载加载项后即可上报）'
    };
  }

  const stale = resource ? reported !== resource : null;
  return {
    reportedBuild: reported,
    resourceBuild: resource,
    deployedBuilds: deployed,
    runningMatchesResource: resource ? reported === resource : null,
    runningMatchesDeployed: deployed.length ? deployed.includes(reported) : null,
    staleSuspect: stale,
    note: stale
      ? 'WPS 进程里运行的加载项与磁盘构建**不是同一份**：新构建已生成（或已部署）但未重载，请重新加载加载项后再验证。'
      : '进程内构建与磁盘构建一致。'
  };
}
