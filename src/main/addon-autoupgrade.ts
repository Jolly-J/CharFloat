/**
 * 加载项自动升级（客户端启动时执行）。
 *
 * ## 解决的问题
 * 客户端包内带着新的加载项构建，而 WPS 里部署的可能还是旧构建。
 * 此前只能靠用户手动点【一键修复 / 升级加载项】，而且**升级后还得彻底退出并重开 WPS**
 * 才生效——使用者为此重启了十几次。
 *
 * ## 三层配合
 * 1. **检测**：`AddonInstaller.checkBuildFreshness()` 比**构建指纹**。
 *    不能只比版本号——`manifest.xml` 的 2.1.0 在多次构建之间不变，
 *    "版本号相同"完全不能说明"跑的是同一份构建"。
 * 2. **部署**：`AddonInstaller.install()`。部署时会给 `index.html` 的
 *    `<script src="./addon-core.js">` 附上构建指纹查询串。
 * 3. **生效**：调桥接的 `wps_reload_addon` 触发加载项 `window.location.reload()`。
 *    因为 `<script src>` 带了新指纹（URL 变了），重载会**真正取到新文件**；
 *    没有这一步时 WPS 只会重跑缓存里的旧 JS。
 *
 * 三者缺一：只检测不部署＝永远提示过期；只部署不重载＝磁盘新、进程旧（ISS-59）。
 */

import { AddonInstaller } from './addon-installer.js';
import { serviceRequest } from '../bridge/service-client.js';

export interface AddonAutoUpgradeResult {
  /** 是否执行了检测 */
  checked: boolean;
  /** 部署副本是否与包内构建不一致；null = 判不了（读不到包内指纹） */
  stale: boolean | null;
  /** 是否执行了重新部署 */
  upgraded: boolean;
  /** 是否触发了加载项重载 */
  reloaded: boolean;
  /** 结果说明（可直接展示给用户） */
  message: string;
  /** 诊断细节 */
  detail?: unknown;
}

/** 加载项重载后等待其重连的时间（毫秒）。 */
const RELOAD_SETTLE_MS = 4000;

/**
 * 确保已部署的加载项与客户端包内的构建一致；不一致则自动重新部署并让它重新加载。
 *
 * @param options.force      忽略"是否需要"的判定，强制重新部署一次
 * @param options.reload     部署后是否触发加载项重载（默认 true）
 * @param options.installFn  可注入的部署函数（默认 `AddonInstaller.install`，便于测试）
 * @param options.reloadFn   可注入的重载函数（默认走桥接的 `wps_reload_addon`）
 */
export async function ensureAddonUpToDate(options: {
  force?: boolean;
  reload?: boolean;
  installFn?: () => Promise<any>;
  reloadFn?: () => Promise<any>;
} = {}): Promise<AddonAutoUpgradeResult> {
  const { force = false, reload = true } = options;

  let freshness: ReturnType<typeof AddonInstaller.checkBuildFreshness>;
  try {
    freshness = AddonInstaller.checkBuildFreshness();
  } catch (e: any) {
    return { checked: false, stale: null, upgraded: false, reloaded: false, message: `无法检测加载项构建状态：${e?.message ?? e}` };
  }

  // 判不了（读不到包内指纹）时**不猜**：宁可不动，也不要盲目重新部署
  if (freshness.stale === null) {
    return { checked: true, stale: null, upgraded: false, reloaded: false, message: freshness.reason, detail: freshness };
  }
  if (!freshness.stale && !force) {
    return { checked: true, stale: false, upgraded: false, reloaded: false, message: freshness.reason, detail: freshness };
  }

  const installFn = options.installFn ?? (() => AddonInstaller.install());
  let installResult: any;
  try {
    installResult = await installFn();
  } catch (e: any) {
    return { checked: true, stale: true, upgraded: false, reloaded: false, message: `重新部署加载项失败：${e?.message ?? e}`, detail: freshness };
  }

  // 部署失败（含权限问题）时不再尝试重载——重载只会加载旧文件，制造"已修复"的假象
  if (installResult && installResult.success === false) {
    return {
      checked: true,
      stale: true,
      upgraded: false,
      reloaded: false,
      message: `重新部署加载项未成功：${installResult.message ?? '见安装器返回'}`,
      detail: { freshness, installResult }
    };
  }

  if (!reload) {
    return {
      checked: true, stale: true, upgraded: true, reloaded: false,
      message: '已重新部署加载项；未触发重载（调用方要求跳过），运行中的加载项要等下次重载或重启 WPS 才生效',
      detail: { freshness, installResult }
    };
  }

  const reloadFn = options.reloadFn ?? (() => serviceRequest('/api/v1/tool/call', { name: 'wps_reload_addon', arguments: {}, sessionId: 'desktop' }));
  try {
    await reloadFn();
    // 加载项 reload 后会断开重连，稍等再让调用方查状态
    await new Promise(resolve => setTimeout(resolve, RELOAD_SETTLE_MS));
    return {
      checked: true, stale: true, upgraded: true, reloaded: true,
      message: '已重新部署加载项并触发重新加载；运行中的加载项无需重启 WPS 即可生效',
      detail: { freshness, installResult }
    };
  } catch (e: any) {
    return {
      checked: true, stale: true, upgraded: true, reloaded: false,
      message: `加载项已重新部署，但触发重载失败（${e?.message ?? e}）：需要手动点击功能区【重新连接】或重启 WPS`,
      detail: { freshness, installResult }
    };
  }
}
