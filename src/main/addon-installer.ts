import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { execSync } from 'node:child_process';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { atomicWrite, getToken, resourcePath, runtimePort, runtimeHome, getOrGenerateCerts, appendServiceLog, VERSION } from '../bridge/runtime.js';
import { detectPermissionIssueCore } from './permission-detect.js';
import { serviceRequest } from '../bridge/service-client.js';

const activeNames = ['Office Agent Bridge', 'Office Agent Bridge (表格)', 'Office Agent Bridge (文字)', 'Office Agent Bridge (演示)'];
const legacyNames = ['WPS Bridge', 'WPS Bridge (表格)', 'WPS Bridge (文字)', 'WPS Bridge (演示)'];
const ownedNames = [...activeNames, ...legacyNames];

export function readXmlSafe(file: string): string {
  if (!fs.existsSync(file)) return '';
  const buf = fs.readFileSync(file);
  if (!buf.length) return '';
  appendServiceLog('AddonInstaller', `读取 ${file}: ${buf.length} 字节, 前导十六进制: ${buf.subarray(0, 16).toString('hex')}`);
  // 1. 检测 UTF-16 LE BOM (ff fe) 或 UTF-16 BE BOM (fe ff) 以及无 BOM 的 UTF-16LE (3c 00)
  if (buf.length >= 2) {
    if (buf[0] === 0xff && buf[1] === 0xfe) return new TextDecoder('utf-16le').decode(buf.subarray(2));
    if (buf[0] === 0xfe && buf[1] === 0xff) return new TextDecoder('utf-16be').decode(buf.subarray(2));
    if (buf[0] === 0x3c && buf[1] === 0x00) return new TextDecoder('utf-16le').decode(buf);
    if (buf[0] === 0x00 && buf[1] === 0x3c) return new TextDecoder('utf-16be').decode(buf);
  }
  // 2. 剥离 UTF-8 BOM (\uFEFF)
  if (buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) {
    return new TextDecoder('utf-8').decode(buf.subarray(3));
  }
  // 3. 检测 XML 声明中的 encoding
  const head = buf.subarray(0, 120).toString('ascii').toLowerCase();
  if (head.includes('encoding="utf-16"') || head.includes("encoding='utf-16'")) {
    try { return new TextDecoder('utf-16le').decode(buf); } catch {}
  }
  if (head.includes('encoding="gbk"') || head.includes("encoding='gbk'") || head.includes('encoding="gb2312"') || head.includes("encoding='gb2312'")) {
    try { return new TextDecoder('gbk').decode(buf); } catch {}
  }
  // 4. 优先 utf-8，若存在非法替换字节 (\uFFFD) 或 NUL 字节，尝试 gbk
  const utf8 = new TextDecoder('utf-8').decode(buf);
  if (utf8.includes('\uFFFD') || utf8.includes('\0')) {
    try {
      const gbk = new TextDecoder('gbk').decode(buf);
      if (!gbk.includes('\uFFFD') && !gbk.includes('\0')) return gbk;
    } catch {}
  }
  return utf8.replace(/\0/g, '');
}

export function mergePluginIndex(raw: string, remove = false) {
  let text = (raw || '').replace(/^\uFEFF/, '').replace(/\0/g, '').trim();
  // 剔除 XML 声明头 (<?xml ... ?>) 与 注释 (<!-- ... -->) 后检测正文是否有标签
  const body = text.replace(/<\?xml[\s\S]*?\?>/gi, '').replace(/<!--[\s\S]*?-->/gi, '').trim();
  if (!body || !body.includes('<')) {
    text = '<jsplugins/>';
  }
  let doc: any;
  try {
    doc = new DOMParser({
      onError: (level, msg) => {
        // 如果是 missing root element，由外层 catch 安全初始化为 <jsplugins/>，不作为致命错误中断
        if (level === 'fatalError' && !msg.includes('missing root element')) {
          throw new Error(`插件索引 XML 无效，停止写入: ${msg}`);
        }
      }
    }).parseFromString(text, 'application/xml');
  } catch (e: any) {
    if (String(e.message || e).includes('missing root element')) {
      appendServiceLog('AddonInstaller', '检测到索引 XML 缺少根节点 (missing root element)，自动安全初始化为 <jsplugins/>');
      doc = new DOMParser().parseFromString('<jsplugins/>', 'application/xml');
    } else {
      appendServiceLog('AddonInstaller', `插件索引 XML 语法错误拦截: ${e.message}`);
      throw e;
    }
  }

  // 兜底：如果解析结果没有 documentElement（空根节点）
  if (!doc || !doc.documentElement) {
    appendServiceLog('AddonInstaller', '解析后未生成 documentElement，自动安全初始化为 <jsplugins/>');
    doc = new DOMParser().parseFromString('<jsplugins/>', 'application/xml');
  } else if (doc.documentElement.tagName !== 'jsplugins') {
    appendServiceLog('AddonInstaller', `插件索引根节点不是 jsplugins (当前为: <${doc.documentElement.tagName}>)，停止写入`);
    throw new Error(`插件索引根节点不是 jsplugins，停止写入`);
  }

  const nodes = Array.from(doc.getElementsByTagName('jsplugin')) as any[];
  for (const node of nodes) if (ownedNames.includes(node.getAttribute('name') || '')) node.parentNode!.removeChild(node);
  if (!remove) for (const [i, type] of ['et', 'wps', 'wpp'].entries()) {
    const node = doc.createElement('jsplugin');
    for (const [key, value] of Object.entries({ name: activeNames[i + 1], type, url: './wps-bridge', enable: 'true', autoload: 'true', version: VERSION })) node.setAttribute(key, value);
    doc.documentElement!.appendChild(node);
  }
  return new XMLSerializer().serializeToString(doc);
}
function isHostBlocked(dir: string): boolean {
  const blockFile = path.join(dir, 'jsaddinblockhost.ini');
  if (!fs.existsSync(blockFile)) return false;
  try {
    const content = fs.readFileSync(blockFile, 'utf8');
    return /127\.0\.0\.1|localhost|19890|19891|wps-bridge|office-agent-bridge/i.test(content);
  } catch {
    return false;
  }
}

export class AddonInstaller {
  static getAllAddonDirectories() {
    if (process.env.WPS_BRIDGE_ADDON_DIR) return [process.env.WPS_BRIDGE_ADDON_DIR];
    const home = os.homedir();
    const candidates = process.platform === 'win32' ? [path.join(process.env.APPDATA || path.join(home, 'AppData/Roaming'), 'kingsoft/wps/jsaddons')] : process.platform === 'darwin' ? [
      'Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons',
      'Library/Containers/cn.wps.moffice_mac/Data/.kingsoft/wps/jsaddons',
      'Library/Containers/com.kingsoft.wpsoffice.mac/Data/.local/share/Kingsoft/wps/jsaddons',
      'Library/Containers/cn.wps.moffice_mac/Data/.local/share/Kingsoft/wps/jsaddons',
      'Library/Application Support/Kingsoft/wps/jsaddons'
    ].map(p => path.join(home, p)) : [];
    return candidates.filter(p => fs.existsSync(p) || fs.existsSync(path.dirname(p)));
  }
  static getAddonDirectory() { return this.getAllAddonDirectories()[0] || null; }
  static getSourceAddonPath() { return resourcePath('wps-addon'); }

  /**
   * 读取加载项产物头部注入的构建指纹。
   *
   * 为什么需要它：`manifest.xml` 的版本号（2.1.0）在多次构建之间**不变**，
   * 所以"版本号相同"完全不能说明"跑的是同一份构建"。指纹每次构建都变，
   * 才是判断"已部署的是不是包内这一份"的可靠依据。
   */
  static readAddonFingerprint(file: string | null): string | null {
    if (!file) return null;
    try {
      const fd = fs.openSync(file, 'r');
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
   * 比对"客户端包内的加载项构建"与"已部署到 WPS 的构建"。
   *
   * 判定口径（任一即视为需要升级）：
   *   - 没装（找不到部署副本，或副本里读不到指纹）；
   *   - 已部署指纹 ≠ 包内指纹（包更新了但没重新部署，或部署后被改过）。
   * 读不到包内指纹时返回 `stale: null`（判不了），**不猜测**。
   */
  static checkBuildFreshness() {
    const bundledPath = path.join(this.getSourceAddonPath(), 'addon-core.js');
    const bundled = this.readAddonFingerprint(fs.existsSync(bundledPath) ? bundledPath : null);
    const deployed = this.getAllAddonDirectories()
      .map(dir => {
        const file = path.join(dir, 'wps-bridge', 'addon-core.js');
        return { dir, file, exists: fs.existsSync(file), fingerprint: this.readAddonFingerprint(fs.existsSync(file) ? file : null) };
      })
      .filter(d => d.exists);

    if (bundled === null) {
      return {
        bundledFingerprint: null,
        bundledPath,
        deployed,
        stale: null as boolean | null,
        reason: '读不到客户端包内的构建指纹（产物缺失或未注入），无法判断是否需要升级'
      };
    }
    if (deployed.length === 0) {
      return {
        bundledFingerprint: bundled,
        bundledPath,
        deployed,
        stale: true as boolean | null,
        reason: '未找到已部署的加载项副本，需要安装'
      };
    }
    const mismatched = deployed.filter(d => d.fingerprint !== bundled);
    return {
      bundledFingerprint: bundled,
      bundledPath,
      deployed,
      stale: mismatched.length > 0,
      mismatchedCount: mismatched.length,
      reason: mismatched.length > 0
        ? `已部署的加载项与客户端包内的构建不一致（${mismatched.length}/${deployed.length} 个副本）：包内 ${bundled.slice(0, 12)}…，部署 ${(mismatched[0].fingerprint || '未知').slice(0, 12)}…`
        : '已部署的加载项与客户端包内的构建一致'
    };
  }

  static checkStatus() {
    const dirs = this.getAllAddonDirectories();
    let installedVersion = '';
    const details = dirs.map(dir => {
      const manifest = path.join(dir, 'wps-bridge/manifest.xml');
      const config = path.join(dir, 'wps-bridge/bridge-config.js');
      const installed = fs.existsSync(manifest);
      let current = false, error = '', ver = '';
      try {
        if (installed) {
          const content = fs.readFileSync(manifest, 'utf8');
          const m = content.match(/<version>(.*?)<\/version>/);
          ver = m ? m[1].trim() : '';
          if (!installedVersion && ver) installedVersion = ver;
          current = ver === VERSION && fs.existsSync(config);
        }
      } catch(e: any) { error = `无法读取加载项文件：${e.code || e.message}`; }
      return { dir, installed, current, version: ver, error, blocked: isHostBlocked(dir) };
    });
    const installed = details.some(d => d.installed);
    const current = installed && (installedVersion === VERSION || details.some(d => d.installed && d.current));
    const hasBlocked = details.some(d => d.blocked);
    // 版本号相同**不代表构建相同**（2.1.0 之间会反复重建）：还要比构建指纹，
    // 否则"装了旧构建但版本号一致"会被判成最新，AI 拿到的一直是旧代码。
    const buildFreshness = this.checkBuildFreshness();
    const buildStale = buildFreshness.stale === true;
    const needsUpgrade = !installed || !current || buildStale;
    return {
      installed,
      current,
      latestVersion: VERSION,
      installedVersion: installedVersion || (installed ? '未知' : '未安装'),
      needsUpgrade,
      hasBlocked,
      buildFreshness,
      platform: process.platform,
      targetPath: dirs.join(' | '),
      details,
      message: !dirs.length
        ? '未找到 WPS 用户加载项目录，请先运行 WPS 或指定目录'
        : hasBlocked
          ? '检测到加载项被 WPS 阻断，请点击【一键修复 / 升级加载项】解除阻断'
          : needsUpgrade
            ? buildStale
              ? `检测到加载项构建不是包内这一份（版本号同为 ${VERSION} 但构建指纹不同），需要重新部署`
              : `检测到加载项需要更新（当前: ${installedVersion || '未知'}，最新: ${VERSION}）`
            : `加载项已是最新版本 (v${VERSION}，构建指纹一致)`
    };
  }
  static install(options: { cleanBlocked?: boolean } = { cleanBlocked: true }) {
    appendServiceLog('AddonInstaller', '=== 开始执行 WPS 加载项安装 / 升级 ===');
    try {
      const dirs = this.getAllAddonDirectories();
      appendServiceLog('AddonInstaller', `检测到 WPS 目标加载项目录: ${dirs.length ? dirs.join(' | ') : '未找到任何目录'}`);
      if (!dirs.length) throw new Error('找不到 WPS 加载项目录，请先安装并运行 WPS。');
      const source = this.getSourceAddonPath();
      appendServiceLog('AddonInstaller', `安装源资源目录: ${source}`);
      // Validate every index before changing anything.
      const indexes = dirs.flatMap(dir => ['publish.xml', 'jsplugins.xml'].map(name => {
        const file = path.join(dir, name);
        appendServiceLog('AddonInstaller', `准备校验/合并索引文件: ${file}`);
        const content = mergePluginIndex(readXmlSafe(file));
        return { file, content };
      }));
      const token = getToken(true);
      for (const dir of dirs) {
        if (options.cleanBlocked !== false) {
          const blockFile = path.join(dir, 'jsaddinblockhost.ini');
          if (fs.existsSync(blockFile)) {
            try {
              fs.copyFileSync(blockFile, `${blockFile}.backup-${Date.now()}`);
              fs.unlinkSync(blockFile);
              appendServiceLog('AddonInstaller', `已清理阻断文件: ${blockFile}`);
            } catch (err: any) {
              appendServiceLog('AddonInstaller', `清理阻断文件异常: ${err.message}`);
            }
          }
        }

        // 清理老旧版本残余目录（避免 WPS 缓存或回退）
        try {
          const entries = fs.readdirSync(dir);
          for (const entry of entries) {
            const isOwned = ownedNames.some(name => entry.startsWith(name));
            if (isOwned && !entry.endsWith(`_${VERSION}`) && entry !== 'wps-bridge') {
              // 保留主软链接/目录，清理其他带历史版本后缀的目录
              const target = path.join(dir, entry);
              if (fs.existsSync(target) && fs.statSync(target).isDirectory() && entry.includes('.')) {
                try { fs.rmSync(target, { recursive: true, force: true }); } catch {}
              }
            }
          }
        } catch {}

        const aliases = ['wps-bridge', 'office-agent-bridge', ...activeNames.slice(1).flatMap(name => [`${name}_`, `${name}_${VERSION}`])];
        for (const alias of aliases) {
          const dest = path.join(dir, alias);
          fs.mkdirSync(dest, { recursive: true, mode: 0o700 });
          for (const file of fs.readdirSync(source)) {
            const srcFile = path.join(source, file);
            if (fs.statSync(srcFile).isFile()) {
              fs.copyFileSync(srcFile, path.join(dest, file));
            }
          }
          atomicWrite(path.join(dest, 'bridge-config.js'), `window.WPS_BRIDGE_CONFIG = ${JSON.stringify({ port: runtimePort(), token, version: VERSION })};\n`);

          // 让"重新加载加载项"能真正吃到新构建。
          //
          // WPS 会缓存 addon-core.js，只调 window.location.reload() 只是**重跑缓存里的旧 JS**
          // （实测：部署后调用 reload，运行中的加载项仍没有新增的构建指纹变量）。
          // 这里给 <script src> 附上构建指纹查询串：每次构建 URL 都不同 → reload 必然重新取文件，
          // 于是"部署 + 重新连接"即可生效，不必彻底退出并重开 WPS。
          const addonEntry = path.join(dest, 'addon-core.js');
          const indexPath = path.join(dest, 'index.html');
          if (fs.existsSync(addonEntry) && fs.existsSync(indexPath)) {
            try {
              const head = fs.readFileSync(addonEntry, 'utf8').slice(0, 4096);
              const fp = head.match(/ADDON_BUILD_FINGERPRINT:\s*([0-9a-f]{16,64})/);
              if (fp) {
                const html = fs.readFileSync(indexPath, 'utf8');
                const patched = html.replace(
                  /src="\.\/addon-core\.js(\?v=[0-9a-f]+)?"/g,
                  `src="./addon-core.js?v=${fp[1].slice(0, 16)}"`
                );
                if (patched !== html) fs.writeFileSync(indexPath, patched, 'utf8');
              }
            } catch (e) {
              appendServiceLog('AddonInstaller', `缓存失效改写失败（不影响部署）: ${(e as Error).message}`);
            }
          }
        }
      }
      for (const index of indexes) {
        atomicWrite(index.file, index.content, true);
        appendServiceLog('AddonInstaller', `写入更新索引文件成功: ${index.file}`);
      }
      appendServiceLog('AddonInstaller', `=== WPS 加载项部署成功完成 (v${VERSION}) ===`);
      return { success: true, message: `加载项已成功部署并更新至 v${VERSION}！请在 WPS 中点击功能区【重新连接】或重启 WPS 生效。`, targetPath: dirs.join(' | '), warnings: dirs.filter(dir => fs.existsSync(path.join(dir, 'jsaddinblockhost.ini'))).map(() => '检测到 WPS 阻断配置，未删除。请在 WPS 中检查加载项权限。') };
    } catch (e: any) {
      appendServiceLog('AddonInstaller', `WPS 加载项安装失败: ${e.message}\n堆栈: ${e.stack || ''}`);
      // 权限受限不是普通失败：交给界面走引导流程，而不是甩一条裸 EPERM 文案给用户。
      const permissionIssue = detectPermissionIssueCore(e);
      if (permissionIssue) return { success: false, message: 'WPS 加载项目录需要 macOS 完全磁盘访问权限', permissionIssue };
      return { success: false, message: e.message };
    }
  }
  static uninstall() {
    try {
      const indexes = this.getAllAddonDirectories().flatMap(dir => ['publish.xml', 'jsplugins.xml'].map(name => path.join(dir, name))).filter(file => fs.existsSync(file)).map(file => ({ file, content: mergePluginIndex(readXmlSafe(file), true) }));
      for (const item of indexes) atomicWrite(item.file, item.content, true);
      return { success: true, message: '已注销本项目加载项；保留文件与备份，重启 WPS 后生效。' };
    } catch(e: any) { return { success: false, message: e.message }; }
  }
}

export class OfficeAddonInstaller {
  static getWefDirectory() {
    const home = os.homedir();
    if (process.platform === 'win32') {
      return path.join(process.env.LOCALAPPDATA || path.join(home, 'AppData/Local'), 'Microsoft/Office/16.0/Wef');
    }
    // macOS 官方 Excel 侧载目录
    return path.join(home, 'Library/Containers/com.microsoft.Excel/Data/Documents/wef');
  }

  static getSourceManifestPath() {
    return resourcePath('office-addon/excel/manifest.xml');
  }

  private static safeRead(filePath: string): string {
    if (!fs.existsSync(filePath)) return '';
    try {
      return fs.readFileSync(filePath, 'utf8');
    } catch {
      if (process.platform === 'darwin') {
        try {
          return execSync(`cat "${filePath}"`, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] });
        } catch {}
      }
    }
    return '';
  }

  private static safeDeploy(sourceFile: string, destFile: string): void {
    appendServiceLog('AddonInstaller', `准备部署 Office 清单: 源文件=${sourceFile}, 目标文件=${destFile}`);
    if (!fs.existsSync(sourceFile)) {
      const msg = `源清单文件不存在: ${sourceFile}`;
      appendServiceLog('AddonInstaller', msg);
      throw new Error(msg);
    }

    const sourceContent = fs.readFileSync(sourceFile, 'utf8');

    // 1. 若目标文件已存在且内容一致，直接跳过覆盖，避免无谓触发系统权限限制
    if (fs.existsSync(destFile)) {
      try {
        const destContent = fs.readFileSync(destFile, 'utf8');
        if (destContent === sourceContent || (destContent.includes(`<Version>${VERSION}`) && destContent.includes('Office Agent Bridge'))) {
          appendServiceLog('AddonInstaller', `目标清单文件已处于最新状态 (v${VERSION})，跳过覆写`);
          if (process.platform === 'darwin') {
            try { execSync(`chmod 644 "${destFile}"`, { stdio: 'ignore' }); } catch {}
          }
          return;
        }
      } catch (e: any) {
        appendServiceLog('AddonInstaller', `比对目标清单失败: ${e.message}，准备覆写`);
      }
    }

    const destDir = path.dirname(destFile);
    try {
      fs.mkdirSync(destDir, { recursive: true, mode: 0o755 });
    } catch (e: any) {
      appendServiceLog('AddonInstaller', `fs.mkdirSync 提示 (${e.code || e.message})，尝试 shell 创建目录`);
      if (process.platform === 'darwin') {
        try { execSync(`mkdir -p "${destDir}"`, { stdio: 'ignore' }); } catch {}
      }
    }

    let writeSucceeded = false;
    // 2. 优先尝试 Node 原生直接写入
    try {
      fs.writeFileSync(destFile, sourceContent, { mode: 0o644 });
      writeSucceeded = true;
      appendServiceLog('AddonInstaller', `fs.writeFileSync 直接写入目标清单成功`);
    } catch (directErr: any) {
      appendServiceLog('AddonInstaller', `fs.writeFileSync 直接写入失败: [${directErr.code || 'UNKNOWN'}] ${directErr.message}`);
    }

    // 3. macOS 平台受限时的系统降级处理
    if (!writeSucceeded && process.platform === 'darwin') {
      appendServiceLog('AddonInstaller', `macOS 下直接写入受限，开始多层提权/系统代理写入流程...`);

      // 3.1 尝试通过系统原生 osascript (AppleScript) 执行写入代理，触发 macOS 授权
      try {
        const tmpFile = path.join(runtimeHome(), 'temp-wps-bridge-manifest.xml');
        fs.writeFileSync(tmpFile, sourceContent, { mode: 0o644 });

        const cmd = `mkdir -p "${destDir}" && cp -f "${tmpFile}" "${destFile}" && chmod 644 "${destFile}"`;
        const appleScript = `do shell script ${JSON.stringify(cmd)}`;
        try {
          execSync(`osascript -e ${JSON.stringify(appleScript)}`, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'pipe'] });
          writeSucceeded = true;
          appendServiceLog('AddonInstaller', `通过 macOS 原生 osascript 代理写入清单成功`);
        } finally {
          try { if (fs.existsSync(tmpFile)) fs.unlinkSync(tmpFile); } catch {}
        }
      } catch (osaErr: any) {
        const osaMsg = osaErr.stderr || osaErr.message || String(osaErr);
        appendServiceLog('AddonInstaller', `osascript 代理写入失败: ${osaMsg}`);

        // 3.2 尝试普通 cp -f 并捕获精确 stderr
        try {
          const tmpFile = path.join(runtimeHome(), 'temp-wps-bridge-manifest.xml');
          fs.writeFileSync(tmpFile, sourceContent, { mode: 0o644 });
          try {
            execSync(`cp -f "${tmpFile}" "${destFile}" && chmod 644 "${destFile}"`, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'pipe'] });
            writeSucceeded = true;
            appendServiceLog('AddonInstaller', `通过 fallback cp -f 写入清单成功`);
          } finally {
            try { if (fs.existsSync(tmpFile)) fs.unlinkSync(tmpFile); } catch {}
          }
        } catch (cpErr: any) {
          const cpMsg = cpErr.stderr || cpErr.message || String(cpErr);
          appendServiceLog('AddonInstaller', `cp -f 命令执行受阻: ${cpMsg}`);
          // 不再只给一句"请去设置里开启"的死路文案：抛带 code 的错误，由统一入口识别为权限问题并走引导。
          const err: any = new Error(`macOS 磁盘访问受限（${cpMsg.trim() || 'EPERM: Operation not permitted'}）`);
          err.code = 'EPERM';
          throw err;
        }
      }
    }

    if (!writeSucceeded && process.platform !== 'darwin') {
      throw new Error(`无法写入 Office 加载项清单文件: ${destFile}`);
    }

    // 4. 平台特有后置处理
    if (process.platform === 'darwin' && fs.existsSync(destFile)) {
      try { execSync(`chmod 644 "${destFile}"`, { stdio: 'ignore' }); } catch {}
    } else if (process.platform === 'win32' && fs.existsSync(destFile)) {
      try {
        appendServiceLog('AddonInstaller', `[Windows] 开始注册 WEF Developer 注册表项`);
        execSync(`reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "55555555-aaaa-bbbb-cccc-777777777777" /t REG_SZ /d "${destFile}" /f`, { stdio: 'ignore' });
        appendServiceLog('AddonInstaller', `[Windows] 注册表写入成功`);
      } catch (regErr: any) {
        appendServiceLog('AddonInstaller', `[Windows] 注册表写入异常: ${regErr.message}`);
      }
      try {
        getOrGenerateCerts();
        const certPath = path.join(runtimeHome(), 'certs/localhost.crt');
        if (fs.existsSync(certPath)) {
          appendServiceLog('AddonInstaller', `[Windows] 导入受信任根证书: ${certPath}`);
          execSync(`certutil -user -addstore "Root" "${certPath}"`, { stdio: 'ignore' });
          appendServiceLog('AddonInstaller', `[Windows] 证书导入完成`);
        }
      } catch (certErr: any) {
        appendServiceLog('AddonInstaller', `[Windows] 证书导入异常: ${certErr.message}`);
      }
    }
  }

  static async checkStatus() {
    appendServiceLog('AddonInstaller', '开始检测 Office 加载项状态');
    try {
      const res = await serviceRequest('/api/v1/office/addon-status');
      appendServiceLog('AddonInstaller', `通过后台服务获取状态成功: installed=${res.installed}, current=${res.current}`);
      return res;
    } catch (e: any) {
      appendServiceLog('AddonInstaller', `后台服务状态请求未响应 (${e.message})，转入本地文件直接检测`);
      const wefDir = this.getWefDirectory();
      const targetFile = path.join(wefDir, 'wps-bridge-manifest.xml');
      const installed = fs.existsSync(targetFile);
      let current = false;
      let ver = '未部署';
      if (installed) {
        try {
          const content = fs.readFileSync(targetFile, 'utf8');
          const m = content.match(/<Version>(.*?)<\/Version>/i);
          ver = m ? m[1].trim() : VERSION;
          current = ver === VERSION || ver === `${VERSION}.0` || ver.startsWith(VERSION);
        } catch {
          current = true;
          ver = VERSION;
        }
      }
      const status = {
        installed,
        current,
        latestVersion: VERSION,
        installedVersion: ver,
        needsUpgrade: !installed || !current,
        targetPath: targetFile,
        platform: process.platform,
        message: !installed ? '尚未部署 Office 官方加载项清单' : !current ? `检测到 Office 加载项需要更新（当前: ${ver}，最新: ${VERSION}）` : `Office 官方加载项已就绪 (v${ver})`
      };
      appendServiceLog('AddonInstaller', `本地检测结果: ${JSON.stringify(status)}`);
      return status;
    }
  }

  static async install() {
    appendServiceLog('AddonInstaller', '触发 Office 加载项部署/升级流程');
    const wefDir = this.getWefDirectory();
    const targetFile = path.join(wefDir, 'wps-bridge-manifest.xml');
    try {
      const source = this.getSourceManifestPath();
      this.safeDeploy(source, targetFile);
      appendServiceLog('AddonInstaller', `Office 加载项部署成功: ${targetFile}`);
      return {
        success: true,
        message: 'Office 官方加载项清单与本地安全配置已就绪！请在 Excel 中重新打开侧边栏。',
        targetPath: targetFile
      };
    } catch (e: any) {
      appendServiceLog('AddonInstaller', `Office 加载项部署失败: ${e.message}`);
      const permissionIssue = detectPermissionIssueCore(e);
      if (permissionIssue) return { success: false, message: 'Office 加载项目录需要 macOS 完全磁盘访问权限', permissionIssue };
      return { success: false, message: e.message };
    }
  }

  static uninstall() {
    appendServiceLog('AddonInstaller', '触发 Office 加载项卸载流程');
    try {
      const wefDir = this.getWefDirectory();
      const target = path.join(wefDir, 'wps-bridge-manifest.xml');
      if (fs.existsSync(target)) {
        try {
          fs.unlinkSync(target);
          appendServiceLog('AddonInstaller', `fs.unlinkSync 删除清单成功: ${target}`);
        } catch (unlinkErr: any) {
          appendServiceLog('AddonInstaller', `fs.unlinkSync 失败 (${unlinkErr.message})，尝试 shell 删除`);
          if (process.platform === 'darwin') {
            try { execSync(`rm -f "${target}"`, { stdio: 'ignore' }); } catch {}
          } else if (process.platform === 'win32') {
            try { execSync('reg delete "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "55555555-aaaa-bbbb-cccc-777777777777" /f', { stdio: 'ignore' }); } catch {}
          }
        }
      }
      return { success: true, message: '已移除 Office 加载项清单，重启 Excel 后生效。' };
    } catch (e: any) {
      appendServiceLog('AddonInstaller', `Office 加载项卸载失败: ${e.message}`);
      const permissionIssue = detectPermissionIssueCore(e);
      if (permissionIssue) return { success: false, message: 'Office 加载项目录需要 macOS 完全磁盘访问权限', permissionIssue };
      return { success: false, message: e.message };
    }
  }
}

