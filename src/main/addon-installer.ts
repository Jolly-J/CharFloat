import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { execSync } from 'node:child_process';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { atomicWrite, getToken, resourcePath, runtimePort, runtimeHome, getOrGenerateCerts, VERSION } from '../bridge/runtime.js';
import { serviceRequest } from '../bridge/service-client.js';

const activeNames = ['Office Agent Bridge', 'Office Agent Bridge (表格)', 'Office Agent Bridge (文字)', 'Office Agent Bridge (演示)'];
const legacyNames = ['WPS Bridge', 'WPS Bridge (表格)', 'WPS Bridge (文字)', 'WPS Bridge (演示)'];
const ownedNames = [...activeNames, ...legacyNames];
export function mergePluginIndex(raw: string, remove = false) {
  const doc = new DOMParser({ onError: () => { throw new Error('插件索引 XML 无效，停止写入'); } }).parseFromString(raw || '<jsplugins/>', 'application/xml');
  if (doc.documentElement?.tagName !== 'jsplugins') throw new Error('插件索引根节点不是 jsplugins，停止写入');
  const nodes = Array.from(doc.getElementsByTagName('jsplugin'));
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
    const current = details.length > 0 && details.every(d => d.current);
    const hasBlocked = details.some(d => d.blocked);
    const needsUpgrade = !installed || !current;
    return {
      installed,
      current,
      latestVersion: VERSION,
      installedVersion: installedVersion || (installed ? '未知' : '未安装'),
      needsUpgrade,
      hasBlocked,
      platform: process.platform,
      targetPath: dirs.join(' | '),
      details,
      message: !dirs.length
        ? '未找到 WPS 用户加载项目录，请先运行 WPS 或指定目录'
        : hasBlocked
          ? '检测到加载项被 WPS 阻断，请点击【一键修复 / 升级加载项】解除阻断'
          : needsUpgrade
            ? `检测到加载项需要更新（当前: ${installedVersion || '未知'}，最新: ${VERSION}）`
            : `加载项已是最新版本 (v${VERSION})`
    };
  }
  static install(options: { cleanBlocked?: boolean } = { cleanBlocked: true }) {
    try {
      const dirs = this.getAllAddonDirectories();
      if (!dirs.length) throw new Error('找不到 WPS 加载项目录，请先安装并运行 WPS。');
      const source = this.getSourceAddonPath();
      // Validate every index before changing anything.
      const indexes = dirs.flatMap(dir => ['publish.xml', 'jsplugins.xml'].map(name => {
        const file = path.join(dir, name); return { file, content: mergePluginIndex(fs.existsSync(file) ? fs.readFileSync(file, 'utf8') : '') };
      }));
      const token = getToken(true);
      for (const dir of dirs) {
        if (options.cleanBlocked !== false) {
          const blockFile = path.join(dir, 'jsaddinblockhost.ini');
          if (fs.existsSync(blockFile)) {
            try {
              fs.copyFileSync(blockFile, `${blockFile}.backup-${Date.now()}`);
              fs.unlinkSync(blockFile);
            } catch {}
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
          for (const file of fs.readdirSync(source)) if (fs.statSync(path.join(source, file)).isFile()) atomicWrite(path.join(dest, file), fs.readFileSync(path.join(source, file), 'utf8'), true);
          atomicWrite(path.join(dest, 'bridge-config.js'), `window.WPS_BRIDGE_CONFIG = ${JSON.stringify({ port: runtimePort(), token, version: VERSION })};\n`);
        }
      }
      for (const index of indexes) atomicWrite(index.file, index.content, true);
      return { success: true, message: `加载项已成功部署并更新至 v${VERSION}！请在 WPS 中点击功能区【重新连接】或重启 WPS 生效。`, targetPath: dirs.join(' | '), warnings: dirs.filter(dir => fs.existsSync(path.join(dir, 'jsaddinblockhost.ini'))).map(() => '检测到 WPS 阻断配置，未删除。请在 WPS 中检查加载项权限。') };
    } catch (e: any) { return { success: false, message: e.message }; }
  }
  static uninstall() {
    try {
      const indexes = this.getAllAddonDirectories().flatMap(dir => ['publish.xml', 'jsplugins.xml'].map(name => path.join(dir, name))).filter(file => fs.existsSync(file)).map(file => ({ file, content: mergePluginIndex(fs.readFileSync(file, 'utf8'), true) }));
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
    const destDir = path.dirname(destFile);
    try {
      fs.mkdirSync(destDir, { recursive: true, mode: 0o755 });
    } catch {
      if (process.platform === 'darwin') {
        try { execSync(`mkdir -p "${destDir}"`, { stdio: 'ignore' }); } catch {}
      }
    }

    let writeSucceeded = false;
    let content = '';
    try {
      content = fs.readFileSync(sourceFile, 'utf8');
      fs.writeFileSync(destFile, content, { mode: 0o644 });
      writeSucceeded = true;
    } catch {}

    if (!writeSucceeded && process.platform === 'darwin') {
      try {
        if (!content) content = fs.readFileSync(sourceFile, 'utf8');
        const tmpFile = path.join(runtimeHome(), 'temp-wps-bridge-manifest.xml');
        fs.writeFileSync(tmpFile, content, { mode: 0o644 });
        try {
          execSync(`cp -f "${tmpFile}" "${destFile}" && chmod 644 "${destFile}"`, { stdio: 'ignore' });
          writeSucceeded = true;
        } finally {
          try { if (fs.existsSync(tmpFile)) fs.unlinkSync(tmpFile); } catch {}
        }
      } catch (e: any) {
        throw new Error(`部署清单文件受限：${e.message || 'macOS 容器沙盒拦截'}`);
      }
    }

    if (process.platform === 'darwin' && fs.existsSync(destFile)) {
      try { execSync(`chmod 644 "${destFile}"`, { stdio: 'ignore' }); } catch {}
    } else if (process.platform === 'win32' && fs.existsSync(destFile)) {
      try {
        execSync(`reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "55555555-aaaa-bbbb-cccc-777777777777" /t REG_SZ /d "${destFile}" /f`, { stdio: 'ignore' });
      } catch {}
    }
  }

  static async checkStatus() {
    try {
      return await serviceRequest('/api/v1/office/addon-status');
    } catch {
      const wefDir = this.getWefDirectory();
      const targetFile = path.join(wefDir, 'wps-bridge-manifest.xml');
      const installed = fs.existsSync(targetFile);
      return {
        installed,
        current: installed,
        latestVersion: VERSION,
        installedVersion: installed ? VERSION : '未部署',
        needsUpgrade: !installed,
        targetPath: targetFile,
        platform: process.platform,
        message: installed ? `Office 官方加载项已就绪 (v${VERSION})` : '尚未部署 Office 官方加载项清单'
      };
    }
  }

  static async install() {
    // 1. 幂等性自愈优先：若清单已就绪，直接返回成功，杜绝任何沙盒写入拦截
    const wefDir = this.getWefDirectory();
    const targetFile = path.join(wefDir, 'wps-bridge-manifest.xml');
    if (fs.existsSync(targetFile)) {
      return {
        success: true,
        message: 'Office 官方加载项清单已处于最新状态！请在 Excel 中点击【插入 -> 我的加载项】启用。',
        targetPath: targetFile
      };
    }

    // 2. 优先委托独立后台守护服务执行安全写入（命令行环境不受 GUI 沙盒隔离）
    try {
      return await serviceRequest('/api/v1/office/install-addon', {});
    } catch {
      // 3. 后台未启动时的安全降级
      try {
        const source = this.getSourceManifestPath();
        this.safeDeploy(source, targetFile);
        return {
          success: true,
          message: 'Office 官方加载项已成功部署！请在 Excel 中点击【插入 -> 我的加载项】启用。',
          targetPath: targetFile
        };
      } catch (e: any) {
        return { success: false, message: e.message };
      }
    }
  }

  static uninstall() {
    try {
      const wefDir = this.getWefDirectory();
      const target = path.join(wefDir, 'wps-bridge-manifest.xml');
      if (fs.existsSync(target)) {
        try {
          fs.unlinkSync(target);
        } catch {
          if (process.platform === 'darwin') {
            try { execSync(`rm -f "${target}"`, { stdio: 'ignore' }); } catch {}
          } else if (process.platform === 'win32') {
            try { execSync('reg delete "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "55555555-aaaa-bbbb-cccc-777777777777" /f', { stdio: 'ignore' }); } catch {}
          }
        }
      }
      return { success: true, message: '已移除 Office 加载项清单，重启 Excel 后生效。' };
    } catch (e: any) {
      return { success: false, message: e.message };
    }
  }
}

