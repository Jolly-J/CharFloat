// src/main/index.ts
import { app as app3, BrowserWindow as BrowserWindow2, ipcMain, Tray, Menu, nativeImage as nativeImage2, clipboard, nativeTheme as nativeTheme2, shell } from "electron";
import path7 from "path";
import fs6 from "fs";
import { fileURLToPath as fileURLToPath2 } from "url";

// src/bridge/service-client.ts
import fs2 from "fs";
import path2 from "path";
import { spawn } from "child_process";

// src/bridge/runtime.ts
import fs from "fs";
import path from "path";
import os from "os";
import crypto from "crypto";
import { execSync } from "child_process";
var VERSION = "2.1.0";
var PROTOCOL = 2;
function runtimeHome() {
  return process.env.WPS_BRIDGE_HOME || (process.platform === "win32" ? path.join(process.env.LOCALAPPDATA || os.homedir(), "WPSBridge") : path.join(os.homedir(), ".wps-bridge"));
}
function runtimePort() {
  const port = Number(process.env.WPS_BRIDGE_PORT || 19890);
  if (!Number.isInteger(port) || port < 1024 || port > 65535) throw new Error("WPS_BRIDGE_PORT \u5FC5\u987B\u4E3A 1024\u201365535");
  return port;
}
function resourcePath(relative) {
  const entryDir = path.dirname(path.resolve(process.argv[1] || "."));
  const roots = [process.env.WPS_BRIDGE_RESOURCES, path.resolve(entryDir, "../.."), process.cwd()].filter(Boolean);
  const found = roots.flatMap((root2) => {
    const p = path.join(root2, relative);
    return [p.replace(/app\.asar([\\/])/, "app.asar.unpacked$1"), p];
  }).find((p) => fs.existsSync(p));
  if (!found) throw new Error(`\u7F3A\u5C11\u8FD0\u884C\u8D44\u6E90 ${relative}\u3002\u8BF7\u91CD\u65B0\u6784\u5EFA\u6216\u4FEE\u590D Bridge \u5B89\u88C5\u3002`);
  return found;
}
function atomicWrite(file, data, backup = false) {
  fs.mkdirSync(path.dirname(file), { recursive: true, mode: 448 });
  if (fs.existsSync(file) && fs.readFileSync(file, "utf8") === data) return;
  if (backup && fs.existsSync(file)) fs.copyFileSync(file, `${file}.backup-${Date.now()}-${crypto.randomBytes(3).toString("hex")}`);
  const tmp = `${file}.${crypto.randomBytes(6).toString("hex")}.tmp`;
  try {
    fs.writeFileSync(tmp, data, { mode: 384, flag: "wx" });
    fs.renameSync(tmp, file);
  } finally {
    if (fs.existsSync(tmp)) fs.unlinkSync(tmp);
  }
}
function getToken(create = false) {
  const file = path.join(runtimeHome(), "token");
  if (create) {
    fs.mkdirSync(runtimeHome(), { recursive: true, mode: 448 });
    try {
      fs.writeFileSync(file, crypto.randomBytes(32).toString("hex"), { flag: "wx", mode: 384 });
    } catch (error) {
      if (error.code !== "EEXIST") throw error;
    }
  }
  return fs.readFileSync(file, "utf8").trim();
}
function getOrGenerateCerts() {
  const certDir = path.join(runtimeHome(), "certs");
  fs.mkdirSync(certDir, { recursive: true, mode: 448 });
  const keyPath = path.join(certDir, "localhost.key");
  const certPath = path.join(certDir, "localhost.crt");
  if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
    return {
      key: fs.readFileSync(keyPath, "utf8"),
      cert: fs.readFileSync(certPath, "utf8")
    };
  }
  try {
    const prebuiltKey = resourcePath("resources/certs/localhost.key");
    const prebuiltCert = resourcePath("resources/certs/localhost.crt");
    if (fs.existsSync(prebuiltKey) && fs.existsSync(prebuiltCert)) {
      fs.copyFileSync(prebuiltKey, keyPath);
      fs.copyFileSync(prebuiltCert, certPath);
      return {
        key: fs.readFileSync(keyPath, "utf8"),
        cert: fs.readFileSync(certPath, "utf8")
      };
    }
  } catch {
  }
  try {
    execSync(`openssl req -x509 -newkey rsa:2048 -keyout "${keyPath}" -out "${certPath}" -days 3650 -nodes -subj "/CN=localhost" -addext "subjectAltName=DNS:localhost,IP:127.0.0.1"`, { stdio: "ignore" });
    if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
      return {
        key: fs.readFileSync(keyPath, "utf8"),
        cert: fs.readFileSync(certPath, "utf8")
      };
    }
  } catch {
  }
  return null;
}
function appendServiceLog(tag, message) {
  try {
    const home = runtimeHome();
    fs.mkdirSync(home, { recursive: true, mode: 448 });
    const logPath = path.join(home, "service.log");
    const time = (/* @__PURE__ */ new Date()).toISOString();
    fs.appendFileSync(logPath, `[${time}] [${tag}] ${message}
`, "utf8");
  } catch {
  }
}

// src/bridge/service-client.ts
var SERVICE_NAME = "wps-bridge";
function serviceAddress() {
  return `http://127.0.0.1:${runtimePort()}`;
}
function serviceRecovery() {
  return [
    "\u786E\u8BA4 Bridge \u540E\u53F0\u6B63\u5728\u8FD0\u884C\uFF1A\u6267\u884C `npx office-agent-bridge --status` \u67E5\u770B ready \u5B57\u6BB5",
    "\u672A\u8FD0\u884C\u65F6\u542F\u52A8\uFF1A`npx office-agent-bridge --start`\uFF08stdio \u5BA2\u6237\u7AEF\u9996\u6B21\u8FDE\u63A5\u4E5F\u4F1A\u81EA\u52A8\u542F\u52A8\uFF09",
    `\u4ECD\u5931\u8D25\u65F6\u67E5\u770B\u65E5\u5FD7 ${path2.join(runtimeHome(), "service.log")}\uFF0C\u5E76\u7528 \`npx office-agent-bridge --doctor\` \u590D\u6838\u51ED\u636E\u4E0E\u7AEF\u53E3`,
    "\u4E0D\u8981\u53CD\u590D\u542F\u52A8\u591A\u4E2A\u5B9E\u4F8B\uFF1B\u7AEF\u53E3\u88AB\u65E7\u7248\u5360\u7528\u65F6\u5148 `npx office-agent-bridge --stop`"
  ];
}
function describeUnreachable(error) {
  const raw = error?.message || String(error);
  const detail = error?.cause?.code ? `${raw}\uFF08${error.cause.code}\uFF09` : raw;
  return `${SERVICE_NAME} \u540E\u53F0\u670D\u52A1\u672A\u8FD0\u884C\u6216\u4E0D\u53EF\u8FBE\uFF08\u5730\u5740 ${serviceAddress()}\uFF09\u3002\u6062\u590D\u6B65\u9AA4\uFF1A${serviceRecovery().join("\uFF1B")}\u3002\u539F\u59CB\u9519\u8BEF\uFF1A${detail}`;
}
async function serviceRequest(route, body) {
  let response;
  try {
    response = await fetch(`${serviceAddress()}${route}`, {
      method: body === void 0 ? "GET" : "POST",
      headers: { Authorization: `Bearer ${getToken()}`, "Content-Type": "application/json" },
      body: body === void 0 ? void 0 : JSON.stringify(body),
      signal: AbortSignal.timeout(35e3)
    });
  } catch (error) {
    throw new Error(describeUnreachable(error));
  }
  const result = await response.json().catch(() => ({}));
  if (!response.ok || result.success === false) throw new Error(result.error || `${SERVICE_NAME} HTTP ${response.status}\uFF08\u5730\u5740 ${serviceAddress()}\uFF09`);
  return result;
}
async function probeService() {
  try {
    const r = await fetch(`${serviceAddress()}/health`, { signal: AbortSignal.timeout(1200) });
    if (!r.ok) return { ready: false, occupied: true, service: SERVICE_NAME, address: serviceAddress(), recovery: serviceRecovery(), message: "\u7AEF\u53E3\u7531\u65E7\u7248 Bridge \u6216\u5176\u4ED6\u7A0B\u5E8F\u5360\u7528\uFF1B\u8BF7\u5148\u9000\u51FA\u65E7\u7248 Bridge\u3002" };
    const info = await r.json();
    if (info.service !== "wps-bridge" || info.protocol !== PROTOCOL) return { ready: false, occupied: true, service: SERVICE_NAME, address: serviceAddress(), recovery: serviceRecovery(), message: "\u7AEF\u53E3\u4E0A\u7684\u670D\u52A1\u7248\u672C\u4E0D\u517C\u5BB9\u3002" };
    await serviceRequest("/api/v1/status");
    return { ready: true, occupied: true, service: SERVICE_NAME, address: serviceAddress(), info };
  } catch (e) {
    const raw = e?.message || String(e);
    if (String(raw).includes("401")) return { ready: false, occupied: true, service: SERVICE_NAME, address: serviceAddress(), recovery: serviceRecovery(), message: "Bridge \u51ED\u636E\u4E0D\u5339\u914D\uFF0C\u8BF7\u68C0\u67E5\u8FD0\u884C\u76EE\u5F55\u3002", error: raw };
    return {
      ready: false,
      occupied: false,
      service: SERVICE_NAME,
      address: serviceAddress(),
      message: `${SERVICE_NAME} \u540E\u53F0\u670D\u52A1\u672A\u8FD0\u884C\uFF08\u5730\u5740 ${serviceAddress()}\uFF09\u3002`,
      recovery: serviceRecovery(),
      error: raw
    };
  }
}
var starting;
function ensureService(entry) {
  if (starting) return starting;
  starting = (async () => {
    const state2 = await probeService();
    if (state2.ready) return;
    if (state2.occupied) throw new Error(state2.message);
    getToken(true);
    const logPath = path2.join(runtimeHome(), "service.log");
    if (fs2.existsSync(logPath) && fs2.statSync(logPath).size > 2 * 1024 * 1024) fs2.renameSync(logPath, `${logPath}.previous`);
    const log = fs2.openSync(logPath, "a", 384);
    const child = spawn(process.execPath, [entry, "--serve"], {
      detached: true,
      windowsHide: true,
      stdio: ["ignore", log, log],
      env: { ...process.env, ELECTRON_RUN_AS_NODE: "1" }
    });
    fs2.closeSync(log);
    let launchError;
    child.on("error", (error) => {
      launchError = error;
    });
    child.unref();
    for (let n = 0; n < 40; n++) {
      await new Promise((r) => setTimeout(r, 150));
      if (launchError) throw launchError;
      if ((await probeService()).ready) return;
    }
    throw new Error(`\u540E\u53F0\u542F\u52A8\u5931\u8D25\u3002\u67E5\u770B ${logPath}\uFF0C\u4E0D\u8981\u53CD\u590D\u542F\u52A8\u591A\u4E2A\u5B9E\u4F8B\u3002`);
  })().finally(() => {
    starting = void 0;
  });
  return starting;
}

// src/main/permissions.ts
import { app } from "electron";
import path3 from "path";
var FULL_DISK_ACCESS_URL = "x-apple.systempreferences:com.apple.preference.security?Privacy_AllFiles";
function appToAuthorize() {
  const exe = process.execPath;
  const appIndex = exe.indexOf(".app/");
  if (appIndex !== -1) return exe.slice(0, appIndex + 4);
  return app.isPackaged ? path3.dirname(exe) : exe;
}

// src/main/permission-window.ts
import { BrowserWindow, app as app2, nativeImage, nativeTheme } from "electron";
import path4 from "path";
import fs3 from "fs";
import { fileURLToPath } from "url";
var guideWindow = null;
var pendingIssue = null;
var here = path4.dirname(fileURLToPath(import.meta.url));
function getPermissionIssue() {
  return pendingIssue;
}
function openPermissionWindow(issue) {
  pendingIssue = issue;
  if (guideWindow && !guideWindow.isDestroyed()) {
    guideWindow.webContents.send("permission:issue", issue);
    guideWindow.show();
    guideWindow.focus();
    return;
  }
  guideWindow = new BrowserWindow({
    width: 430,
    height: 560,
    resizable: false,
    minimizable: false,
    maximizable: false,
    fullscreenable: false,
    alwaysOnTop: true,
    title: "\u9700\u8981\u5B8C\u5168\u78C1\u76D8\u8BBF\u95EE\u6743\u9650",
    backgroundColor: nativeTheme.shouldUseDarkColors ? "#17191c" : "#ffffff",
    webPreferences: { preload: path4.join(here, "../preload/index.cjs"), contextIsolation: true, nodeIntegration: false, sandbox: true }
  });
  guideWindow.setMenuBarVisibility(false);
  guideWindow.setAlwaysOnTop(true, "floating");
  if (process.env.VITE_DEV_SERVER_URL) void guideWindow.loadURL(`${process.env.VITE_DEV_SERVER_URL}#permission-guide`);
  else void guideWindow.loadFile(path4.join(here, "../renderer/index.html"), { hash: "permission-guide" });
  guideWindow.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
  guideWindow.webContents.on("will-navigate", (event) => event.preventDefault());
  guideWindow.webContents.once("did-finish-load", () => guideWindow?.webContents.send("permission:issue", pendingIssue));
  guideWindow.webContents.on("preload-error", (_e, preloadPath, error) => {
    appendServiceLog("PermissionWindow", `\u9884\u52A0\u8F7D\u811A\u672C\u5931\u8D25: ${preloadPath} \u2014 ${error?.message || error}`);
  });
  guideWindow.on("closed", () => {
    guideWindow = null;
  });
}
function closePermissionWindow() {
  if (guideWindow && !guideWindow.isDestroyed()) guideWindow.close();
  guideWindow = null;
}
var ICON_RELATIVE = "resources/icon.png";
function loadAppIcon() {
  try {
    const iconPath = resourcePath(ICON_RELATIVE);
    if (!fs3.existsSync(iconPath)) return "";
    return `data:image/png;base64,${fs3.readFileSync(iconPath).toString("base64")}`;
  } catch (error) {
    appendServiceLog("PermissionWindow", `\u8BFB\u53D6\u56FE\u6807\u5931\u8D25: ${error?.message || error}`);
    return "";
  }
}
function startAppDrag(sender) {
  const file = appToAuthorize();
  const payload = { file };
  try {
    const icon = nativeImage.createFromPath(resourcePath(ICON_RELATIVE));
    if (!icon.isEmpty()) payload.icon = icon.resize({ width: 128, height: 128 });
  } catch {
  }
  sender.startDrag(payload);
}

// src/main/installer-engine.ts
import fs5 from "fs";
import path6 from "path";
import os3 from "os";

// src/main/addon-installer.ts
import fs4 from "fs";
import path5 from "path";
import os2 from "os";
import { execSync as execSync2 } from "child_process";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

// src/main/permission-detect.ts
function extractProtectedPath(message) {
  const quoted = message.match(/['"]([^'"]*\/Library\/[^'"]*)['"]/);
  if (quoted) return quoted[1];
  const bare = message.match(/(\/Users\/[^\s'"]+|\/Library\/[^\s'"]+)/);
  return bare ? bare[1].replace(/[.,:;)\]、。，；）]+$/, "") : "";
}
function detectPermissionIssueCore(error, platform = process.platform) {
  if (platform !== "darwin") return null;
  const code = error?.code;
  const detail = error instanceof Error ? error.message : String(error ?? "");
  const isPermission = code === "EPERM" || code === "EACCES" || /Operation not permitted|not permitted|Permission denied|EPERM|EACCES/i.test(detail);
  if (!isPermission) return null;
  return { kind: "macos-full-disk-access", targetPath: extractProtectedPath(detail), detail };
}

// src/main/addon-installer.ts
var activeNames = ["Office Agent Bridge", "Office Agent Bridge (\u8868\u683C)", "Office Agent Bridge (\u6587\u5B57)", "Office Agent Bridge (\u6F14\u793A)"];
var legacyNames = ["WPS Bridge", "WPS Bridge (\u8868\u683C)", "WPS Bridge (\u6587\u5B57)", "WPS Bridge (\u6F14\u793A)"];
var ownedNames = [...activeNames, ...legacyNames];
function readXmlSafe(file) {
  if (!fs4.existsSync(file)) return "";
  const buf = fs4.readFileSync(file);
  if (!buf.length) return "";
  appendServiceLog("AddonInstaller", `\u8BFB\u53D6 ${file}: ${buf.length} \u5B57\u8282, \u524D\u5BFC\u5341\u516D\u8FDB\u5236: ${buf.subarray(0, 16).toString("hex")}`);
  if (buf.length >= 2) {
    if (buf[0] === 255 && buf[1] === 254) return new TextDecoder("utf-16le").decode(buf.subarray(2));
    if (buf[0] === 254 && buf[1] === 255) return new TextDecoder("utf-16be").decode(buf.subarray(2));
    if (buf[0] === 60 && buf[1] === 0) return new TextDecoder("utf-16le").decode(buf);
    if (buf[0] === 0 && buf[1] === 60) return new TextDecoder("utf-16be").decode(buf);
  }
  if (buf.length >= 3 && buf[0] === 239 && buf[1] === 187 && buf[2] === 191) {
    return new TextDecoder("utf-8").decode(buf.subarray(3));
  }
  const head = buf.subarray(0, 120).toString("ascii").toLowerCase();
  if (head.includes('encoding="utf-16"') || head.includes("encoding='utf-16'")) {
    try {
      return new TextDecoder("utf-16le").decode(buf);
    } catch {
    }
  }
  if (head.includes('encoding="gbk"') || head.includes("encoding='gbk'") || head.includes('encoding="gb2312"') || head.includes("encoding='gb2312'")) {
    try {
      return new TextDecoder("gbk").decode(buf);
    } catch {
    }
  }
  const utf8 = new TextDecoder("utf-8").decode(buf);
  if (utf8.includes("\uFFFD") || utf8.includes("\0")) {
    try {
      const gbk = new TextDecoder("gbk").decode(buf);
      if (!gbk.includes("\uFFFD") && !gbk.includes("\0")) return gbk;
    } catch {
    }
  }
  return utf8.replace(/\0/g, "");
}
function mergePluginIndex(raw, remove = false) {
  let text = (raw || "").replace(/^\uFEFF/, "").replace(/\0/g, "").trim();
  const body = text.replace(/<\?xml[\s\S]*?\?>/gi, "").replace(/<!--[\s\S]*?-->/gi, "").trim();
  if (!body || !body.includes("<")) {
    text = "<jsplugins/>";
  }
  let doc;
  try {
    doc = new DOMParser({
      onError: (level, msg) => {
        if (level === "fatalError" && !msg.includes("missing root element")) {
          throw new Error(`\u63D2\u4EF6\u7D22\u5F15 XML \u65E0\u6548\uFF0C\u505C\u6B62\u5199\u5165: ${msg}`);
        }
      }
    }).parseFromString(text, "application/xml");
  } catch (e) {
    if (String(e.message || e).includes("missing root element")) {
      appendServiceLog("AddonInstaller", "\u68C0\u6D4B\u5230\u7D22\u5F15 XML \u7F3A\u5C11\u6839\u8282\u70B9 (missing root element)\uFF0C\u81EA\u52A8\u5B89\u5168\u521D\u59CB\u5316\u4E3A <jsplugins/>");
      doc = new DOMParser().parseFromString("<jsplugins/>", "application/xml");
    } else {
      appendServiceLog("AddonInstaller", `\u63D2\u4EF6\u7D22\u5F15 XML \u8BED\u6CD5\u9519\u8BEF\u62E6\u622A: ${e.message}`);
      throw e;
    }
  }
  if (!doc || !doc.documentElement) {
    appendServiceLog("AddonInstaller", "\u89E3\u6790\u540E\u672A\u751F\u6210 documentElement\uFF0C\u81EA\u52A8\u5B89\u5168\u521D\u59CB\u5316\u4E3A <jsplugins/>");
    doc = new DOMParser().parseFromString("<jsplugins/>", "application/xml");
  } else if (doc.documentElement.tagName !== "jsplugins") {
    appendServiceLog("AddonInstaller", `\u63D2\u4EF6\u7D22\u5F15\u6839\u8282\u70B9\u4E0D\u662F jsplugins (\u5F53\u524D\u4E3A: <${doc.documentElement.tagName}>)\uFF0C\u505C\u6B62\u5199\u5165`);
    throw new Error(`\u63D2\u4EF6\u7D22\u5F15\u6839\u8282\u70B9\u4E0D\u662F jsplugins\uFF0C\u505C\u6B62\u5199\u5165`);
  }
  const nodes = Array.from(doc.getElementsByTagName("jsplugin"));
  for (const node of nodes) if (ownedNames.includes(node.getAttribute("name") || "")) node.parentNode.removeChild(node);
  if (!remove) for (const [i, type] of ["et", "wps", "wpp"].entries()) {
    const node = doc.createElement("jsplugin");
    for (const [key, value] of Object.entries({ name: activeNames[i + 1], type, url: "./wps-bridge", enable: "true", autoload: "true", version: VERSION })) node.setAttribute(key, value);
    doc.documentElement.appendChild(node);
  }
  return new XMLSerializer().serializeToString(doc);
}
function isHostBlocked(dir) {
  const blockFile = path5.join(dir, "jsaddinblockhost.ini");
  if (!fs4.existsSync(blockFile)) return false;
  try {
    const content = fs4.readFileSync(blockFile, "utf8");
    return /127\.0\.0\.1|localhost|19890|19891|wps-bridge|office-agent-bridge/i.test(content);
  } catch {
    return false;
  }
}
var AddonInstaller = class {
  static getAllAddonDirectories() {
    if (process.env.WPS_BRIDGE_ADDON_DIR) return [process.env.WPS_BRIDGE_ADDON_DIR];
    const home = os2.homedir();
    const candidates = process.platform === "win32" ? [path5.join(process.env.APPDATA || path5.join(home, "AppData/Roaming"), "kingsoft/wps/jsaddons")] : process.platform === "darwin" ? [
      "Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons",
      "Library/Containers/cn.wps.moffice_mac/Data/.kingsoft/wps/jsaddons",
      "Library/Containers/com.kingsoft.wpsoffice.mac/Data/.local/share/Kingsoft/wps/jsaddons",
      "Library/Containers/cn.wps.moffice_mac/Data/.local/share/Kingsoft/wps/jsaddons",
      "Library/Application Support/Kingsoft/wps/jsaddons"
    ].map((p) => path5.join(home, p)) : [];
    return candidates.filter((p) => fs4.existsSync(p) || fs4.existsSync(path5.dirname(p)));
  }
  static getAddonDirectory() {
    return this.getAllAddonDirectories()[0] || null;
  }
  static getSourceAddonPath() {
    return resourcePath("wps-addon");
  }
  /**
   * 读取加载项产物头部注入的构建指纹。
   *
   * 为什么需要它：`manifest.xml` 的版本号（2.1.0）在多次构建之间**不变**，
   * 所以"版本号相同"完全不能说明"跑的是同一份构建"。指纹每次构建都变，
   * 才是判断"已部署的是不是包内这一份"的可靠依据。
   */
  static readAddonFingerprint(file) {
    if (!file) return null;
    try {
      const fd = fs4.openSync(file, "r");
      try {
        const buf = Buffer.alloc(4096);
        const read = fs4.readSync(fd, buf, 0, buf.length, 0);
        const m = buf.toString("utf8", 0, read).match(/ADDON_BUILD_FINGERPRINT:\s*([0-9a-f]{16,64})/);
        return m ? m[1] : null;
      } finally {
        fs4.closeSync(fd);
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
    const bundledPath = path5.join(this.getSourceAddonPath(), "addon-core.js");
    const bundled = this.readAddonFingerprint(fs4.existsSync(bundledPath) ? bundledPath : null);
    const deployed = this.getAllAddonDirectories().map((dir) => {
      const file = path5.join(dir, "wps-bridge", "addon-core.js");
      return { dir, file, exists: fs4.existsSync(file), fingerprint: this.readAddonFingerprint(fs4.existsSync(file) ? file : null) };
    }).filter((d) => d.exists);
    if (bundled === null) {
      return {
        bundledFingerprint: null,
        bundledPath,
        deployed,
        stale: null,
        reason: "\u8BFB\u4E0D\u5230\u5BA2\u6237\u7AEF\u5305\u5185\u7684\u6784\u5EFA\u6307\u7EB9\uFF08\u4EA7\u7269\u7F3A\u5931\u6216\u672A\u6CE8\u5165\uFF09\uFF0C\u65E0\u6CD5\u5224\u65AD\u662F\u5426\u9700\u8981\u5347\u7EA7"
      };
    }
    if (deployed.length === 0) {
      return {
        bundledFingerprint: bundled,
        bundledPath,
        deployed,
        stale: true,
        reason: "\u672A\u627E\u5230\u5DF2\u90E8\u7F72\u7684\u52A0\u8F7D\u9879\u526F\u672C\uFF0C\u9700\u8981\u5B89\u88C5"
      };
    }
    const mismatched = deployed.filter((d) => d.fingerprint !== bundled);
    return {
      bundledFingerprint: bundled,
      bundledPath,
      deployed,
      stale: mismatched.length > 0,
      mismatchedCount: mismatched.length,
      reason: mismatched.length > 0 ? `\u5DF2\u90E8\u7F72\u7684\u52A0\u8F7D\u9879\u4E0E\u5BA2\u6237\u7AEF\u5305\u5185\u7684\u6784\u5EFA\u4E0D\u4E00\u81F4\uFF08${mismatched.length}/${deployed.length} \u4E2A\u526F\u672C\uFF09\uFF1A\u5305\u5185 ${bundled.slice(0, 12)}\u2026\uFF0C\u90E8\u7F72 ${(mismatched[0].fingerprint || "\u672A\u77E5").slice(0, 12)}\u2026` : "\u5DF2\u90E8\u7F72\u7684\u52A0\u8F7D\u9879\u4E0E\u5BA2\u6237\u7AEF\u5305\u5185\u7684\u6784\u5EFA\u4E00\u81F4"
    };
  }
  static checkStatus() {
    const dirs = this.getAllAddonDirectories();
    let installedVersion = "";
    const details = dirs.map((dir) => {
      const manifest = path5.join(dir, "wps-bridge/manifest.xml");
      const config = path5.join(dir, "wps-bridge/bridge-config.js");
      const installed2 = fs4.existsSync(manifest);
      let current2 = false, error = "", ver = "";
      try {
        if (installed2) {
          const content = fs4.readFileSync(manifest, "utf8");
          const m = content.match(/<version>(.*?)<\/version>/);
          ver = m ? m[1].trim() : "";
          if (!installedVersion && ver) installedVersion = ver;
          current2 = ver === VERSION && fs4.existsSync(config);
        }
      } catch (e) {
        error = `\u65E0\u6CD5\u8BFB\u53D6\u52A0\u8F7D\u9879\u6587\u4EF6\uFF1A${e.code || e.message}`;
      }
      return { dir, installed: installed2, current: current2, version: ver, error, blocked: isHostBlocked(dir) };
    });
    const installed = details.some((d) => d.installed);
    const current = installed && (installedVersion === VERSION || details.some((d) => d.installed && d.current));
    const hasBlocked = details.some((d) => d.blocked);
    const buildFreshness = this.checkBuildFreshness();
    const buildStale = buildFreshness.stale === true;
    const needsUpgrade = !installed || !current || buildStale;
    return {
      installed,
      current,
      latestVersion: VERSION,
      installedVersion: installedVersion || (installed ? "\u672A\u77E5" : "\u672A\u5B89\u88C5"),
      needsUpgrade,
      hasBlocked,
      buildFreshness,
      platform: process.platform,
      targetPath: dirs.join(" | "),
      details,
      message: !dirs.length ? "\u672A\u627E\u5230 WPS \u7528\u6237\u52A0\u8F7D\u9879\u76EE\u5F55\uFF0C\u8BF7\u5148\u8FD0\u884C WPS \u6216\u6307\u5B9A\u76EE\u5F55" : hasBlocked ? "\u68C0\u6D4B\u5230\u52A0\u8F7D\u9879\u88AB WPS \u963B\u65AD\uFF0C\u8BF7\u70B9\u51FB\u3010\u4E00\u952E\u4FEE\u590D / \u5347\u7EA7\u52A0\u8F7D\u9879\u3011\u89E3\u9664\u963B\u65AD" : needsUpgrade ? buildStale ? `\u68C0\u6D4B\u5230\u52A0\u8F7D\u9879\u6784\u5EFA\u4E0D\u662F\u5305\u5185\u8FD9\u4E00\u4EFD\uFF08\u7248\u672C\u53F7\u540C\u4E3A ${VERSION} \u4F46\u6784\u5EFA\u6307\u7EB9\u4E0D\u540C\uFF09\uFF0C\u9700\u8981\u91CD\u65B0\u90E8\u7F72` : `\u68C0\u6D4B\u5230\u52A0\u8F7D\u9879\u9700\u8981\u66F4\u65B0\uFF08\u5F53\u524D: ${installedVersion || "\u672A\u77E5"}\uFF0C\u6700\u65B0: ${VERSION}\uFF09` : `\u52A0\u8F7D\u9879\u5DF2\u662F\u6700\u65B0\u7248\u672C (v${VERSION}\uFF0C\u6784\u5EFA\u6307\u7EB9\u4E00\u81F4)`
    };
  }
  static install(options = { cleanBlocked: true }) {
    appendServiceLog("AddonInstaller", "=== \u5F00\u59CB\u6267\u884C WPS \u52A0\u8F7D\u9879\u5B89\u88C5 / \u5347\u7EA7 ===");
    try {
      const dirs = this.getAllAddonDirectories();
      appendServiceLog("AddonInstaller", `\u68C0\u6D4B\u5230 WPS \u76EE\u6807\u52A0\u8F7D\u9879\u76EE\u5F55: ${dirs.length ? dirs.join(" | ") : "\u672A\u627E\u5230\u4EFB\u4F55\u76EE\u5F55"}`);
      if (!dirs.length) throw new Error("\u627E\u4E0D\u5230 WPS \u52A0\u8F7D\u9879\u76EE\u5F55\uFF0C\u8BF7\u5148\u5B89\u88C5\u5E76\u8FD0\u884C WPS\u3002");
      const source = this.getSourceAddonPath();
      appendServiceLog("AddonInstaller", `\u5B89\u88C5\u6E90\u8D44\u6E90\u76EE\u5F55: ${source}`);
      const indexes = dirs.flatMap((dir) => ["publish.xml", "jsplugins.xml"].map((name) => {
        const file = path5.join(dir, name);
        appendServiceLog("AddonInstaller", `\u51C6\u5907\u6821\u9A8C/\u5408\u5E76\u7D22\u5F15\u6587\u4EF6: ${file}`);
        const content = mergePluginIndex(readXmlSafe(file));
        return { file, content };
      }));
      const token = getToken(true);
      for (const dir of dirs) {
        if (options.cleanBlocked !== false) {
          const blockFile = path5.join(dir, "jsaddinblockhost.ini");
          if (fs4.existsSync(blockFile)) {
            try {
              fs4.copyFileSync(blockFile, `${blockFile}.backup-${Date.now()}`);
              fs4.unlinkSync(blockFile);
              appendServiceLog("AddonInstaller", `\u5DF2\u6E05\u7406\u963B\u65AD\u6587\u4EF6: ${blockFile}`);
            } catch (err) {
              appendServiceLog("AddonInstaller", `\u6E05\u7406\u963B\u65AD\u6587\u4EF6\u5F02\u5E38: ${err.message}`);
            }
          }
        }
        try {
          const entries = fs4.readdirSync(dir);
          for (const entry of entries) {
            const isOwned = ownedNames.some((name) => entry.startsWith(name));
            if (isOwned && !entry.endsWith(`_${VERSION}`) && entry !== "wps-bridge") {
              const target = path5.join(dir, entry);
              if (fs4.existsSync(target) && fs4.statSync(target).isDirectory() && entry.includes(".")) {
                try {
                  fs4.rmSync(target, { recursive: true, force: true });
                } catch {
                }
              }
            }
          }
        } catch {
        }
        const aliases = ["wps-bridge", "office-agent-bridge", ...activeNames.slice(1).flatMap((name) => [`${name}_`, `${name}_${VERSION}`])];
        for (const alias of aliases) {
          const dest = path5.join(dir, alias);
          fs4.mkdirSync(dest, { recursive: true, mode: 448 });
          for (const file of fs4.readdirSync(source)) {
            const srcFile = path5.join(source, file);
            if (fs4.statSync(srcFile).isFile()) {
              fs4.copyFileSync(srcFile, path5.join(dest, file));
            }
          }
          atomicWrite(path5.join(dest, "bridge-config.js"), `window.WPS_BRIDGE_CONFIG = ${JSON.stringify({ port: runtimePort(), token, version: VERSION })};
`);
          const addonEntry = path5.join(dest, "addon-core.js");
          const indexPath = path5.join(dest, "index.html");
          if (fs4.existsSync(addonEntry) && fs4.existsSync(indexPath)) {
            try {
              const head = fs4.readFileSync(addonEntry, "utf8").slice(0, 4096);
              const fp = head.match(/ADDON_BUILD_FINGERPRINT:\s*([0-9a-f]{16,64})/);
              if (fp) {
                const html = fs4.readFileSync(indexPath, "utf8");
                const patched = html.replace(
                  /src="\.\/addon-core\.js(\?v=[0-9a-f]+)?"/g,
                  `src="./addon-core.js?v=${fp[1].slice(0, 16)}"`
                );
                if (patched !== html) fs4.writeFileSync(indexPath, patched, "utf8");
              }
            } catch (e) {
              appendServiceLog("AddonInstaller", `\u7F13\u5B58\u5931\u6548\u6539\u5199\u5931\u8D25\uFF08\u4E0D\u5F71\u54CD\u90E8\u7F72\uFF09: ${e.message}`);
            }
          }
        }
      }
      for (const index of indexes) {
        atomicWrite(index.file, index.content, true);
        appendServiceLog("AddonInstaller", `\u5199\u5165\u66F4\u65B0\u7D22\u5F15\u6587\u4EF6\u6210\u529F: ${index.file}`);
      }
      appendServiceLog("AddonInstaller", `=== WPS \u52A0\u8F7D\u9879\u90E8\u7F72\u6210\u529F\u5B8C\u6210 (v${VERSION}) ===`);
      return { success: true, message: `\u52A0\u8F7D\u9879\u5DF2\u6210\u529F\u90E8\u7F72\u5E76\u66F4\u65B0\u81F3 v${VERSION}\uFF01\u8BF7\u5728 WPS \u4E2D\u70B9\u51FB\u529F\u80FD\u533A\u3010\u91CD\u65B0\u8FDE\u63A5\u3011\u6216\u91CD\u542F WPS \u751F\u6548\u3002`, targetPath: dirs.join(" | "), warnings: dirs.filter((dir) => fs4.existsSync(path5.join(dir, "jsaddinblockhost.ini"))).map(() => "\u68C0\u6D4B\u5230 WPS \u963B\u65AD\u914D\u7F6E\uFF0C\u672A\u5220\u9664\u3002\u8BF7\u5728 WPS \u4E2D\u68C0\u67E5\u52A0\u8F7D\u9879\u6743\u9650\u3002") };
    } catch (e) {
      appendServiceLog("AddonInstaller", `WPS \u52A0\u8F7D\u9879\u5B89\u88C5\u5931\u8D25: ${e.message}
\u5806\u6808: ${e.stack || ""}`);
      const permissionIssue = detectPermissionIssueCore(e);
      if (permissionIssue) return { success: false, message: "WPS \u52A0\u8F7D\u9879\u76EE\u5F55\u9700\u8981 macOS \u5B8C\u5168\u78C1\u76D8\u8BBF\u95EE\u6743\u9650", permissionIssue };
      return { success: false, message: e.message };
    }
  }
  static uninstall() {
    try {
      const indexes = this.getAllAddonDirectories().flatMap((dir) => ["publish.xml", "jsplugins.xml"].map((name) => path5.join(dir, name))).filter((file) => fs4.existsSync(file)).map((file) => ({ file, content: mergePluginIndex(readXmlSafe(file), true) }));
      for (const item of indexes) atomicWrite(item.file, item.content, true);
      return { success: true, message: "\u5DF2\u6CE8\u9500\u672C\u9879\u76EE\u52A0\u8F7D\u9879\uFF1B\u4FDD\u7559\u6587\u4EF6\u4E0E\u5907\u4EFD\uFF0C\u91CD\u542F WPS \u540E\u751F\u6548\u3002" };
    } catch (e) {
      return { success: false, message: e.message };
    }
  }
};
var OfficeAddonInstaller = class {
  static getWefDirectory() {
    const home = os2.homedir();
    if (process.platform === "win32") {
      return path5.join(process.env.LOCALAPPDATA || path5.join(home, "AppData/Local"), "Microsoft/Office/16.0/Wef");
    }
    return path5.join(home, "Library/Containers/com.microsoft.Excel/Data/Documents/wef");
  }
  static getSourceManifestPath() {
    return resourcePath("office-addon/excel/manifest.xml");
  }
  static safeRead(filePath) {
    if (!fs4.existsSync(filePath)) return "";
    try {
      return fs4.readFileSync(filePath, "utf8");
    } catch {
      if (process.platform === "darwin") {
        try {
          return execSync2(`cat "${filePath}"`, { encoding: "utf8", stdio: ["pipe", "pipe", "ignore"] });
        } catch {
        }
      }
    }
    return "";
  }
  static safeDeploy(sourceFile, destFile) {
    appendServiceLog("AddonInstaller", `\u51C6\u5907\u90E8\u7F72 Office \u6E05\u5355: \u6E90\u6587\u4EF6=${sourceFile}, \u76EE\u6807\u6587\u4EF6=${destFile}`);
    if (!fs4.existsSync(sourceFile)) {
      const msg = `\u6E90\u6E05\u5355\u6587\u4EF6\u4E0D\u5B58\u5728: ${sourceFile}`;
      appendServiceLog("AddonInstaller", msg);
      throw new Error(msg);
    }
    const sourceContent = fs4.readFileSync(sourceFile, "utf8");
    if (fs4.existsSync(destFile)) {
      try {
        const destContent = fs4.readFileSync(destFile, "utf8");
        if (destContent === sourceContent || destContent.includes(`<Version>${VERSION}`) && destContent.includes("Office Agent Bridge")) {
          appendServiceLog("AddonInstaller", `\u76EE\u6807\u6E05\u5355\u6587\u4EF6\u5DF2\u5904\u4E8E\u6700\u65B0\u72B6\u6001 (v${VERSION})\uFF0C\u8DF3\u8FC7\u8986\u5199`);
          if (process.platform === "darwin") {
            try {
              execSync2(`chmod 644 "${destFile}"`, { stdio: "ignore" });
            } catch {
            }
          }
          return;
        }
      } catch (e) {
        appendServiceLog("AddonInstaller", `\u6BD4\u5BF9\u76EE\u6807\u6E05\u5355\u5931\u8D25: ${e.message}\uFF0C\u51C6\u5907\u8986\u5199`);
      }
    }
    const destDir = path5.dirname(destFile);
    try {
      fs4.mkdirSync(destDir, { recursive: true, mode: 493 });
    } catch (e) {
      appendServiceLog("AddonInstaller", `fs.mkdirSync \u63D0\u793A (${e.code || e.message})\uFF0C\u5C1D\u8BD5 shell \u521B\u5EFA\u76EE\u5F55`);
      if (process.platform === "darwin") {
        try {
          execSync2(`mkdir -p "${destDir}"`, { stdio: "ignore" });
        } catch {
        }
      }
    }
    let writeSucceeded = false;
    try {
      fs4.writeFileSync(destFile, sourceContent, { mode: 420 });
      writeSucceeded = true;
      appendServiceLog("AddonInstaller", `fs.writeFileSync \u76F4\u63A5\u5199\u5165\u76EE\u6807\u6E05\u5355\u6210\u529F`);
    } catch (directErr) {
      appendServiceLog("AddonInstaller", `fs.writeFileSync \u76F4\u63A5\u5199\u5165\u5931\u8D25: [${directErr.code || "UNKNOWN"}] ${directErr.message}`);
    }
    if (!writeSucceeded && process.platform === "darwin") {
      appendServiceLog("AddonInstaller", `macOS \u4E0B\u76F4\u63A5\u5199\u5165\u53D7\u9650\uFF0C\u5F00\u59CB\u591A\u5C42\u63D0\u6743/\u7CFB\u7EDF\u4EE3\u7406\u5199\u5165\u6D41\u7A0B...`);
      try {
        const tmpFile = path5.join(runtimeHome(), "temp-wps-bridge-manifest.xml");
        fs4.writeFileSync(tmpFile, sourceContent, { mode: 420 });
        const cmd = `mkdir -p "${destDir}" && cp -f "${tmpFile}" "${destFile}" && chmod 644 "${destFile}"`;
        const appleScript = `do shell script ${JSON.stringify(cmd)}`;
        try {
          execSync2(`osascript -e ${JSON.stringify(appleScript)}`, { encoding: "utf8", stdio: ["pipe", "pipe", "pipe"] });
          writeSucceeded = true;
          appendServiceLog("AddonInstaller", `\u901A\u8FC7 macOS \u539F\u751F osascript \u4EE3\u7406\u5199\u5165\u6E05\u5355\u6210\u529F`);
        } finally {
          try {
            if (fs4.existsSync(tmpFile)) fs4.unlinkSync(tmpFile);
          } catch {
          }
        }
      } catch (osaErr) {
        const osaMsg = osaErr.stderr || osaErr.message || String(osaErr);
        appendServiceLog("AddonInstaller", `osascript \u4EE3\u7406\u5199\u5165\u5931\u8D25: ${osaMsg}`);
        try {
          const tmpFile = path5.join(runtimeHome(), "temp-wps-bridge-manifest.xml");
          fs4.writeFileSync(tmpFile, sourceContent, { mode: 420 });
          try {
            execSync2(`cp -f "${tmpFile}" "${destFile}" && chmod 644 "${destFile}"`, { encoding: "utf8", stdio: ["pipe", "pipe", "pipe"] });
            writeSucceeded = true;
            appendServiceLog("AddonInstaller", `\u901A\u8FC7 fallback cp -f \u5199\u5165\u6E05\u5355\u6210\u529F`);
          } finally {
            try {
              if (fs4.existsSync(tmpFile)) fs4.unlinkSync(tmpFile);
            } catch {
            }
          }
        } catch (cpErr) {
          const cpMsg = cpErr.stderr || cpErr.message || String(cpErr);
          appendServiceLog("AddonInstaller", `cp -f \u547D\u4EE4\u6267\u884C\u53D7\u963B: ${cpMsg}`);
          const err = new Error(`macOS \u78C1\u76D8\u8BBF\u95EE\u53D7\u9650\uFF08${cpMsg.trim() || "EPERM: Operation not permitted"}\uFF09`);
          err.code = "EPERM";
          throw err;
        }
      }
    }
    if (!writeSucceeded && process.platform !== "darwin") {
      throw new Error(`\u65E0\u6CD5\u5199\u5165 Office \u52A0\u8F7D\u9879\u6E05\u5355\u6587\u4EF6: ${destFile}`);
    }
    if (process.platform === "darwin" && fs4.existsSync(destFile)) {
      try {
        execSync2(`chmod 644 "${destFile}"`, { stdio: "ignore" });
      } catch {
      }
    } else if (process.platform === "win32" && fs4.existsSync(destFile)) {
      try {
        appendServiceLog("AddonInstaller", `[Windows] \u5F00\u59CB\u6CE8\u518C WEF Developer \u6CE8\u518C\u8868\u9879`);
        execSync2(`reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "55555555-aaaa-bbbb-cccc-777777777777" /t REG_SZ /d "${destFile}" /f`, { stdio: "ignore" });
        appendServiceLog("AddonInstaller", `[Windows] \u6CE8\u518C\u8868\u5199\u5165\u6210\u529F`);
      } catch (regErr) {
        appendServiceLog("AddonInstaller", `[Windows] \u6CE8\u518C\u8868\u5199\u5165\u5F02\u5E38: ${regErr.message}`);
      }
      try {
        getOrGenerateCerts();
        const certPath = path5.join(runtimeHome(), "certs/localhost.crt");
        if (fs4.existsSync(certPath)) {
          appendServiceLog("AddonInstaller", `[Windows] \u5BFC\u5165\u53D7\u4FE1\u4EFB\u6839\u8BC1\u4E66: ${certPath}`);
          execSync2(`certutil -user -addstore "Root" "${certPath}"`, { stdio: "ignore" });
          appendServiceLog("AddonInstaller", `[Windows] \u8BC1\u4E66\u5BFC\u5165\u5B8C\u6210`);
        }
      } catch (certErr) {
        appendServiceLog("AddonInstaller", `[Windows] \u8BC1\u4E66\u5BFC\u5165\u5F02\u5E38: ${certErr.message}`);
      }
    }
  }
  static async checkStatus() {
    appendServiceLog("AddonInstaller", "\u5F00\u59CB\u68C0\u6D4B Office \u52A0\u8F7D\u9879\u72B6\u6001");
    try {
      const res = await serviceRequest("/api/v1/office/addon-status");
      appendServiceLog("AddonInstaller", `\u901A\u8FC7\u540E\u53F0\u670D\u52A1\u83B7\u53D6\u72B6\u6001\u6210\u529F: installed=${res.installed}, current=${res.current}`);
      return res;
    } catch (e) {
      appendServiceLog("AddonInstaller", `\u540E\u53F0\u670D\u52A1\u72B6\u6001\u8BF7\u6C42\u672A\u54CD\u5E94 (${e.message})\uFF0C\u8F6C\u5165\u672C\u5730\u6587\u4EF6\u76F4\u63A5\u68C0\u6D4B`);
      const wefDir = this.getWefDirectory();
      const targetFile = path5.join(wefDir, "wps-bridge-manifest.xml");
      const installed = fs4.existsSync(targetFile);
      let current = false;
      let ver = "\u672A\u90E8\u7F72";
      if (installed) {
        try {
          const content = fs4.readFileSync(targetFile, "utf8");
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
        message: !installed ? "\u5C1A\u672A\u90E8\u7F72 Office \u5B98\u65B9\u52A0\u8F7D\u9879\u6E05\u5355" : !current ? `\u68C0\u6D4B\u5230 Office \u52A0\u8F7D\u9879\u9700\u8981\u66F4\u65B0\uFF08\u5F53\u524D: ${ver}\uFF0C\u6700\u65B0: ${VERSION}\uFF09` : `Office \u5B98\u65B9\u52A0\u8F7D\u9879\u5DF2\u5C31\u7EEA (v${ver})`
      };
      appendServiceLog("AddonInstaller", `\u672C\u5730\u68C0\u6D4B\u7ED3\u679C: ${JSON.stringify(status)}`);
      return status;
    }
  }
  static async install() {
    appendServiceLog("AddonInstaller", "\u89E6\u53D1 Office \u52A0\u8F7D\u9879\u90E8\u7F72/\u5347\u7EA7\u6D41\u7A0B");
    const wefDir = this.getWefDirectory();
    const targetFile = path5.join(wefDir, "wps-bridge-manifest.xml");
    try {
      const source = this.getSourceManifestPath();
      this.safeDeploy(source, targetFile);
      appendServiceLog("AddonInstaller", `Office \u52A0\u8F7D\u9879\u90E8\u7F72\u6210\u529F: ${targetFile}`);
      return {
        success: true,
        message: "Office \u5B98\u65B9\u52A0\u8F7D\u9879\u6E05\u5355\u4E0E\u672C\u5730\u5B89\u5168\u914D\u7F6E\u5DF2\u5C31\u7EEA\uFF01\u8BF7\u5728 Excel \u4E2D\u91CD\u65B0\u6253\u5F00\u4FA7\u8FB9\u680F\u3002",
        targetPath: targetFile
      };
    } catch (e) {
      appendServiceLog("AddonInstaller", `Office \u52A0\u8F7D\u9879\u90E8\u7F72\u5931\u8D25: ${e.message}`);
      const permissionIssue = detectPermissionIssueCore(e);
      if (permissionIssue) return { success: false, message: "Office \u52A0\u8F7D\u9879\u76EE\u5F55\u9700\u8981 macOS \u5B8C\u5168\u78C1\u76D8\u8BBF\u95EE\u6743\u9650", permissionIssue };
      return { success: false, message: e.message };
    }
  }
  static uninstall() {
    appendServiceLog("AddonInstaller", "\u89E6\u53D1 Office \u52A0\u8F7D\u9879\u5378\u8F7D\u6D41\u7A0B");
    try {
      const wefDir = this.getWefDirectory();
      const target = path5.join(wefDir, "wps-bridge-manifest.xml");
      if (fs4.existsSync(target)) {
        try {
          fs4.unlinkSync(target);
          appendServiceLog("AddonInstaller", `fs.unlinkSync \u5220\u9664\u6E05\u5355\u6210\u529F: ${target}`);
        } catch (unlinkErr) {
          appendServiceLog("AddonInstaller", `fs.unlinkSync \u5931\u8D25 (${unlinkErr.message})\uFF0C\u5C1D\u8BD5 shell \u5220\u9664`);
          if (process.platform === "darwin") {
            try {
              execSync2(`rm -f "${target}"`, { stdio: "ignore" });
            } catch {
            }
          } else if (process.platform === "win32") {
            try {
              execSync2('reg delete "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "55555555-aaaa-bbbb-cccc-777777777777" /f', { stdio: "ignore" });
            } catch {
            }
          }
        }
      }
      return { success: true, message: "\u5DF2\u79FB\u9664 Office \u52A0\u8F7D\u9879\u6E05\u5355\uFF0C\u91CD\u542F Excel \u540E\u751F\u6548\u3002" };
    } catch (e) {
      appendServiceLog("AddonInstaller", `Office \u52A0\u8F7D\u9879\u5378\u8F7D\u5931\u8D25: ${e.message}`);
      const permissionIssue = detectPermissionIssueCore(e);
      if (permissionIssue) return { success: false, message: "Office \u52A0\u8F7D\u9879\u76EE\u5F55\u9700\u8981 macOS \u5B8C\u5168\u78C1\u76D8\u8BBF\u95EE\u6743\u9650", permissionIssue };
      return { success: false, message: e.message };
    }
  }
};

// src/main/installer-engine.ts
function mergeMcpConfig(file, entry) {
  let config = {};
  if (fs5.existsSync(file)) {
    const raw = fs5.readFileSync(file, "utf8");
    try {
      config = JSON.parse(raw);
    } catch {
      throw new Error("\u73B0\u6709 JSON \u65E0\u6CD5\u89E3\u6790\uFF0C\u672A\u8986\u76D6\u3002\u8BF7\u5148\u4FEE\u590D\u539F\u914D\u7F6E\u3002");
    }
    if (!config || typeof config !== "object" || Array.isArray(config)) throw new Error("\u73B0\u6709\u914D\u7F6E\u4E0D\u662F JSON \u5BF9\u8C61\uFF0C\u672A\u8986\u76D6");
  }
  if (config.mcpServers !== void 0 && (!config.mcpServers || typeof config.mcpServers !== "object" || Array.isArray(config.mcpServers))) throw new Error("mcpServers \u683C\u5F0F\u65E0\u6548\uFF0C\u672A\u8986\u76D6");
  const LEGACY = "wps-bridge";
  const prevLegacy = config.mcpServers[LEGACY];
  const legacyIsOurs = prevLegacy && typeof prevLegacy === "object" && !Array.isArray(prevLegacy) && Boolean(entry && typeof entry === "object" && prevLegacy.command === entry.command && JSON.stringify(prevLegacy.args) === JSON.stringify(entry.args));
  const next = { ...config.mcpServers, "office-agent-bridge": entry };
  if (legacyIsOurs) delete next[LEGACY];
  config.mcpServers = next;
  atomicWrite(file, JSON.stringify(config, null, 2) + "\n", true);
}
var InstallerEngine = class {
  static runtimeEntry = "";
  static getSystemTargetDir() {
    return runtimeHome();
  }
  static configEntry() {
    return { command: process.execPath, args: [this.runtimeEntry || resourcePath("dist/bridge/cli.cjs")], env: { ELECTRON_RUN_AS_NODE: "1", OFFICE_AGENT_BRIDGE_HOME: runtimeHome(), WPS_BRIDGE_HOME: runtimeHome(), WPS_BRIDGE_PORT: String(runtimePort()), WPS_BRIDGE_RESOURCES: resourcePath("package.json").replace(/[\\/]package\.json$/, "") } };
  }
  static detectEnvironment() {
    const home = os3.homedir();
    const appData = process.platform === "darwin" ? path6.join(home, "Library/Application Support") : process.env.APPDATA || path6.join(home, "AppData/Roaming");
    const entries = [
      ["doubao", "\u8C46\u5305\u684C\u9762\u7AEF", path6.join(home, ".doubao/mcp.json"), path6.join(home, ".doubao/skills")],
      ["workbuddy", "WorkBuddy", path6.join(home, ".workbuddy/mcp.json"), path6.join(home, ".workbuddy/skills")],
      ["qwen", "\u5343\u95EE\u529E\u516C", path6.join(home, ".qwen/mcp.json"), path6.join(home, ".qwen/skills")],
      ["kimi", "Kimi", path6.join(home, ".kimi-code/mcp.json"), path6.join(home, ".kimi-code/skills")],
      ["claude-code", "Claude Code", path6.join(home, ".claude.json"), path6.join(home, ".claude/skills")],
      ["codex", "Codex", path6.join(home, ".codex/mcp.json"), path6.join(home, ".codex/skills")]
    ];
    const agents = entries.map(([id, name, configPath, skillsPath]) => {
      let configured = false, detail = "";
      try {
        if (fs5.existsSync(configPath)) {
          const srv = JSON.parse(fs5.readFileSync(configPath, "utf8")).mcpServers;
          configured = Boolean(srv?.["office-agent-bridge"] || srv?.["wps-bridge"]);
        }
      } catch {
        detail = "\u914D\u7F6E JSON \u65E0\u6CD5\u89E3\u6790\uFF0C\u4FEE\u590D\u524D\u4E0D\u4F1A\u8986\u76D6";
      }
      const detected = fs5.existsSync(path6.dirname(configPath));
      return { id, name, configPath, skillsPath, detected, status: configured ? "configured" : detected ? "installed" : "not_found", detail };
    });
    return { platform: process.platform, systemTargetDir: runtimeHome(), addon: AddonInstaller.checkStatus(), agents };
  }
  static async executeInstall(options = {}) {
    const logs = [];
    if (options.addon) {
      const r = AddonInstaller.install();
      logs.push({ step: "WPS \u52A0\u8F7D\u9879", status: r.success ? "success" : "error", detail: r.message });
    }
    for (const agent of this.detectEnvironment().agents.filter((a) => options.agents?.includes(a.id))) {
      try {
        mergeMcpConfig(agent.configPath, this.configEntry());
        if (options.skills && agent.skillsPath) {
          const source = resourcePath("skills");
          for (const name of fs5.readdirSync(source)) {
            const dir = path6.join(source, name);
            if (fs5.existsSync(path6.join(dir, "SKILL.md"))) this.copySkills(dir, path6.join(agent.skillsPath, name));
          }
        }
        logs.push({ step: agent.name, status: "success", detail: "\u5DF2\u5408\u5E76 MCP \u914D\u7F6E\uFF1B\u652F\u6301\u6280\u80FD\u76EE\u5F55\u7684\u5BA2\u6237\u7AEF\u5DF2\u6309\u9009\u62E9\u540C\u6B65\u6280\u80FD\u3002" });
      } catch (e) {
        logs.push({ step: agent.name, status: "error", detail: e.message });
      }
    }
    return { success: logs.every((l) => l.status === "success"), message: logs.length ? "\u6240\u9009\u914D\u7F6E\u5904\u7406\u5B8C\u6210\uFF0C\u8BF7\u67E5\u770B\u6BCF\u9879\u7ED3\u679C\u3002" : "\u672A\u9009\u62E9\u9700\u8981\u5B89\u88C5\u7684\u9879\u76EE\u3002", logs };
  }
  static copySkills(source, dest) {
    for (const entry of fs5.readdirSync(source, { withFileTypes: true })) {
      if (entry.isSymbolicLink()) continue;
      const from = path6.join(source, entry.name), to = path6.join(dest, entry.name);
      if (entry.isDirectory()) this.copySkills(from, to);
      else atomicWrite(to, fs5.readFileSync(from, "utf8"), true);
    }
  }
};

// src/main/addon-autoupgrade.ts
var RELOAD_SETTLE_MS = 4e3;
async function ensureAddonUpToDate(options = {}) {
  const { force = false, reload = true } = options;
  let freshness;
  try {
    freshness = AddonInstaller.checkBuildFreshness();
  } catch (e) {
    return { checked: false, stale: null, upgraded: false, reloaded: false, message: `\u65E0\u6CD5\u68C0\u6D4B\u52A0\u8F7D\u9879\u6784\u5EFA\u72B6\u6001\uFF1A${e?.message ?? e}` };
  }
  if (freshness.stale === null) {
    return { checked: true, stale: null, upgraded: false, reloaded: false, message: freshness.reason, detail: freshness };
  }
  if (!freshness.stale && !force) {
    return { checked: true, stale: false, upgraded: false, reloaded: false, message: freshness.reason, detail: freshness };
  }
  const installFn = options.installFn ?? (() => AddonInstaller.install());
  let installResult;
  try {
    installResult = await installFn();
  } catch (e) {
    return { checked: true, stale: true, upgraded: false, reloaded: false, message: `\u91CD\u65B0\u90E8\u7F72\u52A0\u8F7D\u9879\u5931\u8D25\uFF1A${e?.message ?? e}`, detail: freshness };
  }
  if (installResult && installResult.success === false) {
    return {
      checked: true,
      stale: true,
      upgraded: false,
      reloaded: false,
      message: `\u91CD\u65B0\u90E8\u7F72\u52A0\u8F7D\u9879\u672A\u6210\u529F\uFF1A${installResult.message ?? "\u89C1\u5B89\u88C5\u5668\u8FD4\u56DE"}`,
      detail: { freshness, installResult }
    };
  }
  if (!reload) {
    return {
      checked: true,
      stale: true,
      upgraded: true,
      reloaded: false,
      message: "\u5DF2\u91CD\u65B0\u90E8\u7F72\u52A0\u8F7D\u9879\uFF1B\u672A\u89E6\u53D1\u91CD\u8F7D\uFF08\u8C03\u7528\u65B9\u8981\u6C42\u8DF3\u8FC7\uFF09\uFF0C\u8FD0\u884C\u4E2D\u7684\u52A0\u8F7D\u9879\u8981\u7B49\u4E0B\u6B21\u91CD\u8F7D\u6216\u91CD\u542F WPS \u624D\u751F\u6548",
      detail: { freshness, installResult }
    };
  }
  const reloadFn = options.reloadFn ?? (() => serviceRequest("/api/v1/tool/call", { name: "wps_reload_addon", arguments: {}, sessionId: "desktop" }));
  try {
    await reloadFn();
    await new Promise((resolve) => setTimeout(resolve, RELOAD_SETTLE_MS));
    return {
      checked: true,
      stale: true,
      upgraded: true,
      reloaded: true,
      message: "\u5DF2\u91CD\u65B0\u90E8\u7F72\u52A0\u8F7D\u9879\u5E76\u89E6\u53D1\u91CD\u65B0\u52A0\u8F7D\uFF1B\u8FD0\u884C\u4E2D\u7684\u52A0\u8F7D\u9879\u65E0\u9700\u91CD\u542F WPS \u5373\u53EF\u751F\u6548",
      detail: { freshness, installResult }
    };
  } catch (e) {
    return {
      checked: true,
      stale: true,
      upgraded: true,
      reloaded: false,
      message: `\u52A0\u8F7D\u9879\u5DF2\u91CD\u65B0\u90E8\u7F72\uFF0C\u4F46\u89E6\u53D1\u91CD\u8F7D\u5931\u8D25\uFF08${e?.message ?? e}\uFF09\uFF1A\u9700\u8981\u624B\u52A8\u70B9\u51FB\u529F\u80FD\u533A\u3010\u91CD\u65B0\u8FDE\u63A5\u3011\u6216\u91CD\u542F WPS`,
      detail: { freshness, installResult }
    };
  }
}

// src/main/index.ts
var here2 = path7.dirname(fileURLToPath2(import.meta.url));
process.env.WPS_BRIDGE_RESOURCES ||= app3.getAppPath();
var root = app3.getAppPath();
app3.setPath("userData", path7.join(runtimeHome(), "desktop-state"));
var cliPath = path7.join(root, "dist/bridge/cli.cjs").replace(/app\.asar([\\/])/, "app.asar.unpacked$1");
InstallerEngine.runtimeEntry = cliPath;
var window = null;
var tray = null;
var timer;
var notice = "";
var addonUpgradeResult = null;
var prefsPath = path7.join(runtimeHome(), "desktop.json");
var prefs = {};
try {
  prefs = JSON.parse(fs6.readFileSync(prefsPath, "utf8"));
} catch {
}
if (["system", "light", "dark"].includes(prefs.theme)) nativeTheme2.themeSource = prefs.theme;
function persist() {
  atomicWrite(prefsPath, JSON.stringify(prefs));
}
function createWindow() {
  if (window) {
    window.show();
    window.focus();
    return;
  }
  let appIcon;
  if (process.platform !== "darwin") {
    try {
      appIcon = nativeImage2.createFromPath(resourcePath("resources/icon.png"));
    } catch {
    }
  }
  window = new BrowserWindow2({
    width: Math.max(780, Math.min(prefs.width || 880, 1300)),
    height: Math.max(560, Math.min(prefs.height || 620, 950)),
    minWidth: 780,
    minHeight: 560,
    show: false,
    ...appIcon ? { icon: appIcon } : {},
    title: "Office Agent Bridge",
    titleBarStyle: process.platform === "darwin" ? "hiddenInset" : "default",
    backgroundColor: nativeTheme2.shouldUseDarkColors ? "#17191c" : "#f7f8fa",
    webPreferences: { preload: path7.join(here2, "../preload/index.cjs"), contextIsolation: true, nodeIntegration: false, sandbox: true }
  });
  window.setMenuBarVisibility(false);
  if (process.env.VITE_DEV_SERVER_URL) void window.loadURL(process.env.VITE_DEV_SERVER_URL);
  else void window.loadFile(path7.join(here2, "../renderer/index.html"));
  window.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
  window.webContents.on("will-navigate", (event) => event.preventDefault());
  window.once("ready-to-show", () => window?.show());
  window.on("resized", () => {
    if (window && !window.isMaximized()) {
      const [width, height] = window.getSize();
      prefs = { ...prefs, width, height };
      persist();
    }
  });
  window.on("closed", () => {
    window = null;
  });
}
async function state() {
  try {
    return { online: true, ...await serviceRequest("/api/v1/status"), notice: "" };
  } catch {
    return { online: false, isWpsConnected: false, components: {}, notice };
  }
}
function setupIpc() {
  ipcMain.handle("get-status", state);
  ipcMain.handle("service-start", async () => {
    try {
      await ensureService(cliPath);
      notice = "";
      return { success: true };
    } catch (e) {
      notice = e.message;
      return { success: false, message: e.message };
    }
  });
  ipcMain.handle("service-stop", async () => {
    await serviceRequest("/api/v1/service/stop", {});
    return { success: true };
  });
  ipcMain.handle("diagnose", () => serviceRequest("/api/v1/tool/call", { name: "bridge_diagnose", arguments: {} }).then((r) => r.data));
  ipcMain.handle("office-status", () => serviceRequest("/api/v1/tool/call", { name: "office_get_status", arguments: {} }).then((r) => r.data));
  ipcMain.handle("get-audit-records", (_e, options = {}) => serviceRequest("/api/v1/tool/call", { name: "wps_get_audit_history", arguments: { ...options, limit: 20, view: "summary" } }).then((r) => r.data));
  ipcMain.handle("clear-audit-records", () => serviceRequest("/api/v1/tool/call", { name: "wps_clear_audit_history", arguments: {} }).then((r) => r.data));
  ipcMain.handle("get-audit-record", (_e, id) => serviceRequest("/api/v1/tool/call", { name: "wps_get_audit_record", arguments: { auditId: id } }).then((r) => r.data));
  ipcMain.handle("rollback-record", (_e, id) => serviceRequest("/api/v1/tool/call", { name: "wps_rollback", arguments: { auditId: id }, sessionId: "desktop" }).then((r) => r.data));
  ipcMain.handle("check-addon-status", () => AddonInstaller.checkStatus());
  ipcMain.handle("check-addon-freshness", () => AddonInstaller.checkBuildFreshness());
  ipcMain.handle("get-addon-upgrade-result", () => addonUpgradeResult);
  ipcMain.handle("ensure-addon-uptodate", async (_e, opts) => {
    appendServiceLog("IPC", `\u6536\u5230\u52A0\u8F7D\u9879\u5347\u7EA7\u8BF7\u6C42 (force=${Boolean(opts?.force)})`);
    addonUpgradeResult = await ensureAddonUpToDate({ force: Boolean(opts?.force) });
    return addonUpgradeResult;
  });
  let lastPermissionRetry = null;
  const withPermissionTarget = (r) => {
    if (!r?.permissionIssue) return r;
    const issue = { ...r.permissionIssue, appToAuthorize: appToAuthorize(), isDev: !app3.isPackaged, label: r.message };
    appendServiceLog("PermissionWindow", `\u6743\u9650\u5931\u8D25\uFF0C\u51C6\u5907\u6253\u5F00\u5F15\u5BFC\u6D6E\u7A97: ${issue?.targetPath || "(\u65E0\u8DEF\u5F84)"}`);
    try {
      openPermissionWindow(issue);
      appendServiceLog("PermissionWindow", "\u5F15\u5BFC\u6D6E\u7A97\u5DF2\u521B\u5EFA");
    } catch (e) {
      appendServiceLog("PermissionWindow", `\u6253\u5F00\u5F15\u5BFC\u6D6E\u7A97\u5931\u8D25: ${e?.message || e}`);
    }
    return { ...r, permissionIssue: issue };
  };
  ipcMain.handle("install-addon", async () => {
    appendServiceLog("IPC", "\u6536\u5230\u524D\u7AEF\u4E00\u952E\u5B89\u88C5 / \u5347\u7EA7 WPS \u52A0\u8F7D\u9879\u8BF7\u6C42");
    lastPermissionRetry = async () => withPermissionTarget(await AddonInstaller.install());
    return withPermissionTarget(await AddonInstaller.install());
  });
  ipcMain.handle("check-office-addon-status", () => OfficeAddonInstaller.checkStatus());
  ipcMain.handle("install-office-addon", async () => {
    lastPermissionRetry = async () => withPermissionTarget(await OfficeAddonInstaller.install());
    return withPermissionTarget(await OfficeAddonInstaller.install());
  });
  ipcMain.handle("installer:detect", () => InstallerEngine.detectEnvironment());
  ipcMain.handle("installer:execute", (_e, options) => InstallerEngine.executeInstall(options));
  ipcMain.handle("permission:full-disk-access-info", () => ({ url: FULL_DISK_ACCESS_URL, appToAuthorize: appToAuthorize(), isDev: !app3.isPackaged, platform: process.platform }));
  ipcMain.handle("permission:open-full-disk-access", async () => {
    await shell.openExternal(FULL_DISK_ACCESS_URL);
    return { success: true };
  });
  ipcMain.on("permission:drag-start", (event) => {
    try {
      startAppDrag(event.sender);
    } catch (e) {
      appendServiceLog("PermissionWindow", `\u53D1\u8D77\u62D6\u62FD\u5931\u8D25: ${e?.message || e}`);
    }
  });
  ipcMain.handle("permission:app-icon", () => loadAppIcon());
  ipcMain.handle("permission:open-guide", () => {
    openPermissionWindow(getPermissionIssue());
    return { success: true };
  });
  ipcMain.handle("permission:close-window", () => {
    closePermissionWindow();
    return { success: true };
  });
  ipcMain.handle("permission:retry-install", async () => lastPermissionRetry ? lastPermissionRetry() : { success: false, message: "\u6CA1\u6709\u53EF\u91CD\u8BD5\u7684\u5B89\u88C5" });
  ipcMain.handle("get-app-info", () => ({ version: VERSION, port: runtimePort(), home: runtimeHome(), platform: process.platform, theme: prefs.theme || "system", login: app3.getLoginItemSettings().openAtLogin, config: { mcpServers: { "office-agent-bridge": InstallerEngine.configEntry(), "wps-bridge": InstallerEngine.configEntry() } } }));
  ipcMain.handle("set-theme", (_e, theme) => {
    if (!["system", "light", "dark"].includes(theme)) throw new Error("\u4E3B\u9898\u65E0\u6548");
    nativeTheme2.themeSource = theme;
    prefs.theme = theme;
    persist();
    return true;
  });
  ipcMain.handle("set-login", (_e, enabled) => {
    if (!app3.isPackaged) throw new Error("\u767B\u5F55\u542F\u52A8\u4EC5\u5728\u5B89\u88C5\u7248\u4E2D\u53EF\u7528");
    app3.setLoginItemSettings({ openAtLogin: Boolean(enabled), args: ["--background"] });
    return app3.getLoginItemSettings().openAtLogin;
  });
  ipcMain.handle("copy-text", (_e, text) => {
    if (typeof text !== "string" || text.length > 1e5) throw new Error("\u6587\u672C\u65E0\u6548");
    clipboard.writeText(text);
    return true;
  });
  ipcMain.handle("open-log", () => {
    const file = path7.join(runtimeHome(), "service.log");
    if (!fs6.existsSync(file)) {
      appendServiceLog("System", `\u65E5\u5FD7\u6587\u4EF6\u5DF2\u521D\u59CB\u5316 (\u5E73\u53F0: ${process.platform}, \u7248\u672C: ${VERSION})`);
    }
    shell.openPath(file);
    return true;
  });
  ipcMain.handle("exit-app", () => app3.quit());
}
if (!app3.requestSingleInstanceLock()) app3.quit();
else {
  app3.on("second-instance", createWindow);
  app3.whenReady().then(async () => {
    setupIpc();
    const isMac = process.platform === "darwin";
    let trayIcon;
    if (isMac) {
      trayIcon = nativeImage2.createFromPath(resourcePath("resources/trayTemplate.png"));
      trayIcon.setTemplateImage(true);
    } else {
      trayIcon = nativeImage2.createFromPath(resourcePath("resources/tray-win.png"));
    }
    tray = new Tray(trayIcon);
    tray.setToolTip("Office Agent Bridge");
    tray.setContextMenu(Menu.buildFromTemplate([{ label: "\u6253\u5F00 Office Agent Bridge", click: createWindow }, { label: "\u505C\u6B62\u670D\u52A1", click: () => {
      void serviceRequest("/api/v1/service/stop", {}).catch((e) => {
        notice = e.message;
      });
    } }, { type: "separator" }, { label: "\u9000\u51FA\u7BA1\u7406\u7A97\u53E3\uFF08\u540E\u53F0\u7EE7\u7EED\u8FD0\u884C\uFF09", click: () => app3.quit() }]));
    tray.on("click", createWindow);
    if (!process.argv.includes("--background")) createWindow();
    ensureService(cliPath).then(() => ensureAddonUpToDate()).then((r) => {
      addonUpgradeResult = r;
      appendServiceLog("AutoUpgrade", r.message);
    }).catch((e) => {
      notice = e.message;
    });
    timer = setInterval(async () => {
      if (window && !window.isDestroyed()) window.webContents.send("status-changed", await state());
    }, 2e3);
    app3.on("activate", createWindow);
  });
  app3.on("window-all-closed", () => {
  });
  app3.on("before-quit", () => {
    clearInterval(timer);
    tray?.destroy();
  });
}
//# sourceMappingURL=index.js.map