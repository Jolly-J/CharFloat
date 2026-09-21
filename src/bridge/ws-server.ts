import http from 'node:http';
import https from 'node:https';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import crypto from 'node:crypto';
import { WebSocketServer, WebSocket } from 'ws';
import { SSEServerTransport } from '@modelcontextprotocol/sdk/server/sse.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { createMcpServer } from './mcp-server.js';
import { executeCatalogTool, getTools, capabilities } from './catalog.js';
import { requestContext } from './context.js';
import { VERSION, PROTOCOL, runtimePort, runtimeHttpsPort, getOrGenerateCerts, getToken, validToken, resourcePath } from './runtime.js';
import type { SelectionInfo } from './types.js';

export interface ComponentStatus { connected: boolean; activeDocument?: string; activeSheet?: string; version?: string; lastHeartbeat?: number; summary?: any }
export interface BridgeState {
  isWpsConnected: boolean;
  isMsOfficeConnected?: boolean;
  components: { word: ComponentStatus; excel: ComponentStatus; ppt: ComponentStatus; msExcel?: ComponentStatus; msWord?: ComponentStatus; msPpt?: ComponentStatus };
  activeWorkbook?: string; activeSheet?: string; activeDocument?: string; activePresentation?: string;
  currentSelection?: SelectionInfo | null; wpsClientVersion?: string; addonNeedsUpgrade?: boolean; lastUpdated: number;
}

async function readBody(req: http.IncomingMessage) {
  let size = 0; const chunks: Buffer[] = [];
  for await (const chunk of req) {
    size += chunk.length;
    if (size > 8 * 1024 * 1024) throw new Error('请求超过 8 MB，请分批操作');
    chunks.push(chunk);
  }
  return chunks.length ? JSON.parse(Buffer.concat(chunks).toString('utf8')) : {};
}
export function requestAllowed(req: http.IncomingMessage, port: number) {
  const host = req.headers.host || '';
  const allowedHosts = [`127.0.0.1:${port}`, `localhost:${port}`, '127.0.0.1:19890', 'localhost:19890', '127.0.0.1:19891', 'localhost:19891'];
  if (!allowedHosts.includes(host)) return false;
  const origin = req.headers.origin;
  // File-hosted WPS uses an opaque origin; possession of the install token is still required.
  if (!origin || origin === 'null' || origin.startsWith('file://') || origin.startsWith('ksapp://')) return true;
  const allowedOrigins = [
    `http://127.0.0.1:${port}`, `http://localhost:${port}`,
    `https://127.0.0.1:${port}`, `https://localhost:${port}`,
    'http://127.0.0.1:19890', 'http://localhost:19890',
    'https://127.0.0.1:19891', 'https://localhost:19891'
  ];
  return allowedOrigins.includes(origin);
}
export class WpsBridgeServer {
  private server?: http.Server;
  private httpsServer?: https.Server;
  private wss?: WebSocketServer;
  private sockets = new Map<string, WebSocket>();
  private pending = new Map<string, { socket: WebSocket; resolve: (v: any) => void; reject: (e: Error) => void; timer: NodeJS.Timeout }>();
  private sessions = new Map<string, StreamableHTTPServerTransport>();
  private sse = new Map<string, SSEServerTransport>();
  private listeners = new Set<(s: BridgeState) => void>();
  private heartbeat?: NodeJS.Timeout;
  private state: BridgeState = { isWpsConnected: false, isMsOfficeConnected: false, components: { word: { connected: false }, excel: { connected: false }, ppt: { connected: false }, msExcel: { connected: false } }, lastUpdated: Date.now() };
  constructor(private port = runtimePort(), private httpsPort = runtimeHttpsPort()) {}
  async start() {
    getToken(true);
    this.server = http.createServer((req, res) => { this.handleHttp(req, res).catch(error => {
      if (!res.headersSent) res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: false, error: error.message }));
    }); });
    this.wss = new WebSocketServer({ noServer: true, maxPayload: 32 * 1024 * 1024 });

    const handleUpgrade = (req: http.IncomingMessage, socket: any, head: Buffer) => {
      const isHttps = Boolean((req.socket as any).encrypted);
      const port = isHttps ? this.httpsPort : this.port;
      const url = new URL(req.url || '/', `${isHttps ? 'https' : 'http'}://127.0.0.1:${port}`);
      const token = url.searchParams.get('token');
      const originAllowed = requestAllowed(req, port);
      const isOfficeAddon = url.pathname === '/office-addon';
      const isAddonPath = url.pathname === '/addon' || isOfficeAddon;
      const isTokenValid = isOfficeAddon || validToken(token || undefined);

      if (!isAddonPath || !originAllowed || !isTokenValid) {
        const reason = !isAddonPath
          ? `未知路径: ${url.pathname}`
          : !originAllowed
            ? `Origin 受限: ${req.headers.origin} (Host: ${req.headers.host})`
            : !token
              ? 'Token 缺失'
              : 'Token 无效或已过期';
        console.error(`[Bridge] 握手被拦截 (${reason})。`);
        socket.write('HTTP/1.1 401 Unauthorized\r\nConnection: close\r\n\r\n'); socket.destroy(); return;
      }
      this.wss!.handleUpgrade(req, socket, head, ws => this.acceptSocket(ws));
    };

    this.server.on('upgrade', handleUpgrade);
    await new Promise<void>((resolve, reject) => { this.server!.once('error', reject); this.server!.listen(this.port, '127.0.0.1', resolve); });

    const certs = getOrGenerateCerts();
    if (certs) {
      try {
        this.httpsServer = https.createServer(certs, (req, res) => {
          this.handleHttp(req, res, true).catch(error => {
            if (!res.headersSent) res.writeHead(400, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: false, error: error.message }));
          });
        });
        this.httpsServer.on('upgrade', handleUpgrade);
        await new Promise<void>((resolve) => {
          this.httpsServer!.once('error', (err) => {
            console.warn('[Bridge] 本地 HTTPS 服务启动跳过:', err.message);
            resolve();
          });
          this.httpsServer!.listen(this.httpsPort, '127.0.0.1', resolve);
        });
        console.error(`[Bridge] HTTPS 静态与 WSS 通道已就绪: https://localhost:${this.httpsPort}`);
      } catch (err: any) {
        console.warn('[Bridge] HTTPS 监听初始化跳过:', err.message);
      }
    }

    this.heartbeat = setInterval(() => {
      for (const ws of this.wss!.clients) {
        if ((ws as any).alive === false) { ws.terminate(); continue; }
        (ws as any).alive = false; ws.ping();
      }
    }, 15000);
    console.error(`[Bridge] ${VERSION} listening on 127.0.0.1:${this.port}`);
  }

  private async handleHttp(req: http.IncomingMessage, res: http.ServerResponse, isHttps = false) {
    const json = (value: unknown, code = 200) => { res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }); res.end(JSON.stringify(value)); };
    const port = isHttps ? this.httpsPort : this.port;
    if (!requestAllowed(req, port)) return json({ error: '来源不允许' }, 403);
    const url = new URL(req.url || '/', `${isHttps ? 'https' : 'http'}://127.0.0.1:${port}`);

    // Office Add-in 静态资源托管与图标支持
    if (url.pathname.startsWith('/office-addon/') || url.pathname.startsWith('/assets/')) {
      try {
        const publicDir = resourcePath('office-addon/public');
        const baseName = url.pathname.startsWith('/assets/') ? path.basename(url.pathname) : url.pathname.replace(/^\/office-addon\//, '');
        const filePath = path.join(publicDir, baseName);
        if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
          const ext = path.extname(filePath);
          const mimeTypes: Record<string, string> = {
            '.html': 'text/html; charset=utf-8',
            '.js': 'application/javascript; charset=utf-8',
            '.css': 'text/css; charset=utf-8',
            '.png': 'image/png',
            '.svg': 'image/svg+xml'
          };
          res.writeHead(200, {
            'Content-Type': mimeTypes[ext] || 'application/octet-stream',
            'Access-Control-Allow-Origin': '*',
            'Cache-Control': 'no-cache, no-store, must-revalidate, max-age=0',
            'Pragma': 'no-cache',
            'Expires': '0'
          });
          res.end(fs.readFileSync(filePath));
          return;
        }
      } catch {}
    }

    if (url.pathname === '/health' && req.method === 'GET') return json({ service: 'wps-bridge', protocol: PROTOCOL, version: VERSION, pid: process.pid });
    if (req.method === 'OPTIONS') {
      if (req.headers.origin) res.setHeader('Access-Control-Allow-Origin', req.headers.origin);
      res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, Mcp-Session-Id, MCP-Protocol-Version');
      res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE');
      res.writeHead(204); res.end(); return;
    }
    const auth = req.headers.authorization?.replace(/^Bearer /, '');
    if (!validToken(auth)) return json({ error: 'Bridge HTTP 401：需要本机安装凭据' }, 401);
    if (req.headers.origin) res.setHeader('Access-Control-Allow-Origin', req.headers.origin);
    res.setHeader('Access-Control-Expose-Headers', 'Mcp-Session-Id');
    if (req.method === 'OPTIONS') {
      res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, Mcp-Session-Id, MCP-Protocol-Version');
      res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE'); return json({});
    }
    if (url.pathname === '/mcp') {
      const id = req.headers['mcp-session-id'] as string | undefined;
      if (req.method === 'POST') {
        const body = await readBody(req);
        let transport = id ? this.sessions.get(id) : undefined;
        if (id && !transport) return json({ error: '会话已过期，请重新初始化' }, 404);
        if (!transport) {
          if (body.method !== 'initialize') return json({ error: '请先初始化 MCP 会话' }, 400);
          if (this.sessions.size >= 64) return json({ error: '会话数量达到上限' }, 429);
          const server = createMcpServer();
          transport = new StreamableHTTPServerTransport({ sessionIdGenerator: () => crypto.randomUUID(), onsessioninitialized: sid => { this.sessions.set(sid, transport!); } });
          transport.onclose = () => { if (transport?.sessionId) this.sessions.delete(transport.sessionId); void server.close(); };
          await server.connect(transport);
        }
        await transport.handleRequest(req, res, body); return;
      }
      const transport = id ? this.sessions.get(id) : undefined;
      if (!transport) return json({ error: '会话不存在' }, 404);
      await transport.handleRequest(req, res); return;
    }
    if (url.pathname === '/sse' && req.method === 'GET') {
      const transport = new SSEServerTransport('/messages', res);
      this.sse.set(transport.sessionId, transport);
      const server = createMcpServer();
      res.on('close', () => { this.sse.delete(transport.sessionId); void server.close(); });
      await server.connect(transport); return;
    }
    if (url.pathname === '/messages' && req.method === 'POST') {
      const transport = this.sse.get(url.searchParams.get('sessionId') || '');
      if (!transport) return json({ error: '会话不存在' }, 404);
      await transport.handlePostMessage(req, res, await readBody(req)); return;
    }
    if (req.method === 'GET') {
      if (url.pathname === '/api/v1/status') return json({ ...this.getState(), service: { version: VERSION, pid: process.pid, port: this.port, background: true } });
      if (url.pathname === '/api/v1/mcp-tools') return json(getTools());
      if (url.pathname === '/api/v1/capabilities') return json(capabilities());
      if (url.pathname === '/openapi.json') {
        const { UniversalGateway } = await import('./gateway.js');
        const schema: any = UniversalGateway.getOpenApiSchema(`http://127.0.0.1:${this.port}`);
        schema.info.version = VERSION; schema.security = [{ bearerAuth: [] }];
        schema.components = { securitySchemes: { bearerAuth: { type: 'http', scheme: 'bearer' } } };
        return json(schema);
      }
      if (url.pathname === '/api/v1/tools') return json(getTools().map(t => ({ type: 'function', function: { name: t.name, description: t.description, parameters: t.inputSchema } })));
      if (url.pathname === '/api/v1/prompts/aesthetic') { res.writeHead(200, { 'Content-Type': 'text/markdown; charset=utf-8' }); res.end(fs.readFileSync(resourcePath('prompts/excel_aesthetic_system.md'), 'utf8')); return; }
      if (url.pathname === '/api/v1/office/addon-status') {
        const home = os.homedir();
        const wefDir = process.platform === 'win32'
          ? path.join(process.env.LOCALAPPDATA || path.join(home, 'AppData/Local'), 'Microsoft/Office/16.0/Wef')
          : path.join(home, 'Library/Containers/com.microsoft.Excel/Data/Documents/wef');
        const targetFile = path.join(wefDir, 'wps-bridge-manifest.xml');
        const installed = fs.existsSync(targetFile);
        let current = false;
        let ver = '';
        if (installed) {
          try {
            const content = fs.readFileSync(targetFile, 'utf8');
            const m = content.match(/<Version>(.*?)<\/Version>/i);
            ver = m ? m[1].trim() : '';
            current = ver === VERSION || ver === `${VERSION}.0` || ver.startsWith(VERSION);
          } catch {
            current = true;
            ver = VERSION;
          }
        }
        return json({
          installed,
          current,
          latestVersion: VERSION,
          installedVersion: ver || (installed ? '已部署' : '未部署'),
          needsUpgrade: !installed || !current,
          targetPath: targetFile,
          platform: process.platform,
          message: !installed ? '尚未部署 Office 官方加载项清单' : !current ? `检测到 Office 加载项需要更新（当前: ${ver}，最新: ${VERSION}）` : `Office 官方加载项已就绪 (v${ver})`
        });
      }
    }
    if (url.pathname === '/api/v1/office/install-addon' && req.method === 'POST') {
      try {
        appendServiceLog('WsServer', '收到 /api/v1/office/install-addon 部署请求');
        const home = os.homedir();
        const wefDir = process.platform === 'win32'
          ? path.join(process.env.LOCALAPPDATA || path.join(home, 'AppData/Local'), 'Microsoft/Office/16.0/Wef')
          : path.join(home, 'Library/Containers/com.microsoft.Excel/Data/Documents/wef');
        const targetFile = path.join(wefDir, 'wps-bridge-manifest.xml');
        const sourcePath = resourcePath('office-addon/excel/manifest.xml');
        const content = fs.readFileSync(sourcePath, 'utf8');

        if (fs.existsSync(targetFile)) {
          try {
            const existing = fs.readFileSync(targetFile, 'utf8');
            if (existing === content || existing.includes(`<Version>${VERSION}`)) {
              appendServiceLog('WsServer', '目标清单已为最新版本，跳过覆写');
              return json({
                success: true,
                message: 'Office 官方加载项清单已处于最新状态！请在 Excel 中点击【插入 -> 我的加载项】启用。',
                targetPath: targetFile
              });
            }
          } catch {}
        }

        fs.mkdirSync(wefDir, { recursive: true, mode: 0o755 });
        fs.writeFileSync(targetFile, content, { mode: 0o644 });
        if (process.platform === 'darwin') {
          try {
            const { execSync } = await import('node:child_process');
            execSync(`chmod 644 "${targetFile}"`, { stdio: 'ignore' });
          } catch {}
        } else if (process.platform === 'win32') {
          try {
            const { execSync } = await import('node:child_process');
            execSync(`reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "55555555-aaaa-bbbb-cccc-777777777777" /t REG_SZ /d "${targetFile}" /f`, { stdio: 'ignore' });
          } catch {}
        }

        appendServiceLog('WsServer', `Office 加载项部署就绪: ${targetFile}`);
        return json({
          success: true,
          message: 'Office 官方加载项已成功部署！请在 Excel 中点击【插入 -> 我的加载项】启用。',
          targetPath: targetFile
        });
      } catch (e: any) {
        appendServiceLog('WsServer', `Office 加载项部署失败: ${e.message}`);
        return json({ success: false, error: e.message }, 500);
      }
    }
    if (url.pathname === '/api/v1/office/reload' && req.method === 'POST') {
      const ws = this.sockets.get('ms-excel');
      if (!ws || ws.readyState !== WebSocket.OPEN) {
        return json({ success: false, error: 'Microsoft Excel 任务窗格未连接' }, 404);
      }
      ws.send(JSON.stringify({
        id: 'hot-reload-' + Date.now(),
        method: 'run_script',
        params: { code: 'window.location.reload(true);' }
      }));
      return json({ success: true, message: '已向 Excel 发送热重载信号' });
    }
    if (url.pathname === '/api/v1/tool/call' && req.method === 'POST') {
      const body = await readBody(req);
      const args = body.arguments || {};
      const name = body.name || body.toolName;
      // Only unified excel_* accepts host switching. Legacy names always target WPS.
      const host = name?.startsWith('excel_') && args.host === 'microsoft' ? 'microsoft' : 'wps';
      const sessionId = typeof body.sessionId === 'string' ? body.sessionId : 'http-local';
      try {
        const data = await requestContext.run({ sessionId, host }, () => executeCatalogTool(name, args, body.clientName || 'HTTP Agent'));
        return json({ success: true, data });
      } catch (e: any) { return json({ success: false, error: e.message }, 422); }
    }
    if (url.pathname === '/api/v1/service/stop' && req.method === 'POST') {
      json({ success: true }); setTimeout(() => { this.stop(); process.exit(0); }, 100); return;
    }
    return json({ error: 'Not found' }, 404);
  }
  private acceptSocket(ws: WebSocket) {
    (ws as any).alive = true;
    ws.on('pong', () => { (ws as any).alive = true; });
    ws.on('error', () => {});
    ws.on('message', raw => {
      let p: any; try { p = JSON.parse(raw.toString()); } catch { return; }
      if (p.type === 'register') {
        const client = String(p.client || '');
        const isMs = p.host === 'microsoft' || client.startsWith('ms-');
        const key = isMs
          ? (/word/.test(client) ? 'ms-word' : /ppt|powerpoint/.test(client) ? 'ms-ppt' : 'ms-excel')
          : (/word/.test(client) ? 'word' : /ppt|wpp|presentation/.test(client) ? 'ppt' : 'excel');

        const previous = this.sockets.get(key);
        this.sockets.set(key, ws);
        if (previous && previous !== ws) previous.close(1000, 'replaced');
        const summary = p.summary || {};
        const clientVersion = String(p.version || '未知');

        if (isMs) {
          const compKey = key === 'ms-excel' ? 'msExcel' : key === 'ms-word' ? 'msWord' : 'msPpt';
          (this.state.components as any)[compKey] = {
            connected: true,
            version: clientVersion,
            lastHeartbeat: Date.now(),
            activeDocument: summary.fullName || summary.workbookName || summary.documentName,
            activeSheet: summary.activeSheetName,
            summary
          };
          this.state.isMsOfficeConnected = true;
        } else {
          this.state.components[key as 'excel'] = { connected: true, version: clientVersion, lastHeartbeat: Date.now(), activeDocument: summary.fullName || summary.workbookName || summary.documentName || summary.presentationName, activeSheet: summary.activeSheetName, summary };
          if (key === 'excel') Object.assign(this.state, { activeWorkbook: summary.workbookName, activeSheet: summary.activeSheetName, currentSelection: summary.selection, wpsClientVersion: clientVersion });
          if (key === 'word') this.state.activeDocument = summary.documentName;
          if (key === 'ppt') this.state.activePresentation = summary.presentationName;
        }

        if (clientVersion !== VERSION && !isMs) {
          try {
            ws.send(JSON.stringify({
              type: 'version_notice',
              latestVersion: VERSION,
              currentVersion: clientVersion,
              message: `检测到加载项版本 (${clientVersion}) 与当前 Bridge 守护进程 (${VERSION}) 不一致，建议升级`
            }));
          } catch {}
        }
        this.publish();
      } else if (p.type === 'rpc_response') {
        const pending = this.pending.get(p.id);
        if (!pending || pending.socket !== ws) return;
        clearTimeout(pending.timer); this.pending.delete(p.id);
        if (p.error) pending.reject(new Error(p.error)); else pending.resolve(p.result);
      } else if (p.type === 'event' && p.event === 'selection_change') {
        if (this.sockets.get('excel') === ws || this.sockets.get('ms-excel') === ws) {
          this.state.currentSelection = p.data;
          if (p.data?.sheetName) this.state.activeSheet = p.data.sheetName;
          this.publish();
        }
      }
    });
    ws.on('close', () => {
      for (const [key, current] of this.sockets) if (current === ws) {
        this.sockets.delete(key);
        if (key.startsWith('ms-')) {
          const compKey = key === 'ms-excel' ? 'msExcel' : key === 'ms-word' ? 'msWord' : 'msPpt';
          (this.state.components as any)[compKey] = { connected: false };
        } else {
          this.state.components[key as 'excel'] = { connected: false };
          if (key === 'excel') { this.state.activeWorkbook = undefined; this.state.activeSheet = undefined; this.state.currentSelection = null; }
        }
      }
      for (const [id, p] of this.pending) if (p.socket === ws) { clearTimeout(p.timer); p.reject(new Error('加载项断开，执行结果未知，请先读回确认')); this.pending.delete(id); }
      this.publish();
    });
  }
  async callWps<T = any>(method: string, params: any = {}, timeoutMs = 20000): Promise<T> {
    let key = params.component || (method.startsWith('word_') || params.documentName ? 'word' : method.startsWith('ppt_') || params.presentationName ? 'ppt' : 'excel');
    if (method === 'get_workspace_summary' && !params.workbookName && !this.sockets.has('excel') && this.sockets.size === 1) key = [...this.sockets.keys()][0];
    const ws = this.sockets.get(key);
    if (!ws || ws.readyState !== WebSocket.OPEN) throw new Error(`WPS ${key} 加载项未连接，请先打开对应组件并检查 Bridge 加载项。`);
    const id = crypto.randomUUID();
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => { this.pending.delete(id); reject(new Error(`WPS ${method} 超时，结果未知；先读取状态，不要自动重放写入`)); }, timeoutMs);
      this.pending.set(id, { socket: ws, resolve, reject, timer });
      ws.send(JSON.stringify({ id, method, params }), error => { if (error) { clearTimeout(timer); this.pending.delete(id); reject(error); } });
    });
  }
  async callOfficeAddon<T = any>(method: string, params: any = {}, timeoutMs = 25000): Promise<T> {
    const key = params.component === 'word' ? 'ms-word' : params.component === 'ppt' ? 'ms-ppt' : 'ms-excel';
    const ws = this.sockets.get(key);
    if (!ws || ws.readyState !== WebSocket.OPEN) throw new Error(`Microsoft Office (${key}) 加载项未连接，请在 Excel 中打开【WPS Bridge (Excel AI)】任务窗格。`);
    const id = crypto.randomUUID();
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => { this.pending.delete(id); reject(new Error(`Microsoft Office ${method} 超时，请检查 Excel 任务窗格连接`)); }, timeoutMs);
      this.pending.set(id, { socket: ws, resolve, reject, timer });
      ws.send(JSON.stringify({ id, method, params }), error => { if (error) { clearTimeout(timer); this.pending.delete(id); reject(error); } });
    });
  }
  getState() {
    const copy = structuredClone(this.state);
    const hasOlder = Object.values(copy.components).some(c => c.connected && c.version && c.version !== VERSION);
    copy.addonNeedsUpgrade = hasOlder;
    return copy;
  }
  subscribeState(listener: (s: BridgeState) => void) { this.listeners.add(listener); listener(this.getState()); return () => this.listeners.delete(listener); }
  private publish() {
    this.state.isWpsConnected = Object.entries(this.state.components).some(([k, v]) => !k.startsWith('ms') && v.connected);
    this.state.isMsOfficeConnected = Object.entries(this.state.components).some(([k, v]) => k.startsWith('ms') && v.connected);
    this.state.lastUpdated = Date.now();
    for (const l of this.listeners) l(this.getState());
  }
  stop() {
    clearInterval(this.heartbeat);
    for (const p of this.pending.values()) { clearTimeout(p.timer); p.reject(new Error('Bridge 已停止')); }
    this.pending.clear();
    for (const ws of this.wss?.clients || []) ws.terminate();
    for (const t of this.sessions.values()) void t.close();
    for (const t of this.sse.values()) void t.close();
    this.wss?.close();
    this.server?.close();
    this.httpsServer?.close();
  }
}
export const bridgeServer = new WpsBridgeServer();

