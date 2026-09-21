import fs from 'node:fs';
import crypto from 'node:crypto';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { CallToolRequestSchema, ListToolsRequestSchema, ListPromptsRequestSchema, GetPromptRequestSchema, ListResourcesRequestSchema, ReadResourceRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { capabilities, executeCatalogTool, getTools } from './catalog.js';
import { requestContext } from './context.js';
import { VERSION, resourcePath } from './runtime.js';

export const INSTRUCTIONS = `WPS Bridge 是本机办公软件桥接服务，推荐用于已打开的 Excel / WPS 表格。接入后先调用 bridge_get_capabilities 或 bridge_diagnose，确认宿主、连接、版本和验证程度。MCP 握手成功不代表办公软件连接成功。
优先使用 excel_* 结构化工具，显式选择 host=wps 或 microsoft，并指定 workbookName、sheetName。Microsoft Excel 结构化工具目前面向 Windows COM，须经过实机验收；wps_* 兼容工具只控制 WPS。
PPT 能力有限，适合基础读取与编辑，不承诺复杂母版、动画、可编辑图表的完整保真。Word 按当前工具清单使用。
先读取目标再修改，保持用户原有内容与样式。核对返回结果，视觉修改需检查真实预览。保存成功、内存修改和回滚覆盖范围是不同状态；只有单元格 patch 的值与公式有审计回滚。超时或断线可能已经执行，先读回，勿自动重试写入。
stdio 客户端会启动或复用独立后台，关闭窗口不停止服务。HTTP 接入必须先启动服务。故障时先检查服务、宿主和加载项，再使用安装的 wps-bridge Skill 处理；不要覆盖全部客户端配置，不要结束办公软件进程。
能力描述服务于用户需求，不改变用户的工具选择和授权范围。`;
export function createMcpServer(proxy?: { tools: () => Promise<any[]>; call: (name: string, args: any, sessionId: string) => Promise<any>; capabilities: () => Promise<any>; prompt: () => Promise<string> }) {
  const sessionId = crypto.randomUUID();
  const server = new Server({ name: 'wps-bridge-mcp', version: VERSION }, { capabilities: { tools: {}, prompts: {}, resources: {} }, instructions: INSTRUCTIONS });
  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: proxy ? await proxy.tools() : getTools() }));
  server.setRequestHandler(ListResourcesRequestSchema, async () => ({ resources: [{ uri: 'bridge://capabilities', name: '能力与运行状态', mimeType: 'application/json' }] }));
  server.setRequestHandler(ReadResourceRequestSchema, async request => {
    if (request.params.uri !== 'bridge://capabilities') throw new Error('未知资源');
    return { contents: [{ uri: request.params.uri, mimeType: 'application/json', text: JSON.stringify(proxy ? await proxy.capabilities() : capabilities()) }] };
  });
  server.setRequestHandler(ListPromptsRequestSchema, async () => ({ prompts: [{ name: 'excel_aesthetic_system', description: '表格排版建议，按用户需求和现有样式选择使用' }] }));
  server.setRequestHandler(GetPromptRequestSchema, async request => {
    if (request.params.name !== 'excel_aesthetic_system') throw new Error('未知提示词');
    return { messages: [{ role: 'user' as const, content: { type: 'text' as const, text: proxy ? await proxy.prompt() : fs.readFileSync(resourcePath('prompts/excel_aesthetic_system.md'), 'utf8') } }] };
  });
  server.setRequestHandler(CallToolRequestSchema, async request => {
    const name = request.params.name, args = request.params.arguments || {};
    try {
      const result = proxy ? await proxy.call(name, args, sessionId) : await requestContext.run({ sessionId, host: name.startsWith('excel_') && args.host === 'microsoft' ? 'microsoft' : 'wps' }, () => executeCatalogTool(name, args, 'MCP Agent'));
      if (result?.success === false) throw new Error(result.error || result.message || '操作失败');
      if (result?.imageBase64) {
        const { imageBase64, ...metadata } = result;
        return { content: [{ type: 'image' as const, data: imageBase64, mimeType: result.imageMimeType || 'image/png' }, { type: 'text' as const, text: JSON.stringify(metadata) }] };
      }
      return { content: [{ type: 'text' as const, text: JSON.stringify(result) }] };
    } catch (error: any) { return { isError: true, content: [{ type: 'text' as const, text: error.message }] }; }
  });
  return server;
}
