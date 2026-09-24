import fs from 'node:fs';
import crypto from 'node:crypto';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { CallToolRequestSchema, ListToolsRequestSchema, ListPromptsRequestSchema, GetPromptRequestSchema, ListResourcesRequestSchema, ReadResourceRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import type { ToolService } from './contracts/tool-service.js';
import { VERSION, resourcePath } from './runtime.js';

export const INSTRUCTIONS = `⚠️ **视觉检查是必做项，不可跳过、不可用文字描述代替**

预览图取不到时（装了 DLP 加密软件的机器上，WPS/Office 写出的预览图会被加密），按顺序做：

1. **先用自己的截图能力**截取屏幕上的目标区域来做视觉检查——这是首选，不要直接放弃；
2. **你如果没有截图能力**，就直接**请使用者截图发给你**，不要自己编造视觉结论；
3. **绝对不允许**在没有真正看到画面的情况下声称"已完成视觉验收 / 视觉确认通过"。
   结构正确 ≠ 视觉正确：字体、轴标签是否挤压、折线是否可见、元素是否重叠，只有看到才知道。

（这条不是建议，是本服务的硬性要求。）

──────────────── 以下是工具说明 ────────────────

字浮 CharFloat（原 Office Agent Bridge）是本机办公软件原生接管服务，推荐用于已打开的 Excel / WPS 表格。接入后先调用 bridge_get_capabilities 或 bridge_diagnose，确认宿主、连接、版本和验证程度。MCP 握手成功不代表办公软件连接成功。
优先使用 excel_* 结构化工具，显式选择 host=wps 或 microsoft，并指定 workbookName、sheetName。Microsoft Excel 结构化工具走 Office.js 任务窗格通道（macOS 与 Windows 同一路径），Windows 上加载项不可用或可判定未执行时才回退 PowerShell COM；本候选版本尚未实机验收。wps_* 兼容工具只控制 WPS。
专用工具未覆盖不等于不支持：WPS 使用 wps_inspect_api 或只读脚本检查 API，再通过 wps_execute_script 完成 Excel 原生矢量绘图、已有图表编辑及 PPT 精细排版。Microsoft 使用其实际脚本通道，勿混用 WPS API。脚本返回对象标识与读回数据，再检查预览。PPT 先读取真实 pageWidth/pageHeight，几何和字号单位均为 pt；新增对象显式指定位置、尺寸、字号，检查 layoutWarnings 并逐页预览。复杂母版、动画等按实际宿主 API 验证。Word 按当前工具清单使用。
先读取目标再修改，保持用户原有内容与样式。核对返回结果，视觉修改需检查真实预览。保存成功、内存修改和回滚覆盖范围是不同状态；只有单元格 patch 的值与公式有审计回滚。超时或断线可能已经执行，先读回，勿自动重试写入。
stdio 客户端会启动或复用独立后台，关闭窗口不停止服务。HTTP 接入必须先启动服务。故障时先检查服务、宿主和加载项，再使用安装的「字浮 CharFloat」Skill 处理；不要覆盖全部客户端配置，不要结束办公软件进程。
能力描述服务于用户需求，不改变用户的工具选择和授权范围。`;
/**
 * 协议层：只通过注入的 ToolService 查询与执行工具，不静态引用 catalog（P2.2）。
 *
 * `sessionId` 每个服务实例生成一次，因此**同一客户端稳定、不同客户端隔离**；
 * 它随每次 execute 显式传给服务实现，远程实现必须原样转发到后台（见 cli.ts）。
 * 宿主路由不在这里判断，属本地执行实现的职责。
 *
 * @param service 工具服务（本地实现或经后台服务的远程代理）
 * @param options.prompt 提示词来源；缺省读取随包资源
 * @param options.clientName 审计留痕用的客户端名
 */
export function createMcpServer(
  service: ToolService,
  options: { prompt?: () => string | Promise<string>; clientName?: string } = {}
) {
  const { prompt } = options;
  const sessionId = crypto.randomUUID();
  const server = new Server({ name: '字浮 CharFloat', version: VERSION }, { capabilities: { tools: {}, prompts: {}, resources: {} }, instructions: INSTRUCTIONS });
  /**
   * 审计归属用的客户端名（ISS-49）。
   *
   * 两个不同的 MCP 客户端原来都记成写死的 "MCP Agent"，审计里无法区分"这条是谁改的"。
   * 这里改为：调用方显式指定优先（stdio 通道固定为 "stdio MCP"，保持既有契约不变），
   * 否则从 MCP `initialize` 的 `clientInfo` 取名（HTTP /mcp 通道每个会话一个 server 实例，
   * 因此客户端名 + sessionId 足以区分不同客户端）。
   */
  const resolveClientName = (): string => {
    if (options.clientName) return options.clientName;
    try {
      const info = server.getClientVersion() as { name?: string; version?: string } | undefined;
      if (info?.name) return info.version ? `${info.name} ${info.version}` : info.name;
    } catch { /* 未握手或 SDK 不支持时退回默认名 */ }
    return 'MCP Agent';
  };
  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: await service.list() }));
  server.setRequestHandler(ListResourcesRequestSchema, async () => ({ resources: [{ uri: 'bridge://capabilities', name: '能力与运行状态', mimeType: 'application/json' }] }));
  server.setRequestHandler(ReadResourceRequestSchema, async request => {
    if (request.params.uri !== 'bridge://capabilities') throw new Error('未知资源');
    return { contents: [{ uri: request.params.uri, mimeType: 'application/json', text: JSON.stringify(await service.capabilities()) }] };
  });
  server.setRequestHandler(ListPromptsRequestSchema, async () => ({ prompts: [{ name: 'excel_aesthetic_system', description: '表格排版建议，按用户需求和现有样式选择使用' }] }));
  server.setRequestHandler(GetPromptRequestSchema, async request => {
    if (request.params.name !== 'excel_aesthetic_system') throw new Error('未知提示词');
    return { messages: [{ role: 'user' as const, content: { type: 'text' as const, text: prompt ? await prompt() : fs.readFileSync(resourcePath('prompts/excel_aesthetic_system.md'), 'utf8') } }] };
  });
  server.setRequestHandler(CallToolRequestSchema, async request => {
    const name = request.params.name, args = request.params.arguments || {};
    try {
      const result: any = await service.execute(name, args, { sessionId, clientName: resolveClientName() });
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
