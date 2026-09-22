/**
 * Microsoft Office 原生驱动通道工具定义（P2.3 剩余项：逐类迁移自 gateway.getOpenAiTools）。
 *
 * 已发布顺序中的第 3–7 条，**按前缀聚合**：本模块含全部 `office_*` 定义。
 * 注意它与 `toolClassOf` 的分类标签不同口径：`office_lock_target` / `office_unlock_target`
 * 的标签是 `lock`（与 WPS 目标锁同类），其余三条才是 `microsoft`。
 * 若按标签重排会改变已发布的工具顺序，因此这里保持网关原始顺序。
 */
import type { GatewayToolDefinition } from './shared.js';

/** Microsoft Office 原生驱动通道：已发布顺序第 3–7 条。定义顺序即契约顺序，不得重排。 */
export function microsoftToolDefinitions(): GatewayToolDefinition[] {
  return [
    {
      type: "function",
      function: {
        name: "office_get_status",
        description: "【Microsoft Office 全局状态探测器】探测当前电脑（macOS / Windows）上 Microsoft Word、Excel、PowerPoint 运行进程与打开的文档列表，完全免配置插件。",
        parameters: {
          type: "object",
          properties: {},
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "office_lock_target",
        description: "【Microsoft Office 目标文档锁定器】将 Agent 强隔离锁定在指定的 Microsoft Office 文档上，避免用户切窗口导致的意外漂移。",
        parameters: {
          type: "object",
          properties: {
            component: { type: "string", enum: ["word", "excel", "ppt"], description: "组件类型" },
            targetName: { type: "string", description: "目标文档名称" }
          },
          required: ["component", "targetName"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "office_unlock_target",
        description: "解除 Microsoft Office 文档锁定状态。",
        parameters: {
          type: "object",
          properties: {
            component: { type: "string", enum: ["word", "excel", "ppt"] }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "office_execute_script",
        description: "【Microsoft Office 原生脚本引擎】直接向 Microsoft Word、Excel、PowerPoint 运行实例执行原生自动化代码（macOS 下为 JXA/AppleScript，Windows 下为 PowerShell COM 自动化），实现 100% 任意操作无死角！",
        parameters: {
          type: "object",
          properties: {
            component: { type: "string", enum: ["word", "excel", "ppt"], description: "目标组件类型" },
            script: { type: "string", description: "要执行的原生脚本代码" },
            targetName: { type: "string", description: "目标文档名称（可选，若已锁定则自动复用）" },
            params: { type: "object", description: "传递给脚本的结构化数据" }
          },
          required: ["component", "script"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "office_capture_slide_preview",
        description: "【Microsoft PowerPoint 高清快照】直接从 Microsoft PowerPoint 导出指定幻灯片的矢量渲染快照，用于高保真自检比对。",
        parameters: {
          type: "object",
          properties: {
            slideIndex: { type: "integer", description: "幻灯片页码(1-based)" },
            presentationName: { type: "string", description: "目标演示文稿名称（可选）" }
          },
          required: ["slideIndex"],
          additionalProperties: false
        }
      }
    }
  ];
}
