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
        description: "探测本机（macOS / Windows）Microsoft Word、Excel、PowerPoint 的运行进程与已打开文档列表，不需要安装加载项。返回进程与文档清单，用于在调用前确认目标确实打开。",
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
        description: "锁定 Microsoft Office 目标文档：锁定后不带目标名称的调用都指向该文件，避免用户切换窗口造成漂移。一次只锁一个组件（word/excel/ppt）。",
        parameters: {
          type: "object",
          properties: {
            component: { type: "string", enum: ["word", "excel", "ppt"], description: "组件类型: 'word'|'excel'|'ppt'" },
            targetName: { type: "string", description: "目标文档名称（可含路径），必须与 office_get_status 返回的名称一致" }
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
            component: { type: "string", enum: ["word", "excel", "ppt"], description: "要解锁的组件: 'word'|'excel'|'ppt'；不传则解锁全部组件" }
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
        description: "直接向 Microsoft Word、Excel、PowerPoint 运行实例执行原生自动化代码：macOS 下为 JXA/AppleScript，Windows 下为 PowerShell COM。结构化工具未覆盖的 Microsoft 侧操作走本工具（Office.js 通道未连接时的 Excel、以及 Word/PPT 的全部操作）。脚本自行定位目标文档；执行后先读回验证再继续，本工具不判断结果是否符合预期。该通道在 macOS 与 Windows 均尚未实机验收。",
        parameters: {
          type: "object",
          properties: {
            component: { type: "string", enum: ["word", "excel", "ppt"], description: "目标组件类型: 'word'|'excel'|'ppt'" },
            script: { type: "string", description: "要执行的原生脚本代码。macOS 传 JXA/AppleScript，Windows 传 PowerShell，两平台语法不可混用。" },
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
        description: "从 Microsoft PowerPoint 导出指定幻灯片的渲染快照，返回 imageBase64 与 imagePath，用于视觉自检比对。该通道在 macOS 与 Windows 均尚未实机验收。",
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
