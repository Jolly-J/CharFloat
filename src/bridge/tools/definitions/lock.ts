/**
 * WPS 目标锁工具定义（P2.3 剩余项：逐类迁移自 gateway.getOpenAiTools）。
 *
 * 已发布顺序中的第 8–10 条，分类标签为 `lock`。
 * `wps_get_locked_status` 只读查询锁状态，加解锁会改本地状态。
 * Microsoft 侧的 `office_lock_target` / `office_unlock_target` 留在 `microsoft.ts`
 * （已发布顺序中位于 office_ 区块内，不能移动）。
 */
import type { GatewayToolDefinition } from './shared.js';

/** WPS 目标锁：已发布顺序第 8–10 条。定义顺序即契约顺序，不得重排。 */
export function lockToolDefinitions(): GatewayToolDefinition[] {
  return [
    {
      type: "function",
      function: {
        name: "wps_lock_target_document",
        description: "锁定 WPS 目标文档：锁定后不带目标名称的调用都指向该文件，避免用户切换窗口造成漂移。一次只锁一个组件；要换目标需先解锁或重新锁定。",
        parameters: {
          type: "object",
          properties: {
            component: { type: "string", enum: ["word", "excel", "ppt"], description: "目标组件类型: 'word'|'excel'|'ppt'" },
            targetName: { type: "string", description: "目标文档名称或文件路径，需与已打开文件列表中的名称一致，如 '副本极电光能半年度方针复盘-26.8.20.pptx'" }
          },
          required: ["component", "targetName"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_unlock_target_document",
        description: "解除 WPS 文档锁定状态。解除后无目标名称的调用不再有确定目标，后续操作需显式传工作簿/文档名称。",
        parameters: {
          type: "object",
          properties: {
            component: { type: "string", enum: ["word", "excel", "ppt"], description: "要解锁的组件: 'word'|'excel'|'ppt'；不传则解锁全部" }
          },
          required: [],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_get_locked_status",
        description: "查询当前 Word、Excel、PPT 各组件的目标锁定状态。注意：锁定记录只说明锁指向哪个文件，不代表该文件仍在打开；需要“当前实际打开了哪些文件”时用 wps_get_workspace_summary。",
        parameters: {
          type: "object",
          properties: {},
          required: [],
          additionalProperties: false
        }
      }
    }
  ];
}
