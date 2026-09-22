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
        description: "【强隔离文档锁定器】将 Agent 强制锁定在指定的目标文档/工作簿/演示文稿上。一旦锁定，所有后续操作将严格只针对被锁定文件，用户在电脑上切换窗口绝不漂移！",
        parameters: {
          type: "object",
          properties: {
            component: { type: "string", enum: ["word", "excel", "ppt"], description: "目标应用组件类型 ('word', 'excel', 'ppt')" },
            targetName: { type: "string", description: "目标文档名称或文件路径，如 '副本极电光能半年度方针复盘-26.8.20.pptx'" }
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
        description: "解除文档锁定状态。解除后可重新自由寻址或切换目标文档。",
        parameters: {
          type: "object",
          properties: {
            component: { type: "string", enum: ["word", "excel", "ppt"], description: "要解锁的组件，不传则解锁全部" }
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
        description: "查询当前 Word、Excel、PPT 各组件的目标文档锁定状态及已打开的文件列表。",
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
