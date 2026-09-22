/**
 * 原生脚本与 API 反射工具定义（P2.3 剩余项：逐类迁移自 gateway.getOpenAiTools）。
 *
 * 已发布顺序中的第 1–2 条，分类标签（tools/index.ts 的 toolClassOf）为 `script`。
 * 契约：可跨组件执行原生表达式，**不是**只读工具（见 contracts/host-methods.ts）。
 * 定义顺序即 tools/list 与契约快照顺序，不得调整。
 */
import type { GatewayToolDefinition } from './shared.js';

/** 原生脚本与 API 反射：已发布顺序第 1–2 条。定义顺序即契约顺序，不得重排。 */
export function scriptToolDefinitions(): GatewayToolDefinition[] {
  return [
    {
      type: "function",
      function: {
        name: "wps_execute_script",
        description: "直接向当前 WPS 运行实例执行原生 JavaScript。注入变量：app（WPS 宿主）、wb（当前 Excel 工作簿；Excel 场景用 wb，doc 在该场景为 null）、doc（当前 Word 文档）、pres（当前 PPT 演示文稿）、params（本工具 params 参数）。专用工具未覆盖时使用，例如 Excel 矢量形状、连接线、组合、已有图表属性修改与 PPT 精细布局。先用 wps_inspect_api 或只读脚本确认目标 API，再执行并读回验证；没有专用工具不等于宿主不支持。返回值需是可序列化的扁平 JSON（深层嵌套会丢属性），勿返回宿主对象。",
        parameters: {
          type: "object",
          properties: {
            code: { type: "string", description: "要执行的原生 JavaScript 代码。支持 async/await。执行成功不代表符合预期，需自行读回。" },
            workbookName: { type: "string", description: "精确目标工作簿名" },
            documentName: { type: "string", description: "精确目标 Word 文档名" },
            presentationName: { type: "string", description: "精确目标 PPT 文稿名" },
            component: { type: "string", enum: ["word", "excel", "ppt"], description: "组件类型: 'word'|'excel'|'ppt'；不传则按锁定目标或活动文档自动适配" },
            params: { type: "object", description: "传递给脚本的结构化数据，在代码中通过 params 对象访问" }
          },
          required: ["code"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_inspect_api",
        description: "在运行时反射探测指定 WPS 对象的成员、属性值与方法列表（只反射方法名，不返回枚举常量表）。表达式会被真实求值：一次只探一个表达式，不要探测未防御的原生 getter（如 Comment、SortFields、AutoFilter.Filters），已知会令 WPS 进程崩溃；反射不等于无副作用，表达式本身也能改文档。",
        parameters: {
          type: "object",
          properties: {
            expression: { type: "string", description: "要反射探测的表达式，如 'app', 'pres', 'pres.Slides.Item(1)', 'doc', 'wb'" },
            workbookName: { type: "string", description: "精确目标工作簿名" },
            documentName: { type: "string", description: "精确目标 Word 文档名" },
            presentationName: { type: "string", description: "精确目标 PPT 文稿名" },
            component: { type: "string", enum: ["word", "excel", "ppt"], description: "组件类型: 'word'|'excel'|'ppt'" }
          },
          required: ["expression"],
          additionalProperties: false
        }
      }
    },
    {
      type: "function",
      function: {
        name: "wps_reload_addon",
        description: "让当前 WPS 重新加载桥接加载项，使**刚部署的新构建立即生效**，无需退出并重启 WPS。改了加载项代码并完成部署后调用它；调用成功不代表新代码已生效，随后用 bridge_get_capabilities 或 bridge_diagnose 读回构建指纹核对。无参数。",
        parameters: {
          type: "object",
          properties: {},
          additionalProperties: false
        }
      }
    }
  ];
}
