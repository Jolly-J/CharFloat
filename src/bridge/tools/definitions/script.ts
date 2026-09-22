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
        description: "【WPS 图灵级超级脚本引擎】直接向当前 WPS 运行实例执行原生 JavaScript 自动化代码。已自动绑定锁定的文档上下文: app(WPS宿主), doc(当前Word文档), wb(当前Excel工作簿), pres(当前PPT演示文稿), wps(全局运行时), params(自定义参数)。专用工具未覆盖时使用本工具，例如 Excel 可编辑矢量形状、连接线、组合及已有图表属性修改，PPT 精细布局。先用 wps_inspect_api 或只读脚本检查目标 API，再执行并读回验证；没有专用工具不等于不支持。返回普通 JSON，勿返回宿主对象。",
        parameters: {
          type: "object",
          properties: {
            code: { type: "string", description: "要执行的原生 JavaScript 代码。支持 async/await。" },
            workbookName: { type: "string", description: "精确目标工作簿名" },
            documentName: { type: "string", description: "精确目标 Word 文档名" },
            presentationName: { type: "string", description: "精确目标 PPT 文稿名" },
            component: { type: "string", enum: ["word", "excel", "ppt"], description: "指定组件类型，不传则自动适配" },
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
        description: "【运行时 API 反射探测器】在运行时自省探测指定 WPS 对象的成员、属性值与方法列表，供 Agent 自主探索与调试原生 API。",
        parameters: {
          type: "object",
          properties: {
            expression: { type: "string", description: "要反射探测的表达式，如 'app', 'pres', 'pres.Slides.Item(1)', 'doc', 'wb'" },
            workbookName: { type: "string", description: "精确目标工作簿名" },
            documentName: { type: "string", description: "精确目标 Word 文档名" },
            presentationName: { type: "string", description: "精确目标 PPT 文稿名" },
            component: { type: "string", enum: ["word", "excel", "ppt"], description: "组件类型" }
          },
          required: ["expression"],
          additionalProperties: false
        }
      }
    }
  ];
}
