/**
 * 通用网关门面（P3.1 兼容门面）。
 *
 * 职责只剩四项：
 *   1. 目标锁注入（原逻辑原样保留）；
 *   2. `executeTool`：查表 → 组装上下文 → 调用处理器 → 未知工具报错；
 *   3. 保持既有公开静态方法签名不变：`executeTool` / `extractClipboardImageBase64` /
 *      `getOpenAiTools` / `getOpenApiSchema`（`getOpenAiTools` 已改为委托
 *      `tools/definitions/index.ts` 的 `gatewayToolDefinitions(activeWbHint)`）；
 *   4. `TargetLockStore` 按原路径继续导出（实现在 `gateway/locks.ts`）。
 *
 * 分支实现按类拆到 `gateway/*.ts`；原来的 switch 换成
 * `const HANDLERS: Record<string, Handler>` 显式映射表，`executeTool` 只做
 * "查表 → 调用 → 未知工具报错"。工具定义（原 `getOpenAiTools` 的 59 条数组）
 * 已按类拆到 `tools/definitions/*`（P2.3）。
 * ## 依赖方向
 *
 * 门面 → `gateway/*`（处理器）与 `tools/definitions/*`（纯数据定义）单向依赖。
 * 处理器不 import 门面；定义模块不 import 门面，也不 import `contracts/`、
 * `catalog.ts`、协议层任一模块，因此不产生回环；门面自身仍不 import
 * `catalog.ts` / 协议层任一模块（`tools/index.ts` 也不反向 import 门面）。
 */

import { callOffice } from "./office/adapter.js";
import { currentHost, currentSession, requestContext } from "./context.js";
import { previewPath } from "./runtime.js";
import { bridgeServer } from "./ws-server.js";
import { auditStore } from "./audit-store.js";
import { MsOfficeDriver } from "./office/ms-office-driver.js";

import { TargetLockStore } from "./gateway/locks.js";
import type { GatewayContext, Handler } from "./gateway/types.js";
import * as excel from "./gateway/excel.js";
import * as word from "./gateway/word.js";
import * as ppt from "./gateway/ppt.js";
import * as script from "./gateway/script.js";
import * as audit from "./gateway/audit.js";
import * as lock from "./gateway/lock.js";
import * as microsoft from "./gateway/microsoft.js";
import { gatewayToolDefinitions } from "./tools/definitions/index.js";

export { TargetLockStore };

/** 工具名 → 处理器 的显式注册表（原 switch 的等价物）。键集合由 `registeredGatewayBranches()` 对外暴露。 */
const HANDLERS: Record<string, Handler> = {
  // ---- 目标锁（WPS 网关锁） ----
  "wps_lock_target_document": lock.lockTargetDocument,
  "wps_unlock_target_document": lock.unlockTargetDocument,
  "wps_get_locked_status": lock.getLockedStatus,

  // ---- 脚本执行与 API 反射 ----
  "wps_execute_script": script.executeScript,
  "wps_inspect_api": script.inspectApi,

  // ---- Microsoft 原生驱动 ----
  "office_get_status": microsoft.getStatus,
  "office_lock_target": microsoft.lockTarget,
  "office_unlock_target": microsoft.unlockTarget,
  "office_execute_script": microsoft.executeScript,
  "office_capture_slide_preview": microsoft.captureSlidePreview,

  // ---- Excel（表格） ----
  "wps_get_workspace_summary": excel.getWorkspaceSummary,
  "wps_get_sheet_outline": excel.getSheetOutline,
  "wps_create_sheet": excel.createSheet,
  "wps_delete_sheet": excel.deleteSheet,
  "wps_get_style_token": excel.getStyleToken,
  "wps_format_text_segment": excel.formatTextSegment,
  "wps_configure_print_layout": excel.configurePrintLayout,
  "wps_export_sheet_pdf": excel.exportSheetPdf,
  "wps_set_sheet_view": excel.setSheetView,
  "wps_copy_range": excel.copyRange,
  "wps_manage_hyperlink": excel.manageHyperlink,
  "wps_manage_named_range": excel.manageNamedRange,
  "wps_manage_workbook_views": excel.manageWorkbookViews,
  "wps_manage_document_properties": excel.manageDocumentProperties,
  "wps_manage_table": excel.manageTable,
  "wps_manage_pictures": excel.managePictures,
  "wps_add_shape": excel.addShape,
  "wps_group_shapes": excel.groupShapes,
  "wps_ungroup_shapes": excel.ungroupShapes,
  "wps_set_shape_zorder": excel.setShapeZOrder,
  "wps_export_shape_image": excel.exportShapeImage,
  "wps_list_shapes": excel.listShapes,
  "wps_update_shape": excel.updateShape,
  "wps_create_workbook": excel.createWorkbook,
  "wps_clear_range": excel.clearRange,
  "wps_export_chart_image": excel.exportChartImage,
  "wps_auto_fit_columns": excel.autoFitColumns,
  "wps_read_range": excel.readRange,
  "wps_get_range_styles": excel.getRangeStyles,
  "wps_search_cells": excel.searchCells,
  "wps_patch_cells": excel.patchCells,
  "wps_format_cells": excel.formatCells,
  "wps_add_conditional_formatting": excel.addConditionalFormatting,
  "wps_freeze_panes": excel.freezePanes,
  "wps_modify_rows_columns": excel.modifyRowsColumns,
  "wps_add_chart": excel.addChart,
  "wps_get_charts": excel.getCharts,
  "wps_update_chart": excel.updateChart,
  "wps_delete_chart": excel.deleteChart,
  "wps_create_pivot_table": excel.createPivotTable,
  "wps_set_filter_and_sort": excel.setFilterAndSort,
  "wps_set_data_validation": excel.setDataValidation,
  "wps_manage_sheet": excel.manageSheet,
  "wps_manage_rows_and_columns": excel.manageRowsAndColumns,
  "wps_manage_cell_comments": excel.manageCellComments,
  "wps_find_and_replace": excel.findAndReplace,
  "wps_duplicate_sheet": excel.duplicateSheet,
  "wps_save_workbook": excel.saveWorkbook,
  "wps_capture_sheet_preview": excel.captureSheetPreview,

  // ---- Word（文字） ----
  "wps_word_create_document": word.createDocument,
  "wps_word_save_document": word.saveDocument,
  "wps_word_close_document": word.closeDocument,
  "wps_word_manage_content": word.manageContent,
  "wps_word_read_document": word.readDocument,
  "wps_word_write_content": word.writeContent,
  "wps_word_format_document": word.formatDocument,
  "wps_word_insert_table_of_contents": word.insertTableOfContents,
  "wps_word_manage_table": word.manageTable,
  "wps_word_review_and_comments": word.reviewAndComments,
  "wps_word_page_layout_and_watermark": word.pageLayoutAndWatermark,
  "wps_word_find_and_replace": word.findAndReplace,
  "wps_word_capture_preview": word.capturePreview,
  "wps_word_update_fields": word.updateFields,
  "wps_word_manage_content_controls": word.manageContentControls,

  // ---- PowerPoint（演示） ----
  "wps_ppt_configure_layout": ppt.configureSlideLayout,
  "wps_ppt_read_presentation": ppt.readPresentation,
  "wps_ppt_get_slide_shapes": ppt.getSlideShapes,
  "wps_ppt_generate_deck": ppt.generateDeck,
  "wps_ppt_manage_slides": ppt.manageSlides,
  "wps_ppt_manage_table": ppt.manageTable,
  "wps_ppt_add_business_cards": ppt.addBusinessCards,
  "wps_ppt_add_chart": ppt.insertNativeChart,
  "wps_ppt_insert_native_chart": ppt.insertNativeChart,
  "wps_ppt_manage_shapes_and_media": ppt.manageShapesAndMedia,
  "wps_ppt_capture_slide_preview": ppt.captureSlidePreview,

  // ---- 审计、回滚与收尾 ----
  "wps_rollback": audit.rollback,
  "wps_reload_addon": script.reloadAddon,
  "wps_get_audit_history": audit.getAuditHistory,
  "wps_clear_audit_history": audit.clearAuditHistory,
  "wps_get_audit_record": audit.getAuditRecord,
  "wps_eval_code": script.evalCode,
};


/**
 * 已注册的网关执行分支（真实注册表键，非注释文本）。
 *
 * 供一致性测试直接读取：测试**不得**再从源码里正则抓 `case "..."` 文本，
 * 否则一旦分支改成注册表形式，测试就退化成"校验注释"，而注释与实际执行可以脱节。
 */
export function registeredGatewayBranches(): string[] {
  return Object.keys(HANDLERS);
}

/**
 * "未经值快照记录的、可能改变目标身份"的操作（ISS-46）。
 *
 * 回滚的"是否有后续修改"判定原来只比内容：删表重建后写回**完全相同**的内容时，
 * 旧 auditId 会被放行并把别人的数据清空。这里登记这类操作的发生时间，
 * `wps_rollback` 会拒绝"记录时间之后发生过这类操作"的回滚。
 *
 * 范围刻意收窄：只登记会改变**工作表身份或数据位置**的操作
 * （脚本、清空、工作表增删改、行列插入删除、批量替换）。
 * 纯样式/图表/冻结/校验/批注类操作不改数据位置，不能因此封掉正常的
 * "先 patch、再排版、最后回滚值"流程。
 */
const IDENTITY_CHANGING_TOOLS: Record<string, string> = {
  "wps_execute_script": "任意原生脚本执行（可能重建或删除工作表、整体改写内容）",
  "wps_eval_code": "任意代码执行",
  "wps_inspect_api": "任意表达式探测（表达式可以带副作用）",
  "wps_clear_range": "清空区域",
  "wps_create_sheet": "新建工作表",
  "wps_delete_sheet": "删除工作表",
  "wps_duplicate_sheet": "克隆工作表",
  "wps_manage_sheet": "工作表改名/移动/保护等结构操作",
  "wps_modify_rows_columns": "插入/删除行列（既有数据位置会移动）",
  "wps_manage_rows_and_columns": "插入/删除行列（既有数据位置会移动）",
  "wps_find_and_replace": "批量查找替换（整体改写内容）",
  "office_execute_script": "原生脚本执行（可能改动文档结构或内容）"
};

/** 脚本/无目标类操作无法归属到具体工作簿时，登记为对所有工作簿生效（保守）。 */
const GLOBAL_SCOPE_TOOLS = new Set(["wps_execute_script", "wps_eval_code", "wps_inspect_api", "office_execute_script"]);

function recordIdentityChangingOperation(name: string, args: any) {
  const reason = IDENTITY_CHANGING_TOOLS[name];
  if (!reason) return;
  // M-9：脚本工具原先**无条件**登记为"未快照操作"，于是"写 → 用脚本读回核对 → 回滚"
  // 这个标准自检流程会被自己的读回动作挡住——不是回滚坏了，是守卫把手电筒也算成了脚印。
  //
  // 处置：调用方可以显式声明 `readOnly: true`（声明该脚本只读），此时不登记。
  // **这是调用方的声明，不是系统的保证**——声明了却真去写，回滚安全性由声明方负责。
  // 工具描述里已把这一点写明；真正写入的脚本不要声明 readOnly。
  if (args?.readOnly === true) return;
  const workbookName = GLOBAL_SCOPE_TOOLS.has(name) ? undefined : args?.workbookName;
  auditStore.markUntrackedMutation(workbookName, `${name}：${reason}`);
}

export class UniversalGateway {
  /**
   * 统一执行工具调用
   */
  public static async executeTool(
    name: string,
    args: any = {},
    clientName: string = "AI Agent",
    meta: { ignoredParams?: string[] } = {}
  ): Promise<any> {
    // 强隔离目标文档锁定注入：若已锁定文档且未显式指定（或指定为空），自动强制绑定锁定的文档
    let locks: { word?: string; excel?: string; ppt?: string };
    let targetSource: 'request' | 'session-lock' | 'none' | undefined;
    // ISS-126：记录"目标是从**会话锁**注入的、不是调用方显式传的"。只有这种目标在宿主里查不到时
    // 才能断定是残留锁（宿主重启/文档已关闭）；调用方显式传的名字查不到时不动用户的锁。
    const injectedLockTargets: lock.InjectedLockTarget[] = [];
    if (name.startsWith("wps_ppt_") || name.startsWith("ppt_")) {
      const explicitTarget = typeof args?.presentationName === 'string' && args.presentationName.trim() !== '';
      locks = { ppt: TargetLockStore.resolve("ppt", args?.presentationName) };
      args.presentationName = locks.ppt;
      if (!explicitTarget && locks.ppt) injectedLockTargets.push({ component: "ppt", target: locks.ppt });
    } else if (name.startsWith("wps_word_") || name.startsWith("word_")) {
      const explicitTarget = typeof args?.documentName === 'string' && args.documentName.trim() !== '';
      locks = { word: TargetLockStore.resolve("word", args?.documentName) };
      args.documentName = locks.word;
      if (!explicitTarget && locks.word) injectedLockTargets.push({ component: "word", target: locks.word });
    } else if (name.startsWith("wps_") && !name.includes("ppt") && !name.includes("word") && !name.includes("lock") && !name.includes("rollback")) {
      // 记录目标来源：调用方显式传的 workbookName 优先，否则回落到本会话锁（ISS-02 / ISS-77）。
      const explicit = typeof args?.workbookName === 'string' && args.workbookName.trim() !== '';
      locks = { excel: TargetLockStore.resolve("excel", args?.workbookName) };
      args.workbookName = locks.excel;
      targetSource = explicit ? 'request' : locks.excel ? 'session-lock' : 'none';
      if (!explicit && locks.excel) injectedLockTargets.push({ component: "excel", target: locks.excel });
    } else {
      locks = TargetLockStore.getLocks();
    }

    const handler = HANDLERS[name];
    if (!handler) throw new Error(`未知的 WPS 工具名称: ${name}`);

    let result: any;
    try {
      result = await handler({
        name, args, clientName, locks, targetSource, ignoredParams: meta.ignoredParams,
        callOffice, auditStore, MsOfficeDriver, TargetLockStore,
        bridgeServer, requestContext, currentHost, currentSession, previewPath,
        extractClipboardImageBase64: UniversalGateway.extractClipboardImageBase64,
      } satisfies GatewayContext);
    } catch (error) {
      // ISS-126：宿主重启后残留的目标锁指向已不存在的文档 → 本会话后续所有带锁工具都失败在
      // "未找到目标文档"，而报错不指向"该重新锁目标"。确认是"注入的锁目标不存在"时，
      // 自动解除失效锁并改写为可操作提示；其余失败原样抛出（不改用户的锁、不改错误语义）。
      const reconciled = lock.reconcileStaleTargetLock(error, injectedLockTargets, component => TargetLockStore.unlock(component));
      throw reconciled ?? error;
    }

    // 只在成功之后登记：失败的操作没有改变文档，不应影响后续回滚判定。
    recordIdentityChangingOperation(name, args);
    return result;
  }

  /**
   * 跨平台从系统剪贴板提取图片并转为 Base64 (兼容 macOS Swift 与 Windows PowerShell)
   */
  public static async extractClipboardImageBase64(): Promise<string> {
    return script.extractClipboardImageBase64({
      name: "", args: {}, clientName: "AI Agent", locks: {},
      callOffice, auditStore, MsOfficeDriver, TargetLockStore,
      bridgeServer, requestContext, currentHost, currentSession, previewPath,
      extractClipboardImageBase64: UniversalGateway.extractClipboardImageBase64,
    } satisfies GatewayContext);
  }

  /**
   * 生成通用于各大国产与国际大模型 (豆包、通义千问、DeepSeek、OpenAI) 的 tools 声明定义
   *
   * 定义本体在 `tools/definitions/**`；`activeWbHint` 原样透传（参与
   * `wps_get_workspace_summary` 的 description 与 `workbookName` 的描述文本）。
   */
  public static getOpenAiTools(activeWbHint?: string) {
    // 定义已按类迁到 tools/definitions/**（P2.3）；本方法保留原签名与 activeWbHint 分支语义，
    // 仅委托装配，输出与迁移前逐字节一致。
    return gatewayToolDefinitions(activeWbHint);
  }

  /**
   * 生成供 Dify / Coze 扣子 / GPTs 一键导入的 OpenAPI 3.0 规范
   */
  public static getOpenApiSchema(hostUrl: string = "http://127.0.0.1:19890") {
    return {
      openapi: "3.0.1",
      info: {
        title: "WPS Bridge Universal API",
        description: "让任意 AI Agent (豆包、千问、DeepSeek、Dify、Coze) 直接理解并操作当前打开的 WPS Excel 表格",
        version: "1.0.0"
      },
      servers: [
        {
          url: hostUrl,
          description: "本地 WPS Bridge 守护进程"
        }
      ],
      paths: {
        "/api/v1/tool/call": {
          post: {
            summary: "执行 WPS 操作工具",
            description: "统一执行 WPS 读写、大纲检索、选区定位、格式修改与撤销操作",
            operationId: "callWpsTool",
            requestBody: {
              required: true,
              content: {
                "application/json": {
                  schema: {
                    type: "object",
                    properties: {
                      name: {
                        type: "string",
                        description: "工具名称，如 wps_get_workspace_summary, wps_patch_cells, wps_read_range 等"
                      },
                      arguments: {
                        type: "object",
                        description: "传给该工具的具体参数"
                      },
                      clientName: {
                        type: "string",
                        description: "调用该接口的 Agent 名称（用于留痕显示），如 '豆包Agent', '通义千问', 'Dify'"
                      }
                    },
                    required: ["name"]
                  }
                }
              }
            },
            responses: {
              "200": {
                description: "操作成功返回",
                content: {
                  "application/json": {
                    schema: {
                      type: "object",
                      properties: {
                        success: { type: "boolean" },
                        data: {},
                        error: { type: "string" }
                      }
                    }
                  }
                }
              }
            }
          }
        },
        "/api/v1/status": {
          get: {
            summary: "获取当前 WPS 连接与表格状态",
            operationId: "getBridgeStatus",
            responses: {
              "200": {
                description: "状态信息"
              }
            }
          }
        }
      }
    };
  }
}
