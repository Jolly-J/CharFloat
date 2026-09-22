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
  "wps_clear_range": excel.clearRange,
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

  // ---- PowerPoint（演示） ----
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

export class UniversalGateway {
  /**
   * 统一执行工具调用
   */
  public static async executeTool(
    name: string,
    args: any = {},
    clientName: string = "AI Agent"
  ): Promise<any> {
    // 强隔离目标文档锁定注入：若已锁定文档且未显式指定（或指定为空），自动强制绑定锁定的文档
    let locks: { word?: string; excel?: string; ppt?: string };
    if (name.startsWith("wps_ppt_") || name.startsWith("ppt_")) {
      locks = { ppt: TargetLockStore.resolve("ppt", args?.presentationName) };
      args.presentationName = locks.ppt;
    } else if (name.startsWith("wps_word_") || name.startsWith("word_")) {
      locks = { word: TargetLockStore.resolve("word", args?.documentName) };
      args.documentName = locks.word;
    } else if (name.startsWith("wps_") && !name.includes("ppt") && !name.includes("word") && !name.includes("lock") && !name.includes("rollback")) {
      locks = { excel: TargetLockStore.resolve("excel", args?.workbookName) };
      args.workbookName = locks.excel;
    } else {
      locks = TargetLockStore.getLocks();
    }

    const handler = HANDLERS[name];
    if (!handler) throw new Error(`未知的 WPS 工具名称: ${name}`);

    return handler({
      name, args, clientName, locks,
      callOffice, auditStore, MsOfficeDriver, TargetLockStore,
      bridgeServer, requestContext, currentHost, currentSession, previewPath,
      extractClipboardImageBase64: UniversalGateway.extractClipboardImageBase64,
    } satisfies GatewayContext);
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
