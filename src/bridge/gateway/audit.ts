/**
 * 审计与回滚类处理器（P3.1：从 gateway.ts 的 executeTool switch 原样搬迁，逻辑未改写）。
 *
 * 共享依赖（callOffice / auditStore / bridgeServer / TargetLockStore / requestContext /
 * previewPath / MsOfficeDriver / currentHost 等）一律经 GatewayContext 传入；
 * 本模块不 import 这些执行器，避免与门面产生新的隐式耦合或回环。
 *
 * 每个处理器的前三行固定为：导出签名、ctx 依赖解构、以及原分支正文的第一行；
 * 正文承接原 switch 分支的内容，仅做逐行等量左移 4 个空格的纯缩进变换。
 */
import type { Handler } from "./types.js";

// 仅类型引用（纯类型，无运行时依赖）。

import { RangeData } from "../types.js";

export const rollback: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const auditId = args?.auditId;
    if (!auditId) throw new Error("缺少必要参数: auditId");

    const rec = auditStore.getRecordById(auditId);
    if (!rec) throw new Error(`找不到对应的审计记录: ${auditId}`);
    if (rec.status === "rolled_back") {
      return { success: true, message: "该记录已处于已撤销状态，无需重复操作" };
    }

    const res = await requestContext.run({ sessionId: currentSession(), host: rec.host || 'wps' }, async () => {
      const current = await callOffice<RangeData>('read_range', { workbookName: rec.workbookName, sheetName: rec.sheetName, address: rec.address, includeFormulas: true });
      if (!rec.afterSnapshot || JSON.stringify(current.formulas) !== JSON.stringify(rec.afterSnapshot.formulas) || JSON.stringify(current.values) !== JSON.stringify(rec.afterSnapshot.values)) {
        throw new Error('目标区域已有后续修改或快照缺失，已拒绝覆盖。请先核对当前内容。');
      }
      return callOffice('rollback_cells', { workbookName: rec.workbookName, sheetName: rec.sheetName, address: rec.address, snapshot: rec.beforeSnapshot });
    });

    auditStore.markRolledBack(auditId);

    return {
      success: true,
      message: `成功恢复区域 ${rec.address} 的值与公式`,
      // 说清范围：该记录只覆盖 patch_cells 写入的值/公式，样式、图表、结构变更既不在快照里、
      // 也不在本记录的回滚范围内。原文案"成功恢复区域的数据"被测试者误读为"整表还原"（ISS-47 / ISS-28）。
      restoredScope: { values: true, formulas: true, styles: false, charts: false, structure: false },
      result: res
    };
};

export const getAuditHistory: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const view = args?.view || "summary";
    const requestedLimit = args?.limit ?? 5;
    const limit = Math.max(1, Math.min(view === "detail" ? 5 : 50, Math.trunc(requestedLimit)));
    const { records, total } = auditStore.getRecordsWithTotal({
      limit,
      offset: args?.offset,
      workbookName: args?.workbookName,
      sheetName: args?.sheetName,
      clientName: args?.clientName,
      actionType: args?.actionType,
      status: args?.status,
      fromTimestamp: args?.fromTimestamp,
      toTimestamp: args?.toTimestamp
    });
    if (view === "ids") {
      return { records: records.map((r) => ({ id: r.id, timestamp: r.timestamp, status: r.status })), total };
    }
    if (view === "summary") {
      return {
        records: records.map((r) => ({
          id: r.id,
          host: r.host || "wps",
          timestamp: r.timestamp,
          clientName: r.clientName,
          actionType: r.actionType,
          description: r.description,
          workbookName: r.workbookName,
          sheetName: r.sheetName,
          address: r.address,
          modifiedCount: r.modifiedCount,
          status: r.status
        })),
        total
      };
    }
    return { records, total };
};

export const clearAuditHistory: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    auditStore.clearAllRecords();
    return { success: true, message: "已成功清空所有修改记录留痕" };
};

export const getAuditRecord: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.auditId) throw new Error("缺少必要参数: auditId");
    const record = auditStore.getRecordById(args.auditId);
    if (!record) throw new Error(`找不到对应的审计记录: ${args.auditId}`);
    return record;
};
