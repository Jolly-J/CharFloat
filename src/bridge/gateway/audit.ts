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

/** 解析 "Sheet1!$A$1:$B$2" / "A1:B2" / "$A$1" 形式的区域地址；无法解析时返回 null。 */
function parseRangeAddress(address: unknown): { r1: number; c1: number; r2: number; c2: number } | null {
  if (typeof address !== 'string') return null;
  const cleaned = address.includes('!') ? address.slice(address.lastIndexOf('!') + 1) : address;
  const m = cleaned.replace(/\$/g, '').match(/^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/);
  if (!m) return null;
  const col = (letters: string) => letters.toUpperCase().split('').reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
  const c1 = col(m[1]);
  const r1 = Number(m[2]);
  const c2 = m[3] ? col(m[3]) : c1;
  const r2 = m[4] ? Number(m[4]) : r1;
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

/** 两个区域是否可能重叠；任一侧地址无法解析时返回 null（表示"判不了"，由调用方决定）。 */
function rangesOverlap(a: unknown, b: unknown): boolean | null {
  const ra = parseRangeAddress(a);
  const rb = parseRangeAddress(b);
  if (!ra || !rb) return null;
  return ra.r1 <= rb.r2 && rb.r1 <= ra.r2 && ra.c1 <= rb.c2 && rb.c1 <= ra.c2;
}

export const rollback: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const auditId = args?.auditId;
    if (!auditId) throw new Error("缺少必要参数: auditId");

    const rec = auditStore.getRecordById(auditId);
    if (!rec) throw new Error(`找不到对应的审计记录: ${auditId}`);
    if (rec.status === "rolled_back") {
      return { success: true, message: "该记录已处于已撤销状态，无需重复操作" };
    }

    // ISS-48：只留痕、不可回滚的记录（格式/条件格式/冻结/行列/图表等）必须明确拒绝，
    // 而不是走到下面"快照缺失"的模糊分支或被误判为"没有后续修改"。
    if (rec.rollbackable === false || !rec.beforeSnapshot || !rec.afterSnapshot) {
      throw new Error(
        `该记录不可回滚（操作类型 ${rec.actionType}）：它只登记了操作事实，没有值/公式快照。` +
        `wps_rollback 只能恢复 patch_cells 写入的值与公式；可用 wps_get_audit_record 查看记录内容，` +
        `样式/图表/结构类改动需要用宿主脚本或对应工具反向处理。`
      );
    }

    // ISS-46：回滚判定的第一条是**身份**，不是内容相等。
    // 原来的判定只比内容：别人删表重建、再写回完全相同的内容时，旧 auditId 会被放行并清空这批新数据。
    // 判据一：同一区域内是否还有**更晚**且仍未回滚的审计记录（说明该区域已被别人改过）。
    const laterOverlapping = auditStore.getRecordsWithTotal({ workbookName: rec.workbookName, sheetName: rec.sheetName, limit: 100 }).records
      .filter(other => other.id !== rec.id && other.timestamp > rec.timestamp && other.status === 'applied' && other.rollbackable !== false)
      .filter(other => rangesOverlap(other.address, rec.address) !== false)
      .map(other => `${other.id}(${other.actionType} ${other.address} ${new Date(other.timestamp).toISOString()})`);
    if (laterOverlapping.length) {
      throw new Error(
        `目标区域 ${rec.address} 在本记录之后还有 ${laterOverlapping.length} 条未回滚的修改记录：` +
        `${laterOverlapping.slice(0, 3).join('、')}。回滚会覆盖别人的改动，已拒绝；` +
        `请先处理更新的记录（按其 auditId 回滚），或改用 patch_cells 显式写回 wps_get_audit_record 中的旧快照。`
      );
    }

    // 判据二：本记录之后是否发生过"未经值快照记录、但可能改变工作表身份或数据位置"的操作
    // （任意脚本、清空区域、工作表增删改名、行列插入删除、批量替换）。这类操作没有留痕快照，
    // 无法证明目标还是同一次修改的对象，因此拒绝覆盖。
    const untracked = auditStore.getLastUntrackedMutation(rec.workbookName);
    if (untracked && untracked.at > rec.timestamp) {
      throw new Error(
        `本记录（${new Date(rec.timestamp).toISOString()}）之后，工作簿 [${rec.workbookName}] 上发生过未经快照记录的操作：` +
        `${untracked.reason}（${new Date(untracked.at).toISOString()}${untracked.scope === 'global' ? '，无法归属到具体工作簿' : ''}）。` +
        `此时即使区域内容与记录一致，也无法确认它仍是同一次修改的结果（例如工作表被删除后重建并写回相同内容），已拒绝覆盖。` +
        `如确需恢复，请用 wps_get_audit_record 读取该记录的 beforeSnapshot，再用 patch_cells 显式写回。`
      );
    }

    const res = await requestContext.run({ sessionId: currentSession(), host: rec.host || 'wps' }, async () => {
      // 判据三：目标工作表是否仍然存在。工作表被删除/重建时，宿主会报"找不到工作表"。
      try {
        await callOffice('get_sheet_outline', { workbookName: rec.workbookName, sheetName: rec.sheetName });
      } catch (error: any) {
        const message = String(error?.message || error);
        if (/找不到|未找到|不存在|没有|no such|not found|invalid sheet/i.test(message)) {
          throw new Error(`目标工作表 [${rec.sheetName}] 已不存在或已被重建（宿主返回：${message}），已拒绝回滚。`);
        }
        // 其它错误（超时、连接问题）不在这里下结论：下面的 read_range 会给出真实原因。
      }

      const current = await callOffice<RangeData>('read_range', { workbookName: rec.workbookName, sheetName: rec.sheetName, address: rec.address, includeFormulas: true });
      if (JSON.stringify(current.formulas) !== JSON.stringify(rec.afterSnapshot!.formulas) || JSON.stringify(current.values) !== JSON.stringify(rec.afterSnapshot!.values)) {
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
      identityChecks: ['无更晚的同区域记录', '无未审计的身份变更操作', '目标工作表仍存在', '区域内内容与修改后快照一致'],
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
          sessionId: r.sessionId,
          actionType: r.actionType,
          description: r.description,
          workbookName: r.workbookName,
          sheetName: r.sheetName,
          address: r.address,
          modifiedCount: r.modifiedCount,
          rollbackable: r.rollbackable ?? Boolean(r.afterSnapshot),
          status: r.status
        })),
        total,
        scopeNote: "rollbackable=true 的记录可用 wps_rollback 恢复值/公式；false 的只留痕、不可回滚。"
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
