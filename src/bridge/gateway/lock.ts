/**
 * 目标锁类处理器（WPS 网关锁）（P3.1：从 gateway.ts 的 executeTool switch 原样搬迁，逻辑未改写）。
 *
 * 共享依赖（callOffice / auditStore / bridgeServer / TargetLockStore / requestContext /
 * previewPath / MsOfficeDriver / currentHost 等）一律经 GatewayContext 传入；
 * 本模块不 import 这些执行器，避免与门面产生新的隐式耦合或回环。
 *
 * 每个处理器的前三行固定为：导出签名、ctx 依赖解构、以及原分支正文的第一行；
 * 正文承接原 switch 分支的内容，仅做逐行等量左移 4 个空格的纯缩进变换。
 */
import type { Handler } from "./types.js";

export const lockTargetDocument: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const { component, targetName } = args || {};
    if (!component || !["word", "excel", "ppt"].includes(component)) {
      throw new Error("缺少有效参数: component 必须为 'word', 'excel' 或 'ppt'");
    }
    if (!targetName) throw new Error("缺少必要参数: targetName (目标文档名称/路径)");
    TargetLockStore.lock(component, targetName);
    return {
      success: true,
      component,
      lockedTarget: targetName,
      allLocks: TargetLockStore.getLocks(),
      message: `已成功锁定 ${component.toUpperCase()} 目标文档为 [${targetName}]，后续所有操作将严格针对该文档，无视前台窗口切换`
    };
};

export const unlockTargetDocument: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const { component } = args || {};
    TargetLockStore.unlock(component);
    return {
      success: true,
      unlockedComponent: component || "all",
      allLocks: TargetLockStore.getLocks(),
      message: `已成功解除 ${component ? component.toUpperCase() : "全部"} 文档锁定`
    };
};

export const getLockedStatus: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const gatewayLocks = TargetLockStore.getLocks();
    // ISS-02：这里原来恒返回 `addonStatus: null`，而"当前到底打开了哪些文稿"只有这里能问——
    // 于是锁目标陈旧时调用方没有任何自救路径（自举死锁）。现在改为真的去问一次宿主：
    // 不传目标问摘要 → 拿到真实的已打开列表，并与锁目标比对。
    let openWorkbooks: any[] | null = null;
    let hostError: string | null = null;
    try {
      const summary: any = await callOffice("get_workspace_summary", {});
      openWorkbooks = Array.isArray(summary?.openWorkbooks) ? summary.openWorkbooks
        : summary?.workbookName ? [{ name: summary.workbookName, fullName: summary.fullName || summary.workbookName }] : [];
    } catch (error: any) {
      hostError = error?.message || String(error);
    }
    const names = new Set((openWorkbooks || []).map((w: any) => w?.name).filter(Boolean));
    const lockTargetExists = Object.fromEntries(
      Object.entries(gatewayLocks).map(([component, target]) => [component, {
        target,
        // 只对表格做名称比对：Word/PPT 的锁目标不在表格摘要里，null 表示"未核对"。
        exists: component === 'excel' ? names.has(String(target)) : null
      }])
    );
    return {
      gatewayLocks,
      addonStatus: {
        openWorkbooks,
        openWorkbookCount: openWorkbooks ? openWorkbooks.length : null,
        lockedTargets: gatewayLocks,
        lockTargetExists,
        note: "gatewayLocks 是本会话用 wps_lock_target_document 设的锁（目标可能来自更早的会话）；openWorkbooks 才是宿主当前真实打开的文稿。"
      },
      ...(hostError ? { hostError, hostErrorNote: "宿主不可用或未连接，openWorkbooks 为 null，无法核对锁目标是否存在。" } : {}),
      message: "当前文档锁定状态查询完成"
    };
};
