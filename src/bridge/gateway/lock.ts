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
import { BridgeError, isBridgeError } from "../errors.js";

/** 目标锁组件（与 TargetLockStore 的三类文档一致）。 */
export type LockComponent = "word" | "excel" | "ppt";

/** 本次调用中"由会话锁注入、而非调用方显式传入"的目标。 */
export interface InjectedLockTarget {
  component: LockComponent;
  target: string;
}

/**
 * 宿主在"目标名字在已打开文档里找不到"时的措辞（三个通道各一种，逐条对照源码登记）。
 *
 * - WPS 加载项 `wps-addon/src/shared.js`：`未在 WPS 中找到目标 Word 文档/演示文稿/工作簿 [x]。当前已打开: …`
 * - Windows 原生 COM `resources/office/common.ps1`：`Target must match exactly one open document: x`
 * - macOS JXA `office/ms-office-driver.ts`：`找不到指定目标文档`
 */
const TARGET_NOT_FOUND_PATTERNS: readonly RegExp[] = [
  /未在 WPS 中找到目标\s*(?:Word 文档|演示文稿|工作簿)/,
  /找不到指定目标文档/,
  /Target must match exactly one open document/i,
  /Specify a document name when zero or multiple documents are open/i
];

/** 失败原因是否为"目标名字在宿主里不存在"。 */
export function isTargetNotFoundError(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error ?? "");
  return TARGET_NOT_FOUND_PATTERNS.some(re => re.test(message));
}

/**
 * ISS-126：宿主重启（或文档被关闭）后残留的目标锁，会让本会话**后续所有**带锁工具都拿不到文档：
 * 锁指向一个已不存在的名字 → 宿主报"未找到目标文档"，而报错不指向"该重新锁目标"，
 * 使用者会以为工具坏了，绕法是手工解锁或每次显式传名。
 *
 * 处置：当且仅当①失败原因是"目标名字在宿主里不存在"、且②该目标是**由会话锁注入**（不是调用方
 * 显式传入）时——解除这个失效锁，并把错误改写为可操作提示。
 * 调用方显式传错名字时不动锁，避免把用户刚设好的锁误删。
 *
 * @returns 改写后的错误；不属于残留锁场景时返回 null（调用方应原样抛出原错误）。
 */
export function reconcileStaleTargetLock(
  error: unknown,
  injected: readonly InjectedLockTarget[],
  unlock: (component: LockComponent) => void
): Error | null {
  if (!injected.length || !isTargetNotFoundError(error)) return null;
  const detail = error instanceof Error ? error.message : String(error);
  for (const { component } of injected) unlock(component);
  const list = injected.map(t => `${t.component.toUpperCase()} [${t.target}]`).join("、");
  const message =
    `目标锁已失效并自动解除：${list} 在宿主里已不存在（宿主重启或该文档已关闭），本次调用未完成。` +
    `原错误：${detail}\n` +
    `下一步：① 调 wps_get_locked_status 看宿主当前真正打开了哪些文档；` +
    `② 用 wps_lock_target_document 把目标锁到正确的文档，或本次调用显式传 documentName/presentationName/workbookName；` +
    `③ 失效锁已清理，后续不带目标名的调用不会再被它带偏。`;
  // 保留原失败的分类（kind / channel / executed）：调用方与日志依赖它判断"能不能重试、是否已执行"。
  const rewritten = isBridgeError(error)
    ? new BridgeError(error.kind, message, { channel: error.channel, method: error.method, executed: error.executed, cause: error })
    : new Error(message, { cause: error });
  (rewritten as any).staleTargetLocks = injected.map(t => ({ ...t }));
  (rewritten as any).lockReconciled = true;
  return rewritten;
}

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
    // ISS-126：能核对的部分（表格）直接把"已失效的锁"列出来。带锁调用失败时桥接会**自动解除**
    // 失效锁并在错误里给出重新锁定指引；这里是把同一事实在查询路径上暴露，供调用方提前自救。
    const staleLockTargets = Object.entries(lockTargetExists as Record<string, { target: string; exists: boolean | null }>)
      .filter(([, info]) => info.exists === false)
      .map(([component, info]) => ({ component, target: info.target }));
    return {
      gatewayLocks,
      addonStatus: {
        openWorkbooks,
        openWorkbookCount: openWorkbooks ? openWorkbooks.length : null,
        lockedTargets: gatewayLocks,
        lockTargetExists,
        staleLockTargets,
        note: "gatewayLocks 是本会话用 wps_lock_target_document 设的锁（目标可能来自更早的会话）；openWorkbooks 才是宿主当前真实打开的文稿。锁往一个已不存在的文档时：任何带锁调用都会先失败、桥接会自动解除该失效锁并在错误里提示重新锁定（ISS-126）。",
        ...(staleLockTargets.length
          ? { staleLockNote: "以上锁目标在宿主当前打开的文稿里不存在（已核对表格）：可现在就 wps_unlock_target_document 清掉，或直接重新 wps_lock_target_document 锁到正确的文档。" }
          : {}),
        ...(Object.keys(gatewayLocks).length && gatewayLocks.word
          ? { wordLockNote: "Word/PPT 的锁目标无法用表格摘要核对（这里只列表格）；若 Word 工具报「未在 WPS 中找到目标 Word 文档」，那就是锁已失效，重新锁定即可（桥接会自动清理旧锁）。" }
          : {})
      },
      ...(hostError ? { hostError, hostErrorNote: "宿主不可用或未连接，openWorkbooks 为 null，无法核对锁目标是否存在。" } : {}),
      message: "当前文档锁定状态查询完成"
    };
};
