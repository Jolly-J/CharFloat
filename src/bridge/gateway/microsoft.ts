/**
 * Microsoft Office 原生驱动类处理器（P3.1：从 gateway.ts 的 executeTool switch 原样搬迁，逻辑未改写）。
 *
 * 共享依赖（callOffice / auditStore / bridgeServer / TargetLockStore / requestContext /
 * previewPath / MsOfficeDriver / currentHost 等）一律经 GatewayContext 传入；
 * 本模块不 import 这些执行器，避免与门面产生新的隐式耦合或回环。
 *
 * 每个处理器的前三行固定为：导出签名、ctx 依赖解构、以及原分支正文的第一行；
 * 正文承接原 switch 分支的内容，仅做逐行等量左移 4 个空格的纯缩进变换。
 */
import type { Handler } from "./types.js";

export const getStatus: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    return await MsOfficeDriver.getStatus();
};

export const lockTarget: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const { component, targetName } = args || {};
    if (!component || !["word", "excel", "ppt"].includes(component)) {
      throw new Error("缺少有效参数: component 必须为 'word', 'excel' 或 'ppt'");
    }
    if (!targetName) throw new Error("缺少必要参数: targetName (Microsoft Office 目标文档名称)");
    MsOfficeDriver.lockTarget(component, targetName);
    return {
      success: true,
      component,
      lockedTarget: targetName,
      allLocks: MsOfficeDriver.getLockedTargets(),
      message: `已成功锁定 Microsoft ${component.toUpperCase()} 目标文档为 [${targetName}]`
    };
};

export const unlockTarget: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const { component } = args || {};
    MsOfficeDriver.unlockTarget(component);
    return {
      success: true,
      unlockedComponent: component || "all",
      allLocks: MsOfficeDriver.getLockedTargets(),
      message: `已成功解除 Microsoft ${component ? component.toUpperCase() : "全部"} 文档锁定`
    };
};

export const executeScript: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const { component, script, code, targetName, params } = args || {};
    if (!component || !["word", "excel", "ppt"].includes(component)) {
      throw new Error("缺少有效参数: component 必须为 'word', 'excel' 或 'ppt'");
    }
    const scriptCode = script || code;
    if (!scriptCode) throw new Error("缺少必要参数: script (要执行的原生脚本代码)");
    return await MsOfficeDriver.executeScript(component, scriptCode, targetName, params);
};

export const captureSlidePreview: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const { slideIndex, presentationName } = args || {};
    const res = await MsOfficeDriver.capturePptSlide(Number(slideIndex) || 1, presentationName);
    return {
      success: true,
      slideIndex: Number(slideIndex) || 1,
      imagePath: res.imagePath,
      imageBase64: res.imageBase64,
      message: `已成功生成 Microsoft PowerPoint 第 ${slideIndex || 1} 页高保真快照`
    };
};
