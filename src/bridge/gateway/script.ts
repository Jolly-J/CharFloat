/**
 * 脚本执行与 API 反射类处理器（P3.1：从 gateway.ts 的 executeTool switch 原样搬迁，逻辑未改写）。
 *
 * 共享依赖（callOffice / auditStore / bridgeServer / TargetLockStore / requestContext /
 * previewPath / MsOfficeDriver / currentHost 等）一律经 GatewayContext 传入；
 * 本模块不 import 这些执行器，避免与门面产生新的隐式耦合或回环。
 *
 * 每个处理器的前三行固定为：导出签名、ctx 依赖解构、以及原分支正文的第一行；
 * 正文承接原 switch 分支的内容，仅做逐行等量左移 4 个空格的纯缩进变换。
 */
import type { GatewayContext, Handler } from "./types.js";

import { runProcess } from "../process-runner.js";

export const executeScript: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const { code, script, component, documentName, workbookName, presentationName, params } = args || {};
    const scriptCode = code || script;
    if (!scriptCode) throw new Error("缺少必要参数: code (要执行的原生 JavaScript 代码)");
    return await callOffice("execute_script", {
      code: scriptCode,
      component,
      documentName: documentName || (component === "word" ? TargetLockStore.resolve("word") : undefined),
      workbookName: workbookName || (component === "excel" ? TargetLockStore.resolve("excel") : undefined),
      presentationName: presentationName || (component === "ppt" ? TargetLockStore.resolve("ppt") : undefined),
      params
    });
};

export const inspectApi: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    const { expression, path, component, documentName, workbookName, presentationName, evaluate, maxMembers } = args || {};
    // ISS-89 的护栏参数必须转发到宿主：漏传等于护栏永远按默认走（与 ISS-106 同一类问题——
    // schema 声明了字段、网关却不转发，调用方以为生效了其实没到）。
    return await callOffice("inspect_api", {
      expression: expression || path || "app",
      component,
      documentName,
      workbookName,
      presentationName,
      evaluate,
      maxMembers
    });
};

export const reloadAddon: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    // 必须**逐个组件**发送 reload。
    // 原实现只发一次 `callOffice("reload", {})`，实测只送达一个组件（excel 换成了新构建，
    // word/ppt 仍是旧构建）——"部署后免重启生效"要做全，就不能只重载其中一个。
    const requested = args?.component ? [String(args.component)] : ["excel", "word", "ppt"];
    const results: any[] = [];
    for (const component of requested) {
      try {
        const r = await callOffice("reload", { component });
        results.push({ component, reloaded: true, result: r });
      } catch (e: any) {
        // 该组件没打开/未连接属正常情况，如实记录但不让整批失败
        results.push({ component, reloaded: false, error: e?.message ?? String(e) });
      }
    }
    const okCount = results.filter(r => r.reloaded).length;
    return {
      success: okCount > 0,
      reloadedComponents: results.filter(r => r.reloaded).map(r => r.component),
      results,
      message: okCount > 0
        ? `已触发 ${okCount} 个组件重新加载加载项（${results.filter(r => r.reloaded).map(r => r.component).join(" / ")}）；` +
          `因 index.html 的 <script src> 带构建指纹，重载会真正取到新构建。`
        : `没有组件接受重载请求：${results.map(r => `${r.component}(${r.error})`).join("；")}`
    };
};

export const evalCode: Handler = async (ctx) => {
  const { name, args, clientName, locks, callOffice, auditStore, MsOfficeDriver, TargetLockStore, bridgeServer, requestContext, currentHost, currentSession, previewPath, extractClipboardImageBase64 } = ctx;
    if (!args?.code) throw new Error("缺少必要参数: code");
    return await callOffice("eval_code", { code: args.code });
};

// 跨平台从系统剪贴板提取图片并转为 Base64 (兼容 macOS Swift 与 Windows PowerShell)。
// 原为 UniversalGateway 的静态方法：实现不使用 this，改为 ctx 形参后语义不变；门面仍暴露同名静态方法。
export async function extractClipboardImageBase64(ctx: GatewayContext): Promise<string> {
  if (process.platform === 'darwin') {
    const script = `import Cocoa
let pb = NSPasteboard.general
if let data = pb.data(forType: .png) { print(data.base64EncodedString()) }
else if let tiff = pb.data(forType: .tiff), let rep = NSBitmapImageRep(data: tiff), let png = rep.representation(using: .png, properties: [:]) { print(png.base64EncodedString()) }`;
    const output = await runProcess('swift', ['-'], script, 20000);
    if (output.length > 100) return output;
  } else if (process.platform === 'win32') {
    const result = await ctx.MsOfficeDriver.windows({ action: 'clipboard' });
    if (result?.imageBase64) return result.imageBase64;
  }
  throw new Error('无法提取预览图片。macOS 需可用的 Swift 命令行运行时；Windows 需桌面剪贴板。');
}
