import { defineConfig } from "tsup";

export default defineConfig([
  // 1. Electron 主进程
  {
    entry: {
      "main/index": "src/main/index.ts"
    },
    format: ["esm"],
    target: "node20",
    clean: false,
    sourcemap: true,
    // 打包产物压缩 + 混淆：去掉注释、缩短局部变量名。
    // 目的不只是省体积——**注释里写着大量设计理由与真机踩坑记录**，
    // 随包发出去等于把设计文档一起送人。
    // 注意：esbuild 默认**不重命名属性名**，`args?.workbookName` 这类访问不受影响。
    minify: true,
    external: ["electron", "ws", "@modelcontextprotocol/sdk"]
  },
  // 2. Electron 预加载脚本 (Preload, 输出为 CJS 以获得最大的 Electron 兼容性)
  {
    entry: {
      "preload/index": "src/preload/index.ts"
    },
    format: ["cjs"],
    target: "node20",
    clean: false,
    sourcemap: true,
    // 打包产物压缩 + 混淆：去掉注释、缩短局部变量名。
    // 目的不只是省体积——**注释里写着大量设计理由与真机踩坑记录**，
    // 随包发出去等于把设计文档一起送人。
    // 注意：esbuild 默认**不重命名属性名**，`args?.workbookName` 这类访问不受影响。
    minify: true,
    external: ["electron"]
  },
  // 3. MCP 独立命令行服务 (CLI, CJS 单文件零依赖打包)
  {
    entry: {
      "bridge/cli": "src/bridge/cli.ts"
    },
    format: ["cjs"],
    target: "node20",
    platform: "node",
    clean: false,
    sourcemap: true,
    // 打包产物压缩 + 混淆：去掉注释、缩短局部变量名。
    // 目的不只是省体积——**注释里写着大量设计理由与真机踩坑记录**，
    // 随包发出去等于把设计文档一起送人。
    // 注意：esbuild 默认**不重命名属性名**，`args?.workbookName` 这类访问不受影响。
    minify: true,
    noExternal: [/(.*)/]
  }
]);
