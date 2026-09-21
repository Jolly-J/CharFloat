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
    noExternal: [/(.*)/]
  }
]);
