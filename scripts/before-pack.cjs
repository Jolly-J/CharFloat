// electron-builder 的 beforePack 钩子（**按平台各调用一次**）。
//
// ## 为什么 Windows 要带一份明文兜底
// `cli.jsc` 是 V8 字节码，**同时绑定「V8 版本」与「编译平台」**。
// macOS 上编出来的字节码，到 Windows 上会被 V8 拒绝：
//   Error: Invalid or incompatible cached data (cachedDataRejected)
// 真机已证（2026-09-23）：Windows 包因此起不来，界面报「后台启动失败」，
// 而**同一份产物在 macOS 上正常**；`service.log` 是空的（V8 在原生层中止）。
//
// 在 macOS 上无法生成 Windows 的字节码（跑不了 PE 可执行文件），因此采用
// **加载器自带回退**：`dist/bridge/cli.cjs` 尝试加载字节码，失败且存在
// `cli-full.cjs` 时改用明文版。
//
// 所以 Windows 包必须带上 `cli-full.cjs`，且校验要放行它（`allowPlainCli`）。
// **代价（如实记录）**：Windows 侧保护降为「minify 混淆」，**没有字节码**。
// 要拿到同等级保护，需要一台 Windows 构建机在那边跑 build:bytecode 后再打包。
module.exports = async ({ electronPlatformName } = {}) => {
  const fs = require('node:fs');
  const path = require('node:path');

  const { verifyStage, STAGE, ROOT } = await import('./lib/release-protection.mjs');
  const isWin = electronPlatformName === 'win32';

  if (isWin) {
    const plain = path.join(ROOT, 'dist/bridge/cli-full.cjs');
    if (!fs.existsSync(plain)) {
      throw new Error('Windows 打包需要 dist/bridge/cli-full.cjs，但它不存在。请先跑 npm run build。');
    }
    const target = path.join(STAGE, 'dist/bridge/cli-full.cjs');
    fs.copyFileSync(plain, target);
    console.log(`[before-pack] Windows：已附带明文兜底 cli-full.cjs（${(fs.statSync(target).size / 1048576).toFixed(2)} MB）；Windows 侧为 minify 保护、无字节码`);
  } else {
    // macOS 用字节码，明文兜底不该进包（进了等于白送源码）
    const stray = path.join(STAGE, 'dist/bridge/cli-full.cjs');
    if (fs.existsSync(stray)) fs.unlinkSync(stray);
  }

  // 校验放在最后：Windows 时放行兜底文件，其余规则（清单、加载项、指纹）一律照旧
  verifyStage({ allowPlainCli: isWin });
};
