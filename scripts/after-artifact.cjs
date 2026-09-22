const fs = require('node:fs');
  const path = require('node:path');
module.exports = async event => {
  if (!event.file.endsWith('.zip')) return;
  const { checkRelease } = await import('./lib/check-release.mjs');
  // 按产物名判断平台：Windows 包允许带明文兜底 cli-full.cjs（见 before-pack.cjs 的长注释）
    const allowPlainCli = /win/i.test(path.basename(event.file));
    const result = await checkRelease(event.file, { allowPlainCli });
  fs.writeFileSync(event.file + '.protection.json', JSON.stringify(result, null, 2) + '\n');
  console.log(`[protection] ZIP 检查通过，SHA-256 ${result.sha256}`);
};
