const path = require('node:path');
module.exports = async context => {
  const { checkRelease } = await import('./lib/check-release.mjs');
  const resources = context.electronPlatformName === 'darwin'
    ? path.join(context.appOutDir, context.packager.appInfo.productFilename + '.app', 'Contents/Resources')
    : path.join(context.appOutDir, 'resources');
  const result = await checkRelease(path.join(resources, 'app.asar'), { allowPlainCli: context.electronPlatformName === 'win32' });
  console.log(`[protection] 包内检查通过，${result.protectedFiles} 个发布资源；不代表真实宿主验收。`);
};
