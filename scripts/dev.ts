import { createServer } from 'vite';
import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '..');

/**
 * 开发启动前必须重建的产物。
 *
 * 加载项源码已迁到 `wps-addon/src/**` 与 `office-addon/src/**`，部署入口 `addon-core.js` /
 * `taskpane.js` 是**构建生成物**。只跑 `build:main` 会让开发环境继续加载旧的部署入口——
 * 改了源码却"没生效"，或更糟：以为生效了其实跑的是旧代码。
 *
 * 目前两个构建脚本只做拼接、不监听文件变化，因此**改加载项源码后需要重新执行 `npm run dev`
 * 或单跑 `npm run build:addons`**，然后重新部署/重载加载项（WPS 侧重新打开加载项，
 * Microsoft 侧可用 `POST /api/v1/office/reload` 触发任务窗格热重载）。
 */
const DEV_BUILD_STEPS: { label: string; args: string[] }[] = [
  { label: '主进程与预加载脚本', args: ['run', 'build:main'] },
  { label: '加载项部署入口（WPS 与 Office.js）', args: ['run', 'build:addons'] }
];

function runStep(label: string, args: string[]): Promise<void> {
  console.log(`[Dev] 正在构建${label}...`);
  const child = spawn('npm', args, { cwd: root, stdio: 'inherit', shell: true });
  return new Promise<void>((resolve, reject) => {
    child.on('close', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`${label}构建失败，退出码: ${code}`));
    });
  });
}

async function startDev() {
  for (const step of DEV_BUILD_STEPS) await runStep(step.label, step.args);
  console.log('[Dev] 注意：加载项源码（*/src/**）改动不会自动重建，需重新执行 npm run dev 或 npm run build:addons。');

  console.log('[Dev] 正在启动 Vite 前端开发服务器（支持 HMR 热重载）...');
  const server = await createServer({
    configFile: path.join(root, 'vite.config.ts'),
    server: { port: 5173 }
  });
  await server.listen();

  const address = server.httpServer?.address();
  const port = typeof address === 'object' && address ? address.port : 5173;
  const devUrl = `http://localhost:${port}`;
  console.log(`[Dev] 前端服务就绪: ${devUrl}`);

  console.log('[Dev] 正在启动 Electron 桌面客户端窗口...');
  const electronProcess = spawn('npx', ['electron', '.'], {
    cwd: root,
    stdio: 'inherit',
    shell: true,
    env: {
      ...process.env,
      VITE_DEV_SERVER_URL: devUrl
    }
  });

  const cleanup = async () => {
    try {
      await server.close();
    } catch {}
    process.exit(0);
  };

  electronProcess.on('close', cleanup);
  process.on('SIGINT', cleanup);
  process.on('SIGTERM', cleanup);
}

startDev().catch((err) => {
  console.error('[Dev] 启动失败:', err);
  process.exit(1);
});
