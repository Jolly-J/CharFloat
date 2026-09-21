import { createServer } from 'vite';
import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '..');

async function startDev() {
  console.log('[Dev] 正在编译主进程与预加载脚本...');
  const buildMain = spawn('npm', ['run', 'build:main'], {
    cwd: root,
    stdio: 'inherit',
    shell: true
  });

  await new Promise<void>((resolve, reject) => {
    buildMain.on('close', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`主进程编译失败，退出码: ${code}`));
    });
  });

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
