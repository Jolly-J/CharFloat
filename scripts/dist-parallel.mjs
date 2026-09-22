// 并行构建两个平台的包。
//
// 为什么：`electron-builder --mac --win` 是**串行**的，一个平台打包时另一个干等；
// 而打包本身以 I/O 为主（读几千个小文件、写 asar、压 zip），单进程吃不满多核。
// 两个平台写到 `release/<版本>/mac` 与 `release/<版本>/win`，**目录不重叠**，可安全并行。
//
// 用法：node scripts/dist-parallel.mjs [--mac|--win]  （不带参数=两个都出）
import { spawn } from 'node:child_process';

const argv = process.argv.slice(2);
const want = argv.filter((a) => a === '--mac' || a === '--win');
const targets = want.length ? want : ['--mac', '--win'];
// --no-sign：跳过 macOS 代码签名。
// 实测瓶颈就在这里：mac 204s vs win 43s，而 electron-builder 自身 CPU 只用了 3 秒——
// `codesign` 在**等 Apple 的时间戳服务器**（联网），所以 CPU 闲置、时间全花在等待上。
// 自己测试用的包不需要签名，跳过可把 mac 那段降到与 win 相当。
const noSign = argv.includes('--no-sign');
if (noSign) console.log('[dist:parallel] --no-sign：跳过 macOS 代码签名（产物未签名，双击会被 Gatekeeper 拦，右键打开即可）');

function run(flag) {
  return new Promise((resolve) => {
    const t0 = Date.now();
    const env = { ...process.env };
    // 关掉身份自动发现 = electron-builder 不做 codesign
    if (noSign) env.CSC_IDENTITY_AUTO_DISCOVERY = 'false';
    const p = spawn('npx', ['electron-builder', flag], { stdio: ['ignore', 'pipe', 'pipe'], env });
    let out = '';
    p.stdout.on('data', (d) => { out += d.toString(); });
    p.stderr.on('data', (d) => { out += d.toString(); });
    p.on('close', (code) => {
      const secs = ((Date.now() - t0) / 1000).toFixed(0);
      console.log(`[dist:parallel] ${flag} 完成，${code === 0 ? '成功' : '失败'}，耗时 ${secs}s`);
      if (code !== 0) console.log(out.split('\n').slice(-15).join('\n'));
      resolve(code === 0);
    });
  });
}

const t0 = Date.now();
console.log(`[dist:parallel] 并行构建：${targets.join(' + ')}`);
const results = await Promise.all(targets.map(run));
const total = ((Date.now() - t0) / 1000).toFixed(0);
console.log(`[dist:parallel] 全部完成，总耗时 ${total}s`);
process.exit(results.every(Boolean) ? 0 : 1);
