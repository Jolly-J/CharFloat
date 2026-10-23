// 用法：node scripts/build-windows-bootstrap.mjs <已验证Windows ZIP> <APP源码完整SHA> <引导源码完整SHA>
// 仅写 dist/windows-bootstrap；随后 electron-builder 的 portable 目标生成用户权限自解压 EXE。
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
import { createHash } from 'node:crypto';
import { Writable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { openZip, fileHash, safeName, validateManifest, PRODUCT } from '../windows-bootstrap/core.mjs';
import { checkRelease } from './lib/check-release.mjs';
const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const [zipArg, sourceCommit, bootstrapCommit] = process.argv.slice(2);
if (!zipArg || !/^[a-f0-9]{40}$/.test(sourceCommit) || !/^[a-f0-9]{40}$/.test(bootstrapCommit)) throw new Error('必须提供实际 ZIP 和两份完整源码 SHA');
const payload = path.resolve(zipArg), version = JSON.parse(fs.readFileSync(path.join(root, 'package.json'))).version;
await checkRelease(payload, { allowPlainCli: true });
const zip = await openZip(payload), files = [];
await new Promise((resolve, reject) => {
  zip.on('error', reject); zip.on('end', resolve);
  zip.on('entry', entry => { (async () => {
    const name = safeName(entry.fileName);
    if (!entry.fileName.endsWith('/')) {
      const stream = await new Promise((res, rej) => zip.openReadStream(entry, (e, s) => e ? rej(e) : res(s)));
      const hash = createHash('sha256'); let size = 0;
      await pipeline(stream,new Writable({write(chunk,_encoding,callback){hash.update(chunk);size+=chunk.length;callback();}}));
      files.push({ name, size, sha256: hash.digest('hex') });
    }
    zip.readEntry();
  })().catch(e => { zip.close(); reject(e); }); }); zip.readEntry();
});
const manifest = validateManifest({ product: PRODUCT, version, sourceCommit, bootstrapCommit, payloadSha256: await fileHash(payload), files: files.sort((a,b) => a.name.localeCompare(b.name)), totalBytes: files.reduce((sum,f) => sum+f.size,0) });
const output = path.join(root, 'dist/windows-bootstrap'); fs.mkdirSync(output, { recursive: true });
const app = path.join(output, 'app'); fs.rmSync(app, { recursive: true, force: true }); fs.mkdirSync(app);
await build({ entryPoints: [path.join(root,'windows-bootstrap/main.mjs')], bundle: true, format: 'cjs', platform: 'node', target: 'node22', external: ['electron'], minify: true, outfile: path.join(app,'main.cjs') });
for (const file of ['preload.cjs','index.html','style.css','renderer.js']) fs.copyFileSync(path.join(root,'windows-bootstrap',file),path.join(app,file));
fs.mkdirSync(path.join(app,'assets'));
for (const target of ['hello','work','done']) fs.copyFileSync(path.join(root,'windows-bootstrap/assets',target+'.png'),path.join(app,'assets',target+'.png'));
fs.writeFileSync(path.join(app,'package.json'),JSON.stringify({ name:'charfloat-user-setup',version,main:'main.cjs',description:'字浮 CharFloat 当前用户配置引导',author:'CharFloat',private:true }));
fs.writeFileSync(path.join(output,'payload-manifest.json'),JSON.stringify(manifest));
const config = { extends: null, appId:'com.charfloat.user-setup',productName:'字浮 CharFloat 配置引导',electronVersion:JSON.parse(fs.readFileSync(path.join(root,'node_modules/electron/package.json'))).version,
  directories:{ app,output:path.join(output,'artifacts'),buildResources:path.join(root,'build') }, files:['**/*'],extraResources:[{from:payload,to:'payload.zip'},{from:path.join(output,'payload-manifest.json'),to:'payload-manifest.json'}],
  // app-builder-lib 25.1.8 实现中 true 才不生成固定 UNPACK_DIR_NAME（其 d.ts 注释与实现相反）。
  npmRebuild:false,asar:true,artifactName:`CharFloat-${version}-win-x64.exe`,win:{target:[{target:'portable',arch:['x64']}],icon:path.join(root,'build/icon.ico'),requestedExecutionLevel:'asInvoker'},portable:{requestExecutionLevel:'user',unpackDirName:true},publish:null };
fs.writeFileSync(path.join(output,'builder.json'),JSON.stringify(config,null,2)+'\n');
console.log(JSON.stringify({version,sourceCommit,bootstrapCommit,payloadSha256:manifest.payloadSha256,files:files.length,totalBytes:manifest.totalBytes,config:'dist/windows-bootstrap/builder.json'},null,2));
