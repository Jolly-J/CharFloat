/**
 * 只读检查打包好的 Windows exe 里到底嵌了哪些图标资源。
 *
 * 用途：确认 electron-builder 经 rcedit 写入的 `RT_GROUP_ICON` 条目数**与每个条目的编码格式**。
 * 只看条目数会漏判——条目齐全但全部是 PNG 内嵌时，资源管理器仍然显示空白图标（见
 * [p5.2-windows-cross-build.md](p5.2-windows-cross-build.md)）。
 *
 * 依赖 `resedit`（electron-builder → app-builder-lib 的传递依赖，随 devDependencies 一起安装）。
 *
 * 用法：
 *   node docs/acceptance/2.1.0-p0p1/p5/inspect-pe-icons.mjs release/2.1.0/win/win-unpacked/"Office Agent Bridge.exe"
 *
 * 判定标准（对照原版 `electron.exe`）：
 *   - `RT_GROUP_ICON` 1 条，组内含 7 个尺寸条目；
 *   - 小于 256×256 的 `RT_ICON` 必须是 DIB（BITMAPINFOHEADER 40 字节 / bpp 32 / 高度为两倍）；
 *   - 仅 256×256 允许 PNG 内嵌。
 */
import fs from 'node:fs';
import * as ResEdit from 'resedit';

const file = process.argv[2];
if (!file) {
  console.error('用法：node inspect-pe-icons.mjs <Windows exe 路径>');
  process.exit(1);
}

const resource = ResEdit.NtExecutableResource.from(
  ResEdit.NtExecutable.from(fs.readFileSync(file), { ignoreCert: true }),
);

// resedit 把标准资源类型解析成编号（未知类型才保留字符串），两种形式都接受
const RT_ICON = [3, 'RT_ICON'];
const RT_GROUP_ICON = [14, 'RT_GROUP_ICON'];
const RT_VERSION = [16, 'RT_VERSION'];
const ofType = type => resource.entries.filter(entry => type.includes(entry.type));

const sizes = [];
console.log(`${file}\n`);
console.log('资源类型：');
for (const [name, type] of [['RT_ICON', RT_ICON], ['RT_GROUP_ICON', RT_GROUP_ICON], ['RT_VERSION', RT_VERSION]]) {
  const entries = ofType(type);
  console.log(`  ${name}: ${entries.length} 条${entries.length ? `（id ${entries.map(entry => entry.id).join(', ')}）` : ''}`);
}

if (ofType(RT_ICON).length === 0 || ofType(RT_GROUP_ICON).length === 0) {
  console.error('\n结论：exe 内没有图标资源，资源管理器只能显示默认图标。');
  process.exit(1);
}

console.log('\nRT_GROUP_ICON：');
for (const group of ofType(RT_GROUP_ICON)) {
  const data = Buffer.from(group.bin);
  const count = data.readUInt16LE(4);
  console.log(`  group id=${group.id} 语言=${group.lang} 条目数=${count}`);
  for (let index = 0; index < count; index++) {
    const at = 6 + index * 14;
    const width = data.readUInt8(at) || 256;
    const height = data.readUInt8(at + 1) || 256;
    const bits = data.readUInt16LE(at + 6);
    const bytes = data.readUInt32LE(at + 8);
    const iconId = data.readUInt16LE(at + 12);
    console.log(`    [${index}] ${width}x${height} bpp=${bits} ${bytes} 字节 → RT_ICON id=${iconId}`);
    sizes.push(width);
  }
}

console.log('\nRT_ICON 编码格式：');
let failures = 0;
for (const entry of ofType(RT_ICON)) {
  const data = Buffer.from(entry.bin);
  if (data.subarray(0, 8).toString('hex') === '89504e470d0a1a0a') {
    const width = data.readUInt32BE(16);
    const height = data.readUInt32BE(20);
    const allowed = width >= 256;
    if (!allowed) failures++;
    console.log(`  id=${entry.id} PNG ${width}x${height}${allowed ? '' : '  ← 非法：小于 256 的条目必须是 DIB'}`);
    continue;
  }
  const biSize = data.readUInt32LE(0);
  const width = data.readInt32LE(4);
  const height = data.readInt32LE(8);
  const bits = data.readUInt16LE(14);
  const ok = biSize === 40 && height === width * 2 && bits === 32;
  if (!ok) failures++;
  console.log(`  id=${entry.id} DIB ${width}x${width} biSize=${biSize} biHeight=${height} bpp=${bits}${ok ? '' : '  ← 非法 DIB 头'}`);
}

console.log(
  failures === 0
    ? `\n结论：${sizes.length} 个尺寸（${sizes.join('/')}）全部格式合法。`
    : `\n结论：${failures} 个条目格式非法，Windows 资源管理器会显示空白图标。`,
);
process.exit(failures === 0 ? 0 : 1);
