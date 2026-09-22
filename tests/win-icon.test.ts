/**
 * Windows 图标生成器的格式回归测试。
 *
 * 背景：exe 文件图标在资源管理器里显示空白，根因不是条目数量不足，而是**小于 256px 的条目用了 PNG 内嵌**
 * （Windows 只对 256×256 支持 PNG 内嵌，更小的尺寸必须是 BMP/DIB）。原脚本只校验"条目数 = 7"就放行，
 * 因此这里把"编码格式"变成可断言的回归项。
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { test } from 'node:test';
import zlib from 'node:zlib';

import {
  MAX_ENTRY_BYTES,
  PNG_MIN_SIZE,
  SIZES,
  buildIco,
  decodePng,
  encodePng,
  resizeRgba,
  rgbaToDib,
  verifyIco,
} from '../scripts/build-win-icon.mjs';

const root = path.resolve(import.meta.dirname, '..');
const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

// ---- 测试自带的独立 PNG 写入器（故意不复用被测模块的编码路径） ----

function crc32(buffer: Buffer): number {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, data: Buffer): Buffer {
  const name = Buffer.from(type, 'latin1');
  const size = Buffer.alloc(4);
  size.writeUInt32BE(data.length);
  const checksum = Buffer.alloc(4);
  checksum.writeUInt32BE(crc32(Buffer.concat([name, data])));
  return Buffer.concat([size, name, data, checksum]);
}

function paeth(a: number, b: number, c: number): number {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  if (pa <= pb && pa <= pc) return a;
  return pb <= pc ? b : c;
}

/** 用指定的行过滤器（0–4）把 RGBA 数据写成 PNG，用于验证解码器的逐行反过滤。 */
function pngWithFilter(width: number, height: number, rgba: Buffer, filter: number): Buffer {
  const stride = width * 4;
  const raw = Buffer.alloc((stride + 1) * height);
  let previous = Buffer.alloc(stride);
  for (let y = 0; y < height; y++) {
    const row = rgba.subarray(y * stride, (y + 1) * stride);
    const out = raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1));
    raw[y * (stride + 1)] = filter;
    for (let i = 0; i < stride; i++) {
      const left = i >= 4 ? row[i - 4] : 0;
      const up = previous[i];
      const upLeft = i >= 4 ? previous[i - 4] : 0;
      let value = row[i];
      if (filter === 1) value -= left;
      else if (filter === 2) value -= up;
      else if (filter === 3) value -= (left + up) >> 1;
      else if (filter === 4) value -= paeth(left, up, upLeft);
      out[i] = value & 255;
    }
    previous = row;
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  return Buffer.concat([
    PNG_SIGNATURE,
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', zlib.deflateSync(raw)),
    pngChunk('IEND', Buffer.alloc(0)),
  ]);
}

/** 确定性伪随机 RGBA，避免测试依赖随机源。 */
function sampleImage(width: number, height: number) {
  const rgba = Buffer.alloc(width * height * 4);
  let seed = 0x12345678;
  for (let i = 0; i < rgba.length; i++) {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff;
    rgba[i] = seed >>> 16;
  }
  return { width, height, rgba };
}

function icoEntries(buffer: Buffer) {
  const count = buffer.readUInt16LE(4);
  return Array.from({ length: count }, (_, index) => {
    const at = 6 + index * 16;
    const bytes = buffer.readUInt32LE(at + 8);
    const offset = buffer.readUInt32LE(at + 12);
    return {
      size: buffer.readUInt8(at) || 256,
      bits: buffer.readUInt16LE(at + 6),
      data: buffer.subarray(offset, offset + bytes),
    };
  });
}

// ---- 用例 ----

test('decodePng 能还原 0–4 号行过滤器编码的 PNG', () => {
  const image = sampleImage(7, 5);
  for (const filter of [0, 1, 2, 3, 4]) {
    const decoded = decodePng(pngWithFilter(image.width, image.height, image.rgba, filter));
    assert.equal(decoded.width, image.width, `过滤器 ${filter} 宽度`);
    assert.equal(decoded.height, image.height, `过滤器 ${filter} 高度`);
    assert.deepEqual(decoded.rgba, image.rgba, `过滤器 ${filter} 像素`);
  }
});

test('decodePng 能解码真实源图标 build/icon.png', () => {
  const source = path.join(root, 'build', 'icon.png');
  assert.ok(fs.existsSync(source), `缺少源图标 ${source}`);
  const decoded = decodePng(fs.readFileSync(source));
  assert.equal(decoded.width, decoded.height);
  assert.ok(decoded.width >= PNG_MIN_SIZE);
  assert.equal(decoded.rgba.length, decoded.width * decoded.height * 4);

  assert.equal(decoded.rgba[3], 0, '左上角应为全透明');
  let maxAlpha = 0;
  let minAlpha = 255;
  let transparent = 0;
  for (let i = 3; i < decoded.rgba.length; i += 4) {
    const alpha = decoded.rgba[i];
    if (alpha > maxAlpha) maxAlpha = alpha;
    if (alpha < minAlpha) minAlpha = alpha;
    if (alpha === 0) transparent++;
  }
  // 源图标是透明背景上的彩色图案：两端 alpha 都要出现，且透明区域占比不能太小
  assert.equal(minAlpha, 0);
  assert.equal(maxAlpha, 255);
  assert.ok(transparent > decoded.width * decoded.height * 0.1, `透明像素过少：${transparent}`);
});

test('encodePng 与 decodePng 往返逐字节一致', () => {
  const image = sampleImage(9, 4);
  assert.deepEqual(decodePng(encodePng(image)), image);
});

test('resizeRgba 按面积平均缩放，全透明区域不产生暗色', () => {
  // 4×4 四象限：左上红、右上绿、左下蓝、右下全透明
  const rgba = Buffer.alloc(4 * 4 * 4);
  const set = (x: number, y: number, color: [number, number, number, number]) =>
    color.forEach((value, channel) => (rgba[(y * 4 + x) * 4 + channel] = value));
  for (let y = 0; y < 4; y++) {
    for (let x = 0; x < 4; x++) {
      if (x < 2 && y < 2) set(x, y, [255, 0, 0, 255]);
      else if (x >= 2 && y < 2) set(x, y, [0, 255, 0, 255]);
      else if (x < 2) set(x, y, [0, 0, 255, 255]);
      else set(x, y, [0, 0, 0, 0]);
    }
  }

  const scaled = resizeRgba({ width: 4, height: 4, rgba }, 2);
  const pixel = (x: number, y: number) => Array.from(scaled.rgba.subarray((y * 2 + x) * 4, (y * 2 + x) * 4 + 4));
  assert.deepEqual(pixel(0, 0), [255, 0, 0, 255]);
  assert.deepEqual(pixel(1, 0), [0, 255, 0, 255]);
  assert.deepEqual(pixel(0, 1), [0, 0, 255, 255]);
  assert.deepEqual(pixel(1, 1), [0, 0, 0, 0]);
});

test('rgbaToDib 生成合法的 32bpp 自下而上 DIB 与 AND 掩码', () => {
  // 上排两点不透明，下排左点全透明
  const rgba = Buffer.from([
    10, 20, 30, 255, 40, 50, 60, 255,
    70, 80, 90, 0, 100, 110, 120, 255,
  ]);
  const dib = rgbaToDib(rgba, 2);

  assert.equal(dib.readUInt32LE(0), 40, 'biSize');
  assert.equal(dib.readInt32LE(4), 2, 'biWidth');
  assert.equal(dib.readInt32LE(8), 4, 'biHeight 应为两倍（XOR + AND）');
  assert.equal(dib.readUInt16LE(12), 1, 'biPlanes');
  assert.equal(dib.readUInt16LE(14), 32, 'biBitCount');
  assert.equal(dib.readUInt32LE(16), 0, 'biCompression 必须是 BI_RGB');
  assert.equal(dib.length, 40 + 2 * 2 * 4 + 4 * 2, 'DIB 总长度');

  // 第 0 行取源图最后一行（自下而上），BGRA 顺序
  assert.deepEqual(Array.from(dib.subarray(40, 44)), [90, 80, 70, 0]);
  assert.deepEqual(Array.from(dib.subarray(44, 48)), [120, 110, 100, 255]);

  // AND 掩码：行 0（源图下排）的 x=0 全透明 → 最高位置 1；其余为 0
  const andOffset = 40 + 2 * 2 * 4;
  assert.equal(dib[andOffset], 0b10000000, 'AND 掩码行 0');
  assert.equal(dib[andOffset + 4], 0, 'AND 掩码行 1');
});

test('buildIco 目录项字段正确且偏移连续', () => {
  const entries = [
    { size: 16, data: rgbaToDib(Buffer.alloc(16 * 16 * 4, 200), 16) },
    { size: 256, data: encodePng(sampleImage(256, 256)) },
  ];
  const ico = buildIco(entries);
  assert.equal(ico.readUInt16LE(0), 0);
  assert.equal(ico.readUInt16LE(2), 1);
  assert.equal(ico.readUInt16LE(4), 2);

  const parsed = icoEntries(ico);
  assert.equal(parsed[0].size, 16);
  assert.equal(parsed[1].size, 256, '256px 的宽高字节记 0');
  assert.equal(ico.readUInt8(6), 16);
  assert.equal(ico.readUInt8(6 + 16), 0, '256px 目录项宽字节应为 0');
  assert.deepEqual(parsed[0].data, entries[0].data);
  assert.deepEqual(parsed[1].data, entries[1].data);
});

test('verifyIco 接受「<256 用 DIB、256 用 PNG」的产物', () => {
  const entries = SIZES.map(size => {
    const image = resizeRgba(sampleImage(512, 512), size);
    return { size, data: size >= PNG_MIN_SIZE ? encodePng(image) : rgbaToDib(image.rgba, size) };
  });
  const report = verifyIco(buildIco(entries), SIZES);
  assert.equal(report.length, SIZES.length);
  for (const item of report) {
    assert.equal(item.encoding, item.size >= PNG_MIN_SIZE ? 'png' : 'dib', `${item.size}px 编码`);
    assert.ok(item.bytes <= MAX_ENTRY_BYTES, `${item.size}px 条目 ${item.bytes} 字节超过 PE 资源的 16 位长度上限`);
  }
});

test('buildIco 拒绝超过 64KB 的条目（PE 资源长度字段被截断的回归项）', () => {
  // electron-builder 写入 PE 资源时 dwBytesInRes 只保留低 16 位：128×128 的 DIB 实测 67624 字节会被写成 2088
  assert.throws(() => buildIco([{ size: 128, data: Buffer.alloc(MAX_ENTRY_BYTES + 1) }]), /截断/);
  assert.ok(!SIZES.includes(128), '128px 的 DIB 必然超过 64KB，不能出现在尺寸集合里');
  const largestDib = Math.max(...SIZES.filter(size => size < PNG_MIN_SIZE).map(size => rgbaToDib(Buffer.alloc(size * size * 4), size).length));
  assert.ok(largestDib <= MAX_ENTRY_BYTES, `最大 DIB 条目 ${largestDib} 字节`);
});

test('verifyIco 拒绝小于 256px 的 PNG 条目（本轮空白图标的回归项）', () => {
  const png16 = encodePng(sampleImage(16, 16));
  const ico = buildIco([{ size: 16, data: png16 }]);
  assert.throws(() => verifyIco(ico, [16]), /DIB/);
});

test('verifyIco 拒绝条目数、尺寸与位深不符的产物', () => {
  const dib = rgbaToDib(Buffer.alloc(16 * 16 * 4, 255), 16);
  const ico = buildIco([{ size: 16, data: dib }]);
  assert.throws(() => verifyIco(ico, SIZES), /条目数/);
  assert.throws(() => verifyIco(ico, [32]), /16x16/);
});
