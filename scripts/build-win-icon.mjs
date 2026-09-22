/**
 * 生成 Windows 可用的多尺寸图标 `build/icon.ico`（纯 Node，无外部依赖、跨平台）。
 *
 * ## 为什么条目格式比条目数量更重要
 *
 * Windows 资源管理器取 exe 文件图标时按视图尺寸挑选 `RT_GROUP_ICON` 里的条目，再要求 shell
 * 解码对应 `RT_ICON`。**小于 256×256 的条目必须是 BMP(DIB) 格式**；PNG 内嵌只对 256×256 有效。
 * 若所有条目都是 PNG，资源管理器拿不到可解码的小图标，就显示空白/默认图标——即使条目数量、
 * 尺寸集合、图标画面本身全都正确。
 *
 * 对照实测（Electron 35.7.5 原版 `electron.exe` 与改造前的本仓库 exe）：
 *
 * | 尺寸    | 原版 electron.exe（在 Windows 上显示正常） | 改造前（显示空白） |
 * |---------|--------------------------------------------|--------------------|
 * | 16×16   | DIB 1320B                                  | PNG 732B           |
 * | 32×32   | DIB 5160B                                  | PNG 1834B          |
 * | 48×48   | DIB 11560B                                 | PNG 3237B          |
 * | 256×256 | PNG 18963B                                 | PNG 40753B         |
 *
 * 因此本脚本：**< 256 的尺寸写 32bpp DIB（BITMAPINFOHEADER 40 字节 + 自下而上 BGRA + 1bpp AND 掩码），
 * 256×256 写 PNG**，与原版 Electron 的资源结构一致。写完后再解析自己的输出做自检
 * （[`verifyIco`]），格式不对直接抛错，不允许静默产出坏图标。
 *
 * 用法：`npm run build:win-icon`（只写 `build/icon.ico`）。
 */
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = path.resolve(import.meta.dirname, '..');
const source = path.join(root, 'build', 'icon.png');
const target = path.join(root, 'build', 'icon.ico');

/** 需要生成的尺寸集合（覆盖资源管理器的小图标、中图标、大图标与超大图标视图）。 */
export const SIZES = [16, 24, 32, 48, 64, 96, 256];

/** 达到该尺寸的条目内嵌 PNG；小于该尺寸的条目必须用 DIB。 */
export const PNG_MIN_SIZE = 256;

/**
 * 单个 ICO 条目的字节数上限。
 *
 * 已证实原因：electron-builder 把 ICO 写进 PE 资源时（rcedit 路径），`GRPICONDIRENTRY.dwBytesInRes`
 * 被截断成 16 位——实测 128×128 的 DIB（67624 = 0x10828 字节）在 exe 的 `RT_GROUP_ICON` 里变成了
 * 0x0828 = 2088 字节。资源本身完整，但组目录里的长度错了，Windows 按该长度解析会拿不到这张图。
 * 因此**不生成超过 64KB 的条目**：128×128 的 DIB 必然超标，改用 96×96（38440 字节），
 * 128px 视图由 96 或 256 缩放得到。
 */
export const MAX_ENTRY_BYTES = 0xffff;

const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
const CHANNELS_BY_COLOR_TYPE = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function paeth(a, b, c) {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  if (pa <= pb && pa <= pc) return a;
  return pb <= pc ? b : c;
}

function clampByte(value) {
  return value < 0 ? 0 : value > 255 ? 255 : value;
}

/**
 * 解码 8 位、非隔行的 PNG 为 RGBA。
 *
 * @param {Buffer} buffer PNG 文件内容
 * @returns {{width: number, height: number, rgba: Buffer}}
 */
export function decodePng(buffer) {
  if (!buffer.subarray(0, 8).equals(PNG_SIGNATURE)) throw new Error('不是 PNG 文件');
  let offset = 8;
  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = 0;
  let interlace = 0;
  let palette = null;
  let transparency = null;
  const idat = [];
  let sawEnd = false;

  while (offset + 8 <= buffer.length) {
    const length = buffer.readUInt32BE(offset);
    const type = buffer.subarray(offset + 4, offset + 8).toString('latin1');
    const data = buffer.subarray(offset + 8, offset + 8 + length);
    const expected = crc32(buffer.subarray(offset + 4, offset + 8 + length));
    const actual = buffer.readUInt32BE(offset + 8 + length);
    if (expected !== actual) throw new Error(`PNG 块 ${type} CRC 校验失败`);
    if (type === 'IHDR') {
      width = data.readUInt32BE(0);
      height = data.readUInt32BE(4);
      bitDepth = data[8];
      colorType = data[9];
      interlace = data[12];
    } else if (type === 'PLTE') {
      palette = Buffer.from(data);
    } else if (type === 'tRNS') {
      transparency = Buffer.from(data);
    } else if (type === 'IDAT') {
      idat.push(Buffer.from(data));
    } else if (type === 'IEND') {
      sawEnd = true;
      break;
    }
    offset += 12 + length;
  }

  if (!width || !height) throw new Error('PNG 缺少 IHDR');
  if (!sawEnd) throw new Error('PNG 缺少 IEND');
  if (bitDepth !== 8) throw new Error(`仅支持 8 位 PNG，实际 ${bitDepth} 位`);
  if (interlace !== 0) throw new Error('不支持隔行扫描（Adam7）PNG');
  const channels = CHANNELS_BY_COLOR_TYPE[colorType];
  if (!channels) throw new Error(`不支持的 PNG 颜色类型 ${colorType}`);
  if (colorType === 3 && !palette) throw new Error('调色板 PNG 缺少 PLTE 块');

  const stride = width * channels;
  const raw = zlib.inflateSync(Buffer.concat(idat));
  if (raw.length < (stride + 1) * height) throw new Error('PNG 像素数据长度不足');

  const lines = Buffer.alloc(stride * height);
  let previous = Buffer.alloc(stride);
  for (let y = 0; y < height; y++) {
    const filter = raw[y * (stride + 1)];
    const rowStart = y * (stride + 1) + 1;
    const current = lines.subarray(y * stride, (y + 1) * stride);
    raw.copy(current, 0, rowStart, rowStart + stride);
    if (filter === 1) {
      for (let i = channels; i < stride; i++) current[i] = (current[i] + current[i - channels]) & 255;
    } else if (filter === 2) {
      for (let i = 0; i < stride; i++) current[i] = (current[i] + previous[i]) & 255;
    } else if (filter === 3) {
      for (let i = 0; i < stride; i++) {
        const left = i >= channels ? current[i - channels] : 0;
        current[i] = (current[i] + ((left + previous[i]) >> 1)) & 255;
      }
    } else if (filter === 4) {
      for (let i = 0; i < stride; i++) {
        const left = i >= channels ? current[i - channels] : 0;
        const up = previous[i];
        const upLeft = i >= channels ? previous[i - channels] : 0;
        current[i] = (current[i] + paeth(left, up, upLeft)) & 255;
      }
    } else if (filter !== 0) {
      throw new Error(`未知的 PNG 行过滤器 ${filter}`);
    }
    previous = current;
  }

  const rgba = Buffer.alloc(width * height * 4);
  for (let i = 0, p = 0; i < width * height; i++, p += channels) {
    const at = i * 4;
    if (colorType === 6) {
      rgba[at] = lines[p];
      rgba[at + 1] = lines[p + 1];
      rgba[at + 2] = lines[p + 2];
      rgba[at + 3] = lines[p + 3];
    } else if (colorType === 2) {
      rgba[at] = lines[p];
      rgba[at + 1] = lines[p + 1];
      rgba[at + 2] = lines[p + 2];
      rgba[at + 3] = 255;
    } else if (colorType === 0) {
      rgba[at] = rgba[at + 1] = rgba[at + 2] = lines[p];
      rgba[at + 3] = 255;
    } else if (colorType === 4) {
      rgba[at] = rgba[at + 1] = rgba[at + 2] = lines[p];
      rgba[at + 3] = lines[p + 1];
    } else {
      const index = lines[p];
      rgba[at] = palette[index * 3];
      rgba[at + 1] = palette[index * 3 + 1];
      rgba[at + 2] = palette[index * 3 + 2];
      rgba[at + 3] = transparency && index < transparency.length ? transparency[index] : 255;
    }
  }

  return { width, height, rgba };
}

/**
 * 用面积平均（box filter）缩放 RGBA 图像。透明像素在预乘空间参与平均，避免边缘出现暗边。
 *
 * @param {{width: number, height: number, rgba: Buffer}} image
 * @param {number} size 目标边长（正方形）
 */
export function resizeRgba(image, size) {
  const { width, height, rgba } = image;
  if (!Number.isInteger(size) || size <= 0) throw new Error(`非法目标尺寸 ${size}`);
  const out = Buffer.alloc(size * size * 4);
  const scaleX = width / size;
  const scaleY = height / size;

  for (let y = 0; y < size; y++) {
    const y0 = Math.floor(y * scaleY);
    const y1 = Math.min(height, Math.max(y0 + 1, Math.ceil((y + 1) * scaleY)));
    for (let x = 0; x < size; x++) {
      const x0 = Math.floor(x * scaleX);
      const x1 = Math.min(width, Math.max(x0 + 1, Math.ceil((x + 1) * scaleX)));
      let r = 0;
      let g = 0;
      let b = 0;
      let alphaSum = 0;
      let count = 0;
      for (let sy = y0; sy < y1; sy++) {
        for (let sx = x0; sx < x1; sx++) {
          const at = (sy * width + sx) * 4;
          const alpha = rgba[at + 3] / 255;
          r += rgba[at] * alpha;
          g += rgba[at + 1] * alpha;
          b += rgba[at + 2] * alpha;
          alphaSum += alpha;
          count++;
        }
      }
      const at = (y * size + x) * 4;
      out[at] = alphaSum > 0 ? clampByte(Math.round(r / alphaSum)) : 0;
      out[at + 1] = alphaSum > 0 ? clampByte(Math.round(g / alphaSum)) : 0;
      out[at + 2] = alphaSum > 0 ? clampByte(Math.round(b / alphaSum)) : 0;
      out[at + 3] = clampByte(Math.round((alphaSum / count) * 255));
    }
  }

  return { width: size, height: size, rgba: out };
}

function pngChunk(type, data) {
  const name = Buffer.from(type, 'latin1');
  const size = Buffer.alloc(4);
  size.writeUInt32BE(data.length);
  const checksum = Buffer.alloc(4);
  checksum.writeUInt32BE(crc32(Buffer.concat([name, data])));
  return Buffer.concat([size, name, data, checksum]);
}

/**
 * 把 RGBA 图像编码为 8 位 RGBA 非隔行 PNG（用于 256×256 条目）。
 *
 * @param {{width: number, height: number, rgba: Buffer}} image
 */
export function encodePng(image) {
  const { width, height, rgba } = image;
  const stride = width * 4;
  const raw = Buffer.alloc((stride + 1) * height);
  for (let y = 0; y < height; y++) {
    raw[y * (stride + 1)] = 0;
    rgba.copy(raw, y * (stride + 1) + 1, y * stride, (y + 1) * stride);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  return Buffer.concat([
    PNG_SIGNATURE,
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', zlib.deflateSync(raw, { level: 9 })),
    pngChunk('IEND', Buffer.alloc(0)),
  ]);
}

/**
 * 把 RGBA 图像编码为 ICO 条目所用的 32bpp DIB：
 * `BITMAPINFOHEADER`(40) + 自下而上 BGRA 像素 + 1bpp AND 掩码（行按 4 字节对齐）。
 *
 * @param {Buffer} rgba
 * @param {number} size
 */
export function rgbaToDib(rgba, size) {
  const xorStride = size * 4;
  const xorSize = xorStride * size;
  const andStride = Math.ceil(size / 32) * 4;
  const andSize = andStride * size;

  const header = Buffer.alloc(40);
  header.writeUInt32LE(40, 0);            // biSize
  header.writeInt32LE(size, 4);           // biWidth
  header.writeInt32LE(size * 2, 8);       // biHeight = XOR + AND
  header.writeUInt16LE(1, 12);            // biPlanes
  header.writeUInt16LE(32, 14);           // biBitCount
  header.writeUInt32LE(0, 16);            // biCompression = BI_RGB
  header.writeUInt32LE(xorSize + andSize, 20);

  const xor = Buffer.alloc(xorSize);
  const and = Buffer.alloc(andSize);
  for (let y = 0; y < size; y++) {
    const sourceRow = size - 1 - y; // DIB 自下而上
    for (let x = 0; x < size; x++) {
      const from = (sourceRow * size + x) * 4;
      const to = y * xorStride + x * 4;
      xor[to] = rgba[from + 2];     // BGRA
      xor[to + 1] = rgba[from + 1];
      xor[to + 2] = rgba[from];
      xor[to + 3] = rgba[from + 3];
      if (rgba[from + 3] < 128) and[y * andStride + (x >> 3)] |= 0x80 >> (x & 7);
    }
  }

  return Buffer.concat([header, xor, and]);
}

/**
 * 按 ICO 规范拼装文件。
 *
 * @param {{size: number, data: Buffer}[]} entries
 */
export function buildIco(entries) {
  for (const entry of entries) {
    if (entry.data.length > MAX_ENTRY_BYTES) {
      throw new Error(
        `${entry.size}x${entry.size} 条目 ${entry.data.length} 字节，超过 ${MAX_ENTRY_BYTES} 字节上限：` +
          'PE 资源里的 GRPICONDIRENTRY.dwBytesInRes 会被截断成 16 位，Windows 将读不到这张图。',
      );
    }
  }

  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0);              // reserved
  header.writeUInt16LE(1, 2);              // type = icon
  header.writeUInt16LE(entries.length, 4);

  const directory = Buffer.alloc(16 * entries.length);
  let offset = header.length + directory.length;
  entries.forEach((entry, index) => {
    const at = index * 16;
    directory.writeUInt8(entry.size >= 256 ? 0 : entry.size, at);
    directory.writeUInt8(entry.size >= 256 ? 0 : entry.size, at + 1);
    directory.writeUInt8(0, at + 2);       // 调色板数
    directory.writeUInt8(0, at + 3);       // reserved
    directory.writeUInt16LE(1, at + 4);    // color planes
    directory.writeUInt16LE(32, at + 6);   // bits per pixel
    directory.writeUInt32LE(entry.data.length, at + 8);
    directory.writeUInt32LE(offset, at + 12);
    offset += entry.data.length;
  });

  return Buffer.concat([header, directory, ...entries.map(entry => entry.data)]);
}

/**
 * 解析 ICO 并校验每个条目的尺寸与编码格式，返回可打印的自检报告。
 *
 * 这是本脚本的验收关卡：**小于 [`PNG_MIN_SIZE`] 的条目必须是 DIB，否则抛错**。
 *
 * @param {Buffer} buffer
 * @param {number[]} expectedSizes
 * @returns {{size: number, encoding: 'dib' | 'png', bytes: number}[]}
 */
export function verifyIco(buffer, expectedSizes = SIZES) {
  if (buffer.readUInt16LE(0) !== 0) throw new Error('ICO reserved 字段不为 0');
  if (buffer.readUInt16LE(2) !== 1) throw new Error('ICO type 不是 1（图标）');
  const count = buffer.readUInt16LE(4);
  if (count !== expectedSizes.length) throw new Error(`ICO 条目数 ${count}，期望 ${expectedSizes.length}`);

  const report = [];
  for (let index = 0; index < count; index++) {
    const at = 6 + index * 16;
    const width = buffer.readUInt8(at) || 256;
    const height = buffer.readUInt8(at + 1) || 256;
    const bits = buffer.readUInt16LE(at + 6);
    const bytes = buffer.readUInt32LE(at + 8);
    const offset = buffer.readUInt32LE(at + 12);
    const data = buffer.subarray(offset, offset + bytes);
    if (data.length !== bytes) throw new Error(`条目 ${index} 数据越界`);
    if (bytes > MAX_ENTRY_BYTES) {
      throw new Error(`条目 ${index}（${width}x${height}）${bytes} 字节超过 ${MAX_ENTRY_BYTES} 字节上限，写入 PE 资源时会被截断`);
    }

    const expected = expectedSizes[index];
    if (width !== expected || height !== expected) throw new Error(`条目 ${index} 是 ${width}x${height}，期望 ${expected}x${expected}`);
    if (bits !== 32) throw new Error(`条目 ${index} 位深为 ${bits}，期望 32`);

    const isPng = data.subarray(0, 8).equals(PNG_SIGNATURE);
    if (isPng) {
      const pngWidth = data.readUInt32BE(16);
      const pngHeight = data.readUInt32BE(20);
      if (pngWidth !== expected || pngHeight !== expected) throw new Error(`条目 ${index} 内嵌 PNG 为 ${pngWidth}x${pngHeight}`);
      if (expected < PNG_MIN_SIZE) {
        throw new Error(`条目 ${index}（${expected}x${expected}）是 PNG 内嵌：小于 ${PNG_MIN_SIZE} 的尺寸在 Windows 上不显示，必须用 DIB`);
      }
      report.push({ size: expected, encoding: 'png', bytes });
      continue;
    }

    const biSize = data.readUInt32LE(0);
    const biWidth = data.readInt32LE(4);
    const biHeight = data.readInt32LE(8);
    const biPlanes = data.readUInt16LE(12);
    const biBitCount = data.readUInt16LE(14);
    const biCompression = data.readUInt32LE(16);
    if (biSize !== 40 || biWidth !== expected || biHeight !== expected * 2 || biPlanes !== 1 || biBitCount !== 32 || biCompression !== 0) {
      throw new Error(`条目 ${index}（${expected}x${expected}）DIB 头非法：biSize=${biSize} ${biWidth}x${biHeight} planes=${biPlanes} bpp=${biBitCount} comp=${biCompression}`);
    }
    const needed = 40 + expected * expected * 4 + Math.ceil(expected / 32) * 4 * expected;
    if (data.length < needed) throw new Error(`条目 ${index}（${expected}x${expected}）DIB 数据不足：${data.length} < ${needed}`);
    report.push({ size: expected, encoding: 'dib', bytes });
  }

  return report;
}

function main() {
  if (!fs.existsSync(source)) throw new Error(`缺少源图标：${path.relative(root, source)}`);
  const image = decodePng(fs.readFileSync(source));
  if (image.width !== image.height) throw new Error(`源图标必须是正方形，实际 ${image.width}x${image.height}`);
  if (image.width < PNG_MIN_SIZE) throw new Error(`源图标至少 ${PNG_MIN_SIZE}px，实际 ${image.width}px`);

  const entries = SIZES.map(size => {
    const scaled = resizeRgba(image, size);
    const data = size >= PNG_MIN_SIZE ? encodePng(scaled) : rgbaToDib(scaled.rgba, size);
    return { size, data };
  });

  const ico = buildIco(entries);
  const report = verifyIco(ico, SIZES); // 自检不通过就不落盘
  fs.writeFileSync(target, ico);

  console.log(`已生成 ${path.relative(root, target)}：${report.length} 个尺寸，${ico.length} 字节`);
  for (const item of report) {
    console.log(`  ${String(item.size).padStart(3)}x${String(item.size).padEnd(3)} ${item.encoding === 'dib' ? 'DIB(BMP) 自下而上 BGRA + AND 掩码' : 'PNG 内嵌'}  ${item.bytes} 字节`);
  }
}

const invoked = process.argv[1] ? fileURLToPath(pathToFileURL(fs.realpathSync(process.argv[1]))) : '';
if (invoked && invoked === fileURLToPath(import.meta.url)) main();
