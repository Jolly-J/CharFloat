// 纯构建函数：RGBA 吉祥物合成到 24 位 BMP，不运行安装或写文件。
import { decodePng, resizeRgba } from '../build-win-icon.mjs';

export function buildSplashBitmap(png) {
  const width = 400, height = 260, stride = width * 3;
  const bitmap = Buffer.alloc(54 + stride * height);
  bitmap.write('BM'); bitmap.writeUInt32LE(bitmap.length, 2); bitmap.writeUInt32LE(54, 10);
  bitmap.writeUInt32LE(40, 14); bitmap.writeInt32LE(width, 18); bitmap.writeInt32LE(height, 22);
  bitmap.writeUInt16LE(1, 26); bitmap.writeUInt16LE(24, 28); bitmap.writeUInt32LE(stride * height, 34);
  const { rgba } = resizeRgba(decodePng(png), 220);
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
    let rgb = [233, 244, 255]; const sx = x - 90, sy = y - 20;
    if (sx >= 0 && sx < 220 && sy >= 0 && sy < 220) {
      const at = (sy * 220 + sx) * 4, alpha = rgba[at + 3] / 255;
      rgb = rgb.map((value, channel) => Math.round(value * (1 - alpha) + rgba[at + channel] * alpha));
    }
    const at = 54 + (height - 1 - y) * stride + x * 3;
    bitmap[at] = rgb[2]; bitmap[at + 1] = rgb[1]; bitmap[at + 2] = rgb[0];
  }
  return bitmap;
}
