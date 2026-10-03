// 从原始 App 图标生成工具栏资源；不修改 App/侧栏原图，不增加图像依赖。
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { decodePng, resizeRgba, encodePng } from './build-win-icon.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const RIBBON_ICONS = [['ribbon-icon.png', 32], ['ribbon-icon@2x.png', 64]];

export function buildWpsRibbonIcons(root = ROOT, checkOnly = false) {
  const source = decodePng(fs.readFileSync(path.join(root, 'resources/icon.png')));
  if (source.width !== source.height || source.width < 64) throw new Error('App 图标须为至少 64px 的正方形原图');
  for (const [name, size] of RIBBON_ICONS) {
    // 预乘 alpha 面积平均，避免直接大图缩小造成轮廓丢失和透明边缘暗边。
    const bytes = encodePng(resizeRgba(source, size));
    const target = path.join(root, 'wps-addon', name);
    const matches = fs.existsSync(target) && fs.readFileSync(target).equals(bytes);
    if (checkOnly && !matches) throw new Error(`工具栏图标与源图不一致: wps-addon/${name}`);
    if (!checkOnly && !matches) fs.writeFileSync(target, bytes);
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  buildWpsRibbonIcons(ROOT, process.argv.includes('--check'));
  console.log('WPS 工具栏图标检查通过：32px 与 64px @2x。');
}
