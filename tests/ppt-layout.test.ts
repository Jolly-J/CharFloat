/**
 * PPT 布局与最近一次构建一致性测试（改造计划 P3.6）。
 *
 * 迁移前：从 `wps-addon/addon-core.js` 里按函数名**截取源码字符串**再 eval——
 * 脆的隐式契约，生成物一改排序或插入新代码就假失败。
 * 迁移后：
 *   1. 纯布局计算直接从 `wps-addon/src/ppt-layout.js` **import**（该模块可在 Node 直接运行）；
 *   2. 宿主适配器 `fitGeneratedPptShapes` 从 `wps-addon/src/ppt.js` **片段**加载到注入过布局函数的
 *      vm 上下文里执行——加载的是源码模块，不再对生成物做字符串切割。
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { execFileSync } from 'node:child_process';
import {
  pptPageSize, pptScaleForPage, pptCountTextUnits, pptEstimateLines, pptFitFontSize
} from '../wps-addon/src/ppt-layout.js';
import { UniversalGateway } from '../src/bridge/gateway.js';

/** 宿主适配器所在的源码片段；加载到注入了纯布局函数的上下文，供其自由变量解析。 */
const pptFragment = fs.readFileSync('wps-addon/src/ppt.js', 'utf8');
function adapterContext() {
  const ctx: any = vm.createContext({ console, pptPageSize, pptScaleForPage, pptCountTextUnits, pptEstimateLines, pptFitFontSize });
  vm.runInContext(pptFragment, ctx);
  return ctx;
}

function shape(text = '内容', size = 24) {
  return {
    Id: 1, Left: 50, Top: 30, Width: 620, Height: 50, HasTextFrame: true,
    TextFrame: {
      HasText: true, MarginLeft: 0, MarginRight: 0, MarginTop: 0, MarginBottom: 0,
      TextRange: { Text: text, Font: { Size: size } }
    }
  };
}

test('布局纯函数可直接运行，无需宿主', () => {
  const page = pptPageSize({ PageSetup: { SlideWidth: 960, SlideHeight: 540 } });
  assert.deepEqual(page, { pageWidth: 960, pageHeight: 540, unit: 'pt' });
  const scale = pptScaleForPage(page);
  assert.equal(scale.sx, 960 / 720);
  assert.equal(scale.sy, 540 / 405);
  assert.equal(scale.scale, Math.min(scale.sx, scale.sy));
  assert.equal(pptCountTextUnits('长内容abc'), 3 + 3 * 0.6);
  assert.equal(pptEstimateLines('单行', 24, 620), 1);
});

for (const [width, height] of [[720, 405], [960, 540], [1440, 810], [720, 540]] as const) {
  test(`generated content follows actual ${width} x ${height} page`, () => {
    const ctx = adapterContext();
    const sh = shape();
    const warnings: unknown[] = [];
    const page = ctx.pptPageSize({ PageSetup: { SlideWidth: width, SlideHeight: height } });
    ctx.fitGeneratedPptShapes({ Shapes: { Count: 1, Item: () => sh } }, 1, page, warnings);
    // 全部几何按真实页面比例映射，不允许固定 720x405 坐标直接落到其它尺寸
    assert.ok(Math.abs(sh.Left / width - 50 / 720) < 1e-12);
    assert.ok(Math.abs(sh.Width / width - 620 / 720) < 1e-12);
    assert.ok(Math.abs(sh.Top / height - 30 / 405) < 1e-12);
    assert.equal(sh.TextFrame.TextRange.Font.Size, 24 * Math.min(width / 720, height / 405));
    assert.equal(warnings.length, 0);
  });
}

test('invalid page size stops layout', () => {
  assert.throws(() => pptPageSize({ PageSetup: { SlideWidth: 0, SlideHeight: 405 } }), /页面尺寸/);
});

test('long text warns and existing shapes remain unchanged', () => {
  const ctx = adapterContext();
  const old = shape();
  const long = shape('长内容'.repeat(500), 16);
  const warnings: unknown[] = [];
  ctx.fitGeneratedPptShapes({ Shapes: { Count: 2, Item: (i: number) => i === 1 ? old : long } }, 2,
    { pageWidth: 1440, pageHeight: 810 }, warnings);
  assert.equal(old.Left, 50, 'firstIndex 之前的既有形状不得被改动');
  assert.equal(old.TextFrame.TextRange.Font.Size, 24, '既有形状字号不得被改动');
  assert.equal(long.TextFrame.TextRange.Font.Size, 24);
  assert.equal(warnings.length, 1, '溢出必须给出可定位的警告');
});

test('生成物与源码一致，且不含 import/export', () => {
  // 生成物必须能从源码重建：--check 模式只校验不写文件，漂移即退出码 1。
  execFileSync('node', ['scripts/build-wps-addon.mjs', '--check'], { stdio: 'pipe' });
  const artifact = fs.readFileSync('wps-addon/addon-core.js', 'utf8');
  assert.match(artifact.split('\n')[0], /请勿手改/, '生成物顶部必须有"请勿手改"声明');
  assert.ok(!/^\s*(import|export)\s/m.test(artifact), '构建必须剥掉测试用 export 块，产物不得含 import/export');
});

test('native script tools expose exact target names', () => {
  for (const name of ['wps_execute_script', 'wps_inspect_api']) {
    const tool = UniversalGateway.getOpenAiTools().find(t => t.function.name === name)!;
    for (const key of ['workbookName', 'documentName', 'presentationName']) {
      assert.ok(key in tool.function.parameters.properties);
    }
  }
});
