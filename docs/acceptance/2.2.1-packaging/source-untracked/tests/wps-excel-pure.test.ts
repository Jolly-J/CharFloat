import test from 'node:test';
import assert from 'node:assert/strict';
// @ts-ignore 无宿主依赖，构建时剥离测试导出。
import { normalize2DArray, excelColumnName } from '../wps-addon/src/excel/pure.js';

test('空区域矩阵维度完整，修改一行不污染其他行', () => {
  const matrix = normalize2DArray(null, 2, 3);
  assert.deepEqual(matrix, [[null, null, null], [null, null, null]]);
  matrix[0][0] = '写入';
  assert.equal(matrix[1][0], null);
  assert.deepEqual(normalize2DArray(undefined, 0, 2), []);
});
test('矩阵转换保留零、布尔值、空字符串和原有二维数据', () => {
  for (const value of [0, false, '']) assert.deepEqual(normalize2DArray(value, 1, 1), [[value]]);
  assert.deepEqual(normalize2DArray([0, '内容'], 1, 2), [[0, '内容']]);
  const ragged = [[0], ['内容', null]];
  assert.equal(normalize2DArray(ragged, 2, 2), ragged, '搬迁不得悄悄填充或复制原二维矩阵');
});
test('列名覆盖字母进位与 Excel 最后一列，保留现有输入归一规则', () => {
  for (const [index, name] of [[1, 'A'], [26, 'Z'], [27, 'AA'], [52, 'AZ'], [53, 'BA'], [702, 'ZZ'], [703, 'AAA'], [16384, 'XFD']] as const) {
    assert.equal(excelColumnName(index), name);
  }
  assert.equal(excelColumnName('27'), 'AA');
  assert.equal(excelColumnName(26.9), 'Z');
  for (const value of [0, -1, null, '非法']) assert.equal(excelColumnName(value), '');
});
