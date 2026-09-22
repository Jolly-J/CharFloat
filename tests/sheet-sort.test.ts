/**
 * 排序读回校验的纯函数测试（问题台账 ISS-38）。
 *
 * 为什么单独测它：排序"写完了"必须在宿主里**读回校验**才敢返回 success，
 * 而校验函数一旦漏比较（比如漏掉最后一行）就会"恒真"，保护直接失效——
 * 上一轮盘点有子代理据此提出 `isSortedByRules` 存在 off-by-one 的质疑，
 * 这里用可复现的用例把它钉死。
 */
import assert from 'node:assert/strict';
import { test } from 'node:test';

import { compareCellValues, isSortedByRules } from '../wps-addon/src/sheet-sort.js';

const asc = (colIndex = 1) => [{ colIndex, order: 'asc' }];
const desc = (colIndex = 1) => [{ colIndex, order: 'desc' }];

test('compareCellValues：数字按数值比较，不走字典序', () => {
  assert.equal(compareCellValues(2, 10), -1, '2 应小于 10（字典序会得出相反结论）');
  assert.equal(compareCellValues('10', '9'), 1, '纯数字字符串按数值比较：10 > 9（字典序会得出相反的 -1）');
  assert.equal(compareCellValues(5, 5), 0);
});

test('compareCellValues：空值最小，文本按字典序', () => {
  assert.equal(compareCellValues(null, 0), -1);
  assert.equal(compareCellValues('', 'a'), -1);
  assert.equal(compareCellValues('a', 'a'), 0);
  assert.equal(compareCellValues('a', 'b'), -1);
});

test('isSortedByRules：已排序的升序/降序返回 true', () => {
  assert.equal(isSortedByRules([['h'], ['a'], ['b'], ['c']], asc()), true);
  assert.equal(isSortedByRules([['h'], [3], [2], [1]], desc()), true);
});

test('isSortedByRules：中间相邻行逆序返回 false', () => {
  assert.equal(isSortedByRules([['h'], ['a'], ['c'], ['b'], ['d']], asc()), false);
});

test('isSortedByRules 必须比较**最后一对**相邻行（"漏最后一行"回归项）', () => {
  // 只有最后一对逆序：表头 + a, c, b
  assert.equal(isSortedByRules([['h'], ['a'], ['c'], ['b']], asc()), false);
  // 降序镜像：表头 + c, b, d —— 最后一对 (b, d) 在降序下逆序
  assert.equal(isSortedByRules([['h'], ['c'], ['b'], ['d']], desc()), false);
});

test('isSortedByRules：首行按表头跳过，表头与首行数据"逆序"不得误判', () => {
  // 表头 "zzz" 远大于首行数据 "a"；若把表头也拿去比较就会误报 false
  assert.equal(isSortedByRules([['zzz'], ['a'], ['b']], asc()), true);
});

test('isSortedByRules：少于 3 行（表头 + 1 行数据）无需排序，返回 true', () => {
  assert.equal(isSortedByRules([['h'], ['b']], asc()), true);
  assert.equal(isSortedByRules([['h']], asc()), true);
});

test('isSortedByRules：多列规则，前一列相等时看后一列', () => {
  const matrix = [['k', 'v'], ['x', 2], ['x', 1]];
  assert.equal(isSortedByRules(matrix, [{ colIndex: 1, order: 'asc' }, { colIndex: 2, order: 'asc' }]), false);
  assert.equal(isSortedByRules(matrix, [{ colIndex: 1, order: 'asc' }, { colIndex: 2, order: 'desc' }]), true);
});

test('isSortedByRules：规则为空或非法列号时不误判为已排序', () => {
  assert.equal(isSortedByRules([['h'], ['b'], ['a']], []), true, '没有规则时无可校验，返回 true');
  assert.equal(isSortedByRules([['h'], ['b'], ['a']], [{ colIndex: 0, order: 'asc' }]), false, ' colIndex 小于 1 属非法');
});

test('isSortedByRules：二维数组缺失时不抛错', () => {
  assert.equal(isSortedByRules(null as never, asc()), true);
  assert.equal(isSortedByRules([], asc()), true);
});
