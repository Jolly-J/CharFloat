import test from 'node:test';
import assert from 'node:assert/strict';
import { normalizeOfficeRequest, normalizeOfficeResponse } from '../src/bridge/office/normalizer.js';

test('normalizeOfficeRequest transforms flat styling parameters to Office.js structures', () => {
  const req = normalizeOfficeRequest('wps_format_cells', {
    sheetName: 'Sheet1',
    address: 'B2:D2',
    bold: true,
    fontColor: '#FF0000',
    backgroundColor: '#00FF00',
    numberFormat: '0.00%',
    horizontalAlignment: 'center'
  });

  assert.equal(req.method, 'format_cells');
  assert.equal(req.params.bold, true);
  assert.equal(req.params.font.bold, true);
  assert.equal(req.params.font.color, '#FF0000');
  assert.equal(req.params.fill.color, '#00FF00');
  assert.equal(req.params.alignment.horizontal, 'center');
  assert.equal(req.params.numberFormat, '0.00%');
});

test('normalizeOfficeResponse normalizes patch_cells result to standard PatchResult', () => {
  const res = normalizeOfficeResponse('wps_patch_cells', {
    success: true,
    address: 'A1:C3',
    sheetName: 'Sheet1',
    rowCount: 3,
    columnCount: 3,
    modifiedCount: 9
  }, { address: 'A1:C3', sheetName: 'Sheet1' });

  assert.equal(res.success, true);
  assert.equal(res.modifiedCount, 9);
  assert.equal(res.address, 'A1:C3');
  assert.equal(res.sheetName, 'Sheet1');
});

test('normalizeOfficeResponse normalizes get_workspace_summary', () => {
  const res = normalizeOfficeResponse('excel_get_workspace_summary', {
    hasOpenWorkbook: true,
    workbookName: '工作簿1.xlsx',
    activeSheet: 'MCP测试',
    sheetCount: 2,
    sheets: [{ name: 'Sheet1', position: 1 }, { name: 'MCP测试', position: 2 }]
  }, { workbookName: '工作簿1.xlsx' });

  assert.equal(res.hasOpenWorkbook, true);
  assert.equal(res.activeSheetName, 'MCP测试');
  assert.equal(res.sheetCount, 2);
  assert.equal(res.sheets.length, 2);
});

test('normalizeOfficeRequest routes save_workbook to save', () => {
  const req = normalizeOfficeRequest('wps_save_workbook', { workbookName: 'Test.xlsx' });
  assert.equal(req.method, 'save');

  const res = normalizeOfficeResponse('wps_save_workbook', { success: true });
  assert.equal(res.success, true);
});

test('CAP-50：Microsoft 通道的 set_data_validation(action="read") 判为不安全的参数组合', async () => {
  const { isUnsafeParamError } = await import('../src/bridge/office/normalizer.js');
  // Office.js 的 handleSetDataValidation 与 excel.ps1 的 set_data_validation 都没有 action 分支：
  // 继续下发会把一次"读"变成"写"（COM 侧还会先清掉既有校验）。必须判为 unsafe → 禁止换通道重放。
  let caught: unknown = null;
  try { normalizeOfficeRequest('excel_set_data_validation', { action: 'read', address: 'E5:E20' }); }
  catch (error) { caught = error; }
  assert.ok(caught, 'action=read 必须报错而不是被当成写入下发');
  assert.equal(isUnsafeParamError(caught), true, '必须标记为参数安全护栏（否则 Windows 上会被转手给 COM）');
  assert.match((caught as Error).message, /host=wps|wps_set_data_validation/, '必须给出可用的替代路径');

  // 写入路径不受影响
  const req = normalizeOfficeRequest('excel_set_data_validation', { address: 'E5:E20', validationType: 'list', listItems: ['已通过'] });
  assert.equal(req.method, 'set_data_validation');

  // 空搜索串同样判为不安全（find_and_replace 在 COM 白名单内，不能换通道重放）
  let searchErr: unknown = null;
  try { normalizeOfficeRequest('excel_find_and_replace', { searchQuery: '', replaceText: 'X' }); }
  catch (error) { searchErr = error; }
  assert.equal(isUnsafeParamError(searchErr), true);
});
