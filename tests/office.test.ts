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
