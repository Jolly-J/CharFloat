#!/usr/bin/env python3
"""用例5：参数边界。走 HTTP /api/v1/tool/call（同一 catalog 校验，无需 MCP 会话）。"""
import json, urllib.request, urllib.error
from pathlib import Path

TOKEN = (Path.home() / '.wps-bridge' / 'token').read_text().strip()
URL = 'http://127.0.0.1:19890/api/v1/tool/call'
WB = "钙钛矿各家企业现状.xlsx"
SH = "fault_case5_params"


def call(name, args):
    body = {'name': name, 'arguments': args, 'clientName': 'fault-param'}
    req = urllib.request.Request(URL, headers={
        'Content-Type': 'application/json', 'Authorization': 'Bearer ' + TOKEN},
        data=json.dumps(body).encode())
    try:
        with urllib.request.urlopen(req, timeout=45) as r:
            return r.status, r.read().decode()
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode()


CASES = [
    ("缺必填 host", 'excel_read_range', {'address': 'A1'}),
    ("host 非法枚举", 'excel_read_range', {'host': 'libreoffice', 'address': 'A1', 'sheetName': SH, 'workbookName': WB}),
    ("address 类型错误(数字)", 'excel_read_range', {'host': 'wps', 'address': 123, 'sheetName': SH, 'workbookName': WB}),
    ("未知参数", 'excel_read_range', {'host': 'wps', 'address': 'A1', 'sheetName': SH, 'workbookName': WB, 'foo': 1}),
    ("maxCells 超上限 99999", 'excel_get_range_styles', {'host': 'wps', 'address': 'A1:B2', 'sheetName': SH, 'workbookName': WB, 'mode': 'cells', 'maxCells': 99999}),
    ("maxCells 负数", 'excel_get_range_styles', {'host': 'wps', 'address': 'A1:B2', 'sheetName': SH, 'workbookName': WB, 'mode': 'cells', 'maxCells': -5}),
    ("freezeRowIndex 负数", 'excel_freeze_panes', {'host': 'wps', 'sheetName': SH, 'workbookName': WB, 'freezeRowIndex': -3}),
    ("freezeRowIndex 字符串'3'", 'excel_freeze_panes', {'host': 'wps', 'sheetName': SH, 'workbookName': WB, 'freezeRowIndex': '3'}),
    ("chartType 非法枚举", 'excel_add_chart', {'host': 'wps', 'sheetName': SH, 'workbookName': WB, 'dataRange': 'A1:B2', 'chartType': 'radar'}),
    ("create_sheet 空名", 'excel_create_sheet', {'host': 'wps', 'workbookName': WB, 'sheetName': ''}),
    ("create_sheet 缺 sheetName", 'excel_create_sheet', {'host': 'wps', 'workbookName': WB}),
    ("patch values 一维数组", 'excel_patch_cells', {'host': 'wps', 'sheetName': SH, 'workbookName': WB, 'address': 'A1:B2', 'values': ['x', 'y']}),
    ("patch values 与 address 尺寸不符", 'excel_patch_cells', {'host': 'wps', 'sheetName': SH, 'workbookName': WB, 'address': 'A1:B2', 'values': [['only-one']]}),
    ("patch 无 values 也无 formulas", 'excel_patch_cells', {'host': 'wps', 'sheetName': SH, 'workbookName': WB, 'address': 'A1:B2'}),
    ("address 超范围 ZZZ9999999", 'excel_read_range', {'host': 'wps', 'address': 'ZZZ9999999', 'sheetName': SH, 'workbookName': WB}),
    ("不存在的 sheet", 'excel_read_range', {'host': 'wps', 'address': 'A1', 'sheetName': 'fault_不存在', 'workbookName': WB}),
    ("不存在的 workbook", 'excel_read_range', {'host': 'wps', 'address': 'A1', 'sheetName': SH, 'workbookName': '不存在.xlsx'}),
    ("audit limit 超上限", 'wps_get_audit_history', {'limit': 100000}),
    ("audit limit 类型错误", 'wps_get_audit_history', {'limit': 'abc'}),
    ("audit view 非法枚举", 'wps_get_audit_history', {'view': 'full'}),
    ("rollback 非法参数名", 'wps_rollback', {'auditID': 'x'}),
    ("未知工具名", 'wps_not_a_tool', {}),
]

if __name__ == '__main__':
    call('excel_create_sheet', {'host': 'wps', 'workbookName': WB, 'sheetName': SH})
    for title, name, args in CASES:
        code, body = call(name, args)
        try:
            d = json.loads(body)
            if d.get('success') is False:
                disp = 'SUCCESS=false :: ' + json.dumps(d.get('error', d), ensure_ascii=False)
            else:
                disp = 'SUCCESS=true  :: ' + json.dumps(d.get('data', d), ensure_ascii=False)
        except Exception:
            disp = body
        print('### %s\n  工具=%s 入参=%s\n  HTTP=%s -> %s\n' %
              (title, name, json.dumps(args, ensure_ascii=False), code, disp[:500]))
