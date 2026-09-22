#!/usr/bin/env python3
"""按组件逐表达式反射，单表达式中断可恢复。落盘 out/<comp>.json（增量写）。"""
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from drv import reflect, call, XL, WD, PP  # noqa: E402

HERE = Path(__file__).parent
OUT = HERE / 'out'
OUT.mkdir(exist_ok=True)

EXCEL = [
    'wb',
    'wb.Worksheets.Item("Data")',
    'wb.Worksheets.Item("Probe")',
    'wb.Worksheets.Item("Data").Range("A1:B2")',
    'wb.Worksheets.Item("Data").Rows',
    'wb.Worksheets.Item("Data").Columns',
    'wb.Worksheets.Item("Data").ListObjects.Item(1)',
    'wb.Worksheets.Item("Data").ListObjects',
    'wb.Worksheets.Item("Probe").PivotTables()',
    'wb.PivotCaches()',
    'wb.Worksheets.Item("Data").Range("D2:D11").FormatConditions.Item(1)',
    'wb.Worksheets.Item("Data").Range("D2:D11").FormatConditions',
    'wb.Names',
    'wb.Names.Item(1)',
    'wb.Worksheets.Item("Data").Range("B2:B11").Validation',
    'wb.Worksheets.Item("Data").Range("A1").Comment',
    'wb.Worksheets.Item("Data").Range("A2").CommentThreaded',
    'wb.Worksheets.Item("Data").Comments',
    'wb.Worksheets.Item("Data").CommentsThreaded',
    'wb.Worksheets.Item("Data").Sort',
    'wb.Worksheets.Item("Data").Sort.SortFields',
    'wb.Worksheets.Item("Data").AutoFilter',
    'wb.Worksheets.Item("Data").AutoFilter.Filters',
    'wb.Worksheets.Item("Data").PageSetup',
    'wb.Worksheets.Item("Probe").Shapes.Item(1)',
    'wb.Worksheets.Item("Probe").Shapes',
    'wb.Worksheets.Item("Probe").ChartObjects().Item(1)',
    'wb.Worksheets.Item("Probe").ChartObjects().Item(1).Chart',
    'wb.Worksheets.Item("Data").QueryTables',
    'wb.Connections',
    'app.Windows.Item(1)',
    'app.Windows.Item(1).Panes',
    'wb.Worksheets.Item("Data").Tab',
    'wb.Worksheets.Item("Data").Outline',
    'wb.Worksheets.Item("Data").Hyperlinks',
    'wb.Worksheets.Item("Data").Protection',
    'wb.Worksheets.Item("Data").CustomProperties',
    'wb.Worksheets.Item("Data").NamedSheetViews',
    'wb.Worksheets.Item("Data").Names',
    'wb.Styles',
    'wb.Worksheets.Item("Data").UsedRange',
    'wb.Worksheets.Item("Data").Range("A1").Font',
    'wb.Worksheets.Item("Data").Range("A1").Interior',
    'wb.Worksheets.Item("Data").Range("A1").Borders',
    'wb.Worksheets.Item("Data").Range("A1").NumberFormat',
    'app.ActiveWindow',
    'app.CommandBars',
    'wb.Worksheets.Item("Data").Cells.Item(1,1)',
    'wb.Worksheets.Item("Data").Sheets',
    'wb.Worksheets.Item("Data").PivotTables()',
    'wb.Worksheets.Item("Data").PivotTableWizard',
    'wb.Worksheets.Item("Data").Scenarios',
    'wb.Worksheets.Item("Data").MailEnvelope',
    'wb.Worksheets.Item("Data").SmartTags',
    'wb.Worksheets.Item("Data").Scripts',
    'wb.Worksheets.Item("Data").XmlMapQuery',
    'wb.CustomDocumentProperties',
    'wb.Container',
    'app.ActiveWindow.SelectedSheets',
    'wb.Worksheets.Item("Probe").ChartObjects().Item(1).Chart.SeriesCollection()',
    'wb.Worksheets.Item("Probe").Shapes.Item(1).TextFrame',
    'wb.Worksheets.Item("Probe").Shapes.Item(1).Fill',
    'wb.Worksheets.Item("Probe").Shapes.Item(1).Line',
    'wb.Worksheets.Item("Data").ListObjects.Item(1).Range',
    'wb.Worksheets.Item("Data").ListObjects.Item(1).Sort',
    'wb.Worksheets.Item("Data").ListObjects.Item(1).AutoFilter',
    'wb.Worksheets.Item("Data").ListObjects.Item(1).ListColumns',
    'wb.Worksheets.Item("Data").Names.Item(1)',
]


def healthy(comp):
    r = call('bridge_get_capabilities', {})
    c = ((r.get('data') or {}).get('connection') or {})
    if not c.get('isWpsConnected'):
        return False
    if comp in ('excel', 'word', 'ppt'):
        return bool(((c.get('components') or {}).get(comp) or {}).get('connected'))
    return True


def wait_health(comp, tries=30, delay=3):
    for _ in range(tries):
        if healthy(comp):
            return True
        time.sleep(delay)
    return False


def run(comp, exprs):
    f = OUT / ('%s.json' % comp)
    data = json.loads(f.read_text()) if f.exists() else {}
    for i, e in enumerate(exprs):
        if e in data and not data[e].get('drop') and not str(data[e].get('err','')).startswith('CALL:'):
            print('[--] skip %s' % e[:56]); continue
        r = reflect(comp, [e])
        rec = {}
        if not r.get('success'):
            rec = {'err': 'CALL: ' + str(r.get('error'))[:200], 'drop': True}
            print('[%2d/%d] %-56s CALL-FAIL %s' % (i + 1, len(exprs), e[:56], rec['err'][:80]))
            data[e] = rec
            f.write_text(json.dumps(data, ensure_ascii=False, indent=1))
            wait_health(comp)
            continue
        d = r.get('data') or {}
        if not d.get('success'):
            rec = {'err': 'SCRIPT: ' + json.dumps(d, ensure_ascii=False)[:200]}
        else:
            res = ((d.get('returnValue') or {}).get('results') or [{}])[0]
            rec = res
        data[e] = rec
        status = rec.get('err') or ('type=%s n=%s' % (rec.get('t'), rec.get('n')))
        print('[%2d/%d] %-56s %s' % (i + 1, len(exprs), e[:56], str(status)[:90]))
        f.write_text(json.dumps(data, ensure_ascii=False, indent=1))
        time.sleep(0.15)
    return data


if __name__ == '__main__':
    comp = sys.argv[1]
    exprs = EXCEL if comp == 'excel' else json.loads(Path(sys.argv[2]).read_text())
    run(comp, exprs)
