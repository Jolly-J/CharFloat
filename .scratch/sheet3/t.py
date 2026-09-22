import json, sys, time
sys.path.insert(0, '.scratch/sheet3')
from call import call
WB="钙钛矿各家企业现状.xlsx"; SH="adv_高级能力实测"
def run(tool, args, label=""):
    args = dict(args); args.setdefault('workbookName', WB); args.setdefault('sheetName', SH); args.setdefault('host','wps')
    t0=time.time(); r = call(tool, args); dt=time.time()-t0
    print("### %s [%s] %.2fs" % (label or tool, tool, dt))
    print(json.dumps(r, ensure_ascii=False)[:1800])
    print()
    return r
def rd(addr):
    r = call('excel_read_range', {"workbookName":WB,"sheetName":SH,"address":addr,"host":"wps"})
    d = r.get('data') or {}
    return d.get('values'), d.get('formulas')
def sty(addr, include=None):
    a={"workbookName":WB,"sheetName":SH,"address":addr,"host":"wps","mode":"cells","maxCells":60}
    if include: a['include']=include
    r=call('excel_get_range_styles', a)
    print("### styles %s" % addr); print(json.dumps(r,ensure_ascii=False)[:1500]); print()
    return r
