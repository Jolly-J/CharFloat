import sys, json
sys.path.insert(0,'.scratch/sheet3')
from call import call
WB="钙钛矿各家企业现状.xlsx"
def js(code, params=None, comp="excel", wbname=WB):
    r = call('wps_execute_script', {"workbookName":wbname,"component":comp,"params":params or {},"code":code})
    print(json.dumps(r, ensure_ascii=False)[:6000])
    return r
