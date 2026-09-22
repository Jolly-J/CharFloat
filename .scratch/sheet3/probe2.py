import json, sys
sys.path.insert(0, '.scratch/sheet3')
from call import call
WB="钙钛矿各家企业现状.xlsx"; SH="adv_高级能力实测"
# 复现原始失败：7 列首行含 6 个空字符串，整块 15 行一次写
rows = [["钙钛矿企业能力实测数据（adv_ 自建）","","","","","",""],
        ["区域","企业","技术路线","产能(MW)","效率(%)","良率(%)","项目状态"]]
rows += [["华东","协鑫光电","单结组件",100,19.2,92.5,"已投产"]]*12
r = call('excel_patch_cells', {"workbookName":WB,"sheetName":SH,"address":"A1:G15","values":rows,"reason":"adv_ 复现探针","host":"wps"})
print("FULL:", json.dumps(r, ensure_ascii=False)[:1200])
