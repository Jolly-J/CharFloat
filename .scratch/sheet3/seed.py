import json, sys
sys.path.insert(0, '.scratch/sheet3')
from call import call

WB = "钙钛矿各家企业现状.xlsx"
SH = "adv_高级能力实测"
HEAD = ["区域","企业","技术路线","产能(MW)","效率(%)","良率(%)","项目状态"]
DATA = [
 ["华东","协鑫光电","单结组件",100,19.2,92.5,"已投产"],
 ["华东","极电光能","单结组件",150,18.8,88.0,"已投产"],
 ["华北","纤纳光电","单结组件",120,19.6,95.1,"已投产"],
 ["华南","脉络能源","叠层",80,22.4,86.3,"中试"],
 ["西南","通威股份","叠层",300,21.1,90.7,"建设中"],
 ["华东","隆基绿能","叠层",250,26.5,83.2,"研发"],
 ["华北","晶科能源","叠层",200,25.8,85.4,"中试"],
 ["华南","华晟新能源","单结组件",90,17.9,91.2,"已投产"],
 ["西南","宁德时代","叠层",180,23.7,89.6,"建设中"],
 ["华东","杭萧绿建","单结组件",60,16.5,78.4,"研发"],
 ["华北","京东方能源","叠层",140,20.3,87.1,"中试"],
 ["西南","天合光能","叠层",220,24.6,93.8,"已投产"],
]
rows = [["钙钛矿企业能力实测数据（adv_ 自建）","","","","","",""], HEAD] + DATA
last = 1 + len(rows)
r = call('excel_patch_cells', {"workbookName": WB, "sheetName": SH, "address": "A1:G%d" % last,
    "values": rows, "reason": "adv_ 高级能力实测：写入演示数据", "host": "wps"})
print(json.dumps({"last": last, "write": r.get('success'), "data": r.get('data')}, ensure_ascii=False)[:900])
# read back
rb = call('excel_read_range', {"workbookName": WB, "sheetName": SH, "address": "A1:G%d" % last, "host": "wps"})
print(json.dumps(rb.get('data'), ensure_ascii=False)[:1200])
