import sys
sys.path.insert(0, '/Users/Python/Office Agent Bridge/.scratch/apimap')
from drv import refl, ev

print('########## 0. 目标确认 ##########')
ev('excel', [
    ('wb.Name', 'wb.Name'),
    ('wb.FullName', 'wb.FullName'),
    ('wb.Sheets.Count', 'wb.Sheets.Count'),
    ('sheetNames', '(function(){var a=[];for(var i=1;i<=wb.Worksheets.Count;i++)a.push(wb.Worksheets.Item(i).Name);return a.join("/");})()'),
])

print('########## 1. Worksheet 成员 ##########')
refl('excel', [
    "wb.Worksheets.Item('Data')",
])
