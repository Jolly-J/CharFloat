const sheet = wb.Worksheets.Item(params.sheetName);
const res = [];
// 图标集：良率 F4:F15
try {
  const r1 = sheet.Range('F4:F15');
  const fc1 = r1.FormatConditions.AddIconSetCondition();
  fc1.IconSet = wb.IconSets(2); // 3 个符号（无圆圈）之类，探测用
  res.push('icon: ok type=' + String(fc1.Type) + ' iconID=' + String(fc1.IconSet.ID));
} catch (e) { res.push('icon FAIL: ' + e.message); }
// 公式规则：效率高于均值 -> 绿色字体，E4:E15
try {
  const r2 = sheet.Range('E4:E15');
  const fc2 = r2.FormatConditions.Add(2, 0, '=E4>AVERAGE($E$4:$E$15)');
  fc2.Font.Color = 0x2E7D32;
  res.push('formula: ok type=' + String(fc2.Type) + ' f1=' + String(fc2.Formula1) + ' font=' + String(fc2.Font.Color));
} catch (e) { res.push('formula FAIL: ' + e.message); }
return { sheet: sheet.Name, res: res.join(' ;; '), total: sheet.Cells.FormatConditions.Count };
