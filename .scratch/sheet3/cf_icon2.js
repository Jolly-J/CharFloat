const sheet = wb.Worksheets.Item(params.sheetName);
const out = [];
// 探测 IconSets 的多种入口
try { out.push('wb.IconSets=' + typeof wb.IconSets + '/' + (wb.IconSets ? wb.IconSets.Count : 'null')); } catch (e) { out.push('wb.IconSets ERR ' + e.message); }
try { out.push('app.IconSets=' + typeof app.IconSets); } catch (e) { out.push('app.IconSets ERR ' + e.message); }
try { out.push('wps.IconSets=' + typeof wps.IconSets); } catch (e) { out.push('wps.IconSets ERR ' + e.message); }
try { out.push('wb.Application.IconSets=' + typeof wb.Application.IconSets + '/' + (wb.Application.IconSets ? wb.Application.IconSets.Count : 'null')); } catch (e) { out.push('wb.Application.IconSets ERR ' + e.message); }
// 尝试创建图标集条件格式但不设 IconSet
try {
  const r1 = sheet.Range('F4:F15');
  const fc1 = r1.FormatConditions.AddIconSetCondition();
  out.push('AddIconSetCondition ok, type=' + String(fc1.Type) + ' iconObj=' + String(fc1.IconSet));
  try { fc1.IconSet = 5; out.push('setIconSet(5) ok -> ' + String(fc1.IconSet.ID)); } catch (e) { out.push('setIconSet(5) ERR ' + e.message); }
} catch (e) { out.push('AddIconSetCondition FAIL: ' + e.message); }
// 列出 FormatCondition 可用方法
try { out.push('AddIconSetCondition type=' + typeof sheet.Range('F4').FormatConditions.AddIconSetCondition); } catch (e) {}
return { dump: out.join(' ;; '), total: sheet.Cells.FormatConditions.Count };
