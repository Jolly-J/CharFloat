const sheet = wb.Worksheets.Item(params.sheetName);
const out = [];
const rng = sheet.Range('N17:O19');
function reset() { sheet.Range('N17:O19').Value2 = [[3,'c'],[1,'a'],[2,'b']]; }
// sheet.Sort 对象
try { out.push('sheet.Sort typeof=' + typeof sheet.Sort); } catch (e) { out.push('sheet.Sort ERR ' + e.message); }
reset();
try {
  const s = sheet.Sort;
  s.SortFields.Clear();
  s.SortFields.Add(rng.Columns.Item(1), 0, 1);
  s.SetRange(rng);
  s.Header = 2;
  s.Apply();
  out.push('sheet.Sort -> ' + String(sheet.Range('N17:O19').Value2));
} catch (e) { out.push('sheet.Sort FAIL ' + e.message); }
// wb.Sort?
try { out.push('wb.Sort typeof=' + typeof wb.Sort); } catch (e) {}
// Range.AutoFilter 定位
reset();
try {
  const af = rng.AutoFilter(1, undefined, 1); // 字段1, 升序
  out.push('AutoFilter(1,_,1) -> ' + String(sheet.Range('N17:O19').Value2));
} catch (e) { out.push('AutoFilter FAIL ' + e.message); }
try { out.push('AutoFilterMode=' + sheet.AutoFilterMode + ' range=' + sheet.AutoFilter.Range.Address()); } catch (e) {}
if (sheet.AutoFilterMode) { try { sheet.AutoFilterMode = false; } catch(e){} }
return { dump: out.join(' ;; ') };
