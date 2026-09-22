const sheet = wb.Worksheets.Item(params.sheetName);
const out = [];
try { out.push('typeof sheet.PivotTables=' + typeof sheet.PivotTables); } catch (e) { out.push('PT ERR ' + e.message); }
try { out.push('pivotTables=' + String(sheet.PivotTables)); } catch (e) {}
try { out.push('caches=' + wb.PivotCaches().Count); } catch (e) { out.push('caches ERR ' + e.message); }
try {
  const pc = wb.PivotCaches().Item(1);
  out.push('cache1 recordCount=' + pc.RecordCount + ' srcData=' + String(pc.SourceData));
} catch (e) { out.push('cache1 ERR ' + e.message); }
// 枚举 sheet 上所有可能的透视对象入口
const keys = [];
for (const k in sheet) { keys.push(k); }
out.push('keys=' + keys.join(','));
return { dump: out.join(' ;; ') };
