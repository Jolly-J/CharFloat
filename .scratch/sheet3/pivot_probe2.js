const sheet = wb.Worksheets.Item(params.sheetName);
const out = [];
try { out.push('PivotTables(1)=' + String(sheet.PivotTables(1))); } catch (e) { out.push('PivotTables(1) ERR ' + e.message); }
try { out.push('new PivotTables()=' + String(new sheet.PivotTables())); } catch (e) { out.push('new ERR ' + e.message); }
try { const pt = sheet.PivotTables(1); out.push('name=' + pt.Name + ' range=' + pt.TableRange1.Address()); } catch (e) { out.push('pt ERR ' + e.message); }
// 强制刷新缓存
try { wb.PivotCaches().Item(1).Refresh(); out.push('cache refresh ok, recordCount=' + wb.PivotCaches().Item(1).RecordCount); } catch (e) { out.push('refresh ERR ' + e.message); }
try { const pt = sheet.PivotTables(1); out.push('after refresh tableRange=' + pt.TableRange1.Address()); } catch (e) {}
return { dump: out.join(' ;; ') };
