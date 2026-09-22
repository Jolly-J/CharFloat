const sheet = wb.Worksheets.Item(params.sheetName);
const out = [];
function reset() { sheet.Range('N17:O19').Value2 = [[3,'c'],[1,'a'],[2,'b']]; }
function snap(tag, r) { out.push(tag + ' -> ' + String(sheet.Range('N17:O19').Value2)); }
const rng = sheet.Range('N17:O19');

reset();
try { rng.Sort(sheet.Range('N17'), 1); snap('A Sort(N17Range,1)', rng); } catch(e){ out.push('A FAIL '+e.message); }

reset();
try { rng.Sort(sheet.Range('N17'), 1, undefined, undefined, undefined, undefined, 0); snap('B Sort(N17Range,1,...,0)', rng); } catch(e){ out.push('B FAIL '+e.message); }

reset();
try { rng.Sort(rng.Columns.Item(1), 1); snap('C Sort(colItem,1)', rng); } catch(e){ out.push('C FAIL '+e.message); }

reset();
try { rng.Sort(rng.Columns.Item(1), 0); snap('D Sort(colItem,0)', rng); } catch(e){ out.push('D FAIL '+e.message); }

reset();
try { rng.Sort(rng.Columns(1), 1); snap('E Sort(Columns(1),1)', rng); } catch(e){ out.push('E FAIL '+e.message); }

reset();
try { const s = rng.Sort; out.push('sortKeys=' + Object.keys(s).join(',')); } catch(e){ out.push('keys FAIL '+e.message); }
reset();
try { const s = rng.Sort; s.SortFields.Add(rng.Columns.Item(1), 0, 1); s.SetRange(rng); s.Header = 2; s.Apply(); snap('F SortFields', rng); } catch(e){ out.push('F FAIL '+e.message); }
reset();
try { rng.Sort(rng.Columns.Item(1), 1, rng.Columns.Item(2), 1, undefined, undefined, 0); snap('G 2key', rng); } catch(e){ out.push('G FAIL '+e.message); }
return { dump: out.join(' ;; ') };
