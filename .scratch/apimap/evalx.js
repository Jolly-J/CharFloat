// 表达式求值器：返回 JSON 安全的值。params.exprs = [{k: 标签, e: 表达式}]
const items = (params && params.items) || [];
function safe(v, depth) {
  depth = depth || 0;
  if (v === null) return 'null';
  if (v === undefined) return 'undefined';
  const t = typeof v;
  if (t === 'string') return 'str(' + (v.length > 160 ? v.slice(0, 160) + '…' : v) + ')';
  if (t === 'number' || t === 'boolean') return t + '(' + v + ')';
  if (t === 'function') return 'function';
  if (depth > 1) return 'obj';
  try {
    if (v.Name !== undefined && v.Name !== null && typeof v.Name !== 'function') return 'obj{Name=' + v.Name + '}';
  } catch (e) { /* ignore */ }
  try { if (typeof v.Count === 'number') return 'obj{Count=' + v.Count + '}'; } catch (e) { /* ignore */ }
  return 'obj';
}
const out = [];
for (const it of items) {
  let line;
  try { line = it.k + ' = ' + safe(eval(it.e)); }
  catch (e) { line = it.k + ' !! ERR ' + String(e && e.message ? e.message : e); }
  out.push(line);
}
return out;
