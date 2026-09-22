// 通用只读反射器：从 params.exprs 读取表达式列表，逐个反射成员
// 成员序列化为紧凑字符串 "name|type|count"（宿主跨界序列化会吃掉嵌套对象）
const exprs = (params && params.exprs) || [];
const opts = (params && params.opts) || {};
const withCount = opts.withCount !== false;
const limit = opts.limit || 0;

function ev(expr) {
  try { return { ok: true, v: eval(expr) }; }
  catch (e) { return { ok: false, e: String(e && e.message ? e.message : e) }; }
}

function reflectValue(v) {
  if (v === null || v === undefined) return { count: -1, members: ['<nullish>'], errs: [] };
  const names = [];
  let forInErr = null;
  try { for (const k in v) names.push(String(k)); } catch (e) { forInErr = String(e); }
  const members = [];
  const errs = [];
  for (const k of names) {
    if (limit && members.length >= limit) break;
    let t, val;
    try { val = v[k]; t = typeof val; }
    catch (e) { errs.push(k + '==' + String(e && e.message ? e.message : e)); continue; }
    let c = '';
    if (withCount && t === 'object' && val !== null) {
      try { const n = val.Count; if (typeof n === 'number') c = String(n); } catch (e) { /* 非集合 */ }
    }
    members.push(k + '|' + t + '|' + c);
  }
  return { count: names.length, members: members, errs: errs, forInErr: forInErr };
}

const out = [];
for (const expr of exprs) {
  const rec = { e: expr };
  const r = ev(expr);
  if (!r.ok) { rec.err = r.e; out.push(rec); continue; }
  rec.t = typeof r.v;
  const R = reflectValue(r.v);
  rec.n = R.count;
  rec.m = R.members.join(',');
  if (R.errs && R.errs.length) rec.re = R.errs.join(' ;; ');
  if (R.forInErr) rec.fe = R.forInErr;
  out.push(rec);
}
return { results: out, count: out.length };
