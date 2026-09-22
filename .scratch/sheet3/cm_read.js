const sheet = wb.Worksheets.Item(params.sheetName);
const addrs = params.addrs;
const lines = [];
for (let i = 0; i < addrs.length; i++) {
  const c = sheet.Range(addrs[i]).Comment;
  if (!c) { lines.push(addrs[i] + '=NONE'); continue; }
  let author = '?', text = '?';
  try { author = c.Author; } catch (e) { author = 'ERR'; }
  try { text = c.Text(); } catch (e) { try { text = c.Text; } catch (e2) { text = 'ERR:' + e2.message; } }
  lines.push(addrs[i] + '=author[' + author + '] text[' + String(text).replace(/\n/g, '\\n') + ']');
}
return { dump: lines.join(' ;; ') };
