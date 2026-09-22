await Excel.run(async (context) => {
  const s = context.workbook.worksheets.getActiveWorksheet();
  s.getRange("A1").values = [["MINI-" + Date.now()]];
  await context.sync();
});
return { ok: true, mini: true, at: new Date().toISOString() };
