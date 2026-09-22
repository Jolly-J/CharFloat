const sheet = wb.Worksheets.Item(params.sheetName);
const out = [];
// 清掉 F4:F15 上的全部图标集规则（探测期产生的），重建一条干净的
const rng = sheet.Range('F4:F15');
let removed = 0;
for (let i = rng.FormatConditions.Count; i >= 1; i--) {
  const fc = rng.FormatConditions.Item(i);
  let t = -1; try { t = fc.Type; } catch (e) {}
  if (t === 6) { fc.Delete(); removed++; }
}
out.push('removedIconRules=' + removed);
const fc1 = rng.FormatConditions.AddIconSetCondition();
fc1.IconSet = wb.IconSets.Item(5);
out.push('newIconRule iconID=' + String(fc1.IconSet.ID) + ' type=' + String(fc1.Type));
return { dump: out.join(' ;; '), F4Count: rng.FormatConditions.Count };
