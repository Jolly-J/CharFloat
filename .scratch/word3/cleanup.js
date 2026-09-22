const out = {};
const D = doc;
const R = {};
// 删除被误插入目录的测试段落
let removed = 0;
for (let i = D.Paragraphs.Count; i >= 1; i--) {
  if (String(D.Paragraphs.Item(i).Range.Text).indexOf("书签定位写入测试") >= 0) {
    D.Paragraphs.Item(i).Range.Delete();
    removed++;
  }
}
R.removedPollutedPara = removed;
for (let i = 1; i <= D.TablesOfContents.Count; i++) {
  try { D.TablesOfContents.Item(i).Update(); } catch (e) { R["updErr" + i] = String(e.message); }
}
R.toc1Head = (() => { try { return String(D.TablesOfContents.Item(1).Range.Text).slice(0, 80).replace(/[\r\a\t]/g, "|"); } catch (e) { return "ERR"; } })();
out.r = JSON.stringify(R);
return out;
