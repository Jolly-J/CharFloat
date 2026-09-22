const out = {};
const D = doc;
const R = {};
R.echo = params.s;
const p = D.Paragraphs.Add();
p.Range.InsertAfter(params.s + "\r");
R.insertAfter = String(p.Range.Text);
// 第二种写法：Content 末尾插入
const endPos = D.Content.End - 1;
D.Range(endPos, endPos).InsertAfter(params.s + "\r");
R.tailText = String(D.Paragraphs.Item(D.Paragraphs.Count).Range.Text);
// 第三种：写 Text 后再读同一对象
const p3 = D.Paragraphs.Add();
p3.Range.Text = params.s;
R.objText = String(p3.Range.Text);
R.objTextRaw = JSON.stringify(String(p3.Range.Text));
out.r = JSON.stringify(R);
return out;
