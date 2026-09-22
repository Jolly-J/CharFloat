const out = {};
const D = doc;
const R = {};
R.echo = params.s;
const p = D.Paragraphs.Add();
p.Range.Text = params.s + "\n";
R.wrote = String(p.Range.Text).replace(/[\r\a\n]/g, "");
out.r = JSON.stringify(R);
return out;
