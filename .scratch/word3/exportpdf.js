const out = {};
const R = {};
const D = doc;
try {
  const p = "/Users/Python/Office Agent Bridge/.scratch/word3/word3-final.pdf";
  D.ExportAsFixedFormat(p, 17);
  R.exported = p;
} catch (e) { R.exportErr = String(e.message); }
try { R.pages = D.ComputeStatistics(2); } catch (e) { R.pagesErr = String(e.message); }
out.r = JSON.stringify(R);
return out;
