const out = {};
const D = doc;
const T = D.Tables.Item(2);
const R = {};
R.cell_2_2_before = { shading: T.Cell(2,2).Shading.BackgroundPatternColor, font: T.Cell(2,2).Range.Font.Name, size: T.Cell(2,2).Range.Font.Size, bold: T.Cell(2,2).Range.Font.Bold };
R.uniform = T.Uniform;
R.row5cells = T.Rows.Item(5).Cells.Count;
out.r = JSON.stringify(R);
return out;
