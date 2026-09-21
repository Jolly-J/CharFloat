$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'resources/office/common.ps1')
. (Join-Path $root 'resources/office/excel.ps1')
function Assert($value,$message){if(!$value){throw $message}}
class FakeCollection {
  [object[]]$Items
  [int]$Count
  FakeCollection([object[]]$items){$this.Items=$items;$this.Count=$items.Length}
  [object] Item([int]$i){return $this.Items[$i-1]}
}
$doc=[pscustomobject]@{Name='Exact.xlsx';FullName='C:\Test\Exact.xlsx'}
$app=[pscustomobject]@{Workbooks=[FakeCollection]::new(@($doc))}
Assert ((Get-Target $app 'excel' $null).Name -eq 'Exact.xlsx') 'Single workbook selection failed'
Assert ((Get-Target $app 'excel' 'C:\Test\Exact.xlsx').Name -eq 'Exact.xlsx') 'Full path selection failed'
$rejected=$false;try{Get-Target $app 'excel' 'Exact'}catch{$rejected=$true};Assert $rejected 'Partial names must be rejected'
$other=[pscustomobject]@{Name='Exact.xlsx';FullName='D:\Test\Exact.xlsx'}
$app.Workbooks=[FakeCollection]::new(@($doc,$other))
$rejected=$false;try{Get-Target $app 'excel' 'Exact.xlsx'}catch{$rejected=$true};Assert $rejected 'Duplicate names must be rejected'
Assert ((Get-Target $app 'excel' 'D:\Test\Exact.xlsx').FullName -eq $other.FullName) 'Exact path resolution failed'
$raw=[Array]::CreateInstance([object],[int[]]@(2,2),[int[]]@(1,1))
$raw.SetValue('text',1,1);$raw.SetValue(0,1,2);$raw.SetValue($null,2,1);$raw.SetValue(0.72,2,2)
$range=[pscustomobject]@{Rows=[pscustomobject]@{Count=2};Columns=[pscustomobject]@{Count=2};Value2=$raw;Formula=$raw}
$matrix=Get-Matrix $range 'Value2'
Assert ($matrix.Count -eq 2 -and $matrix[0].Count -eq 2) '2D matrix shape lost'
Assert ($matrix[0][1] -eq 0 -and $matrix[1][1] -eq 0.72) 'Numeric matrix values changed'
$one=[pscustomobject]@{Rows=[pscustomobject]@{Count=1};Columns=[pscustomobject]@{Count=1};Value2=0;Formula=0}
$single=Get-Matrix $one 'Value2';Assert ($single.Count -eq 1 -and $single[0].Count -eq 1 -and $single[0][0] -eq 0) 'Scalar zero normalization failed'
Set-Matrix $range 'Value2' @(@(1,2),@(3,4));Assert ($range.Value2.Rank -eq 2 -and $range.Value2[1,1] -eq 4) 'COM matrix write shape failed'
$before=$range.Value2;$rejected=$false;try{Set-Matrix $range 'Value2' @(@(1),@(3,4))}catch{$rejected=$true};Assert $rejected 'Invalid shape accepted';Assert ([object]::ReferenceEquals($before,$range.Value2)) 'Invalid shape mutated range'
Assert ((Format-Color (Get-Color '#12ABEF')) -eq '#12ABEF') 'COM color order failed'
Write-Output 'Native data contracts: passed (mock objects; not Office COM acceptance)'
