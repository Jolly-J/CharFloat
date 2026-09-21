# Opt-in desktop acceptance. Creates its own Excel instance and a disposable workbook.
# Does not attach to or close a user's existing Excel instance.
param([switch]$RunOffice)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$files=Get-ChildItem (Join-Path $root 'resources/office') -Filter '*.ps1'
foreach($file in $files) {
  $tokens=$null;$errors=$null
  [Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$tokens,[ref]$errors) | Out-Null
  if($errors.Count){throw ($errors | Out-String)}
}
Write-Output 'PowerShell syntax: passed'
if(!$RunOffice){Write-Output 'Office acceptance NOT RUN. Use -RunOffice on a Windows desktop with Excel installed.';exit 0}
. (Join-Path $root 'resources/office/common.ps1')
. (Join-Path $root 'resources/office/excel.ps1')
$script:testExcel=New-Object -ComObject Excel.Application
$script:testExcel.Visible=$true
$script:wb=$script:testExcel.Workbooks.Add()
$folder=Join-Path ([IO.Path]::GetTempPath()) ('bridge-acceptance-'+[Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($folder) | Out-Null
$script:wb.SaveAs((Join-Path $folder 'Bridge-中文.xlsx'),51)
function Get-OfficeApp($component){if($component -ne 'excel'){throw 'Test allows Excel only'};return $script:testExcel}
function Invoke($method,$params){$params.workbookName=$script:wb.Name;return Invoke-ExcelTool $method ([pscustomobject]$params)}
function Assert($condition,$message){if(!$condition){throw $message}}
try {
  $sheet=$script:wb.Worksheets.Item(1);$sheet.Name='测试'
  $base=@{sheetName='测试';address='A1:B3';values=@(@('项目','数值'),@('甲',3),@('乙',7))}
  $patch=Invoke 'patch_cells' $base;Assert ($patch.modifiedCount -eq 6) 'Patch did not change six cells'
  $read=Invoke 'read_range' @{sheetName='测试';address='A1:B3';includeFormulas=$true;includeNumberFormats=$true}
  Assert ($read.values[1][0] -eq '甲') 'Unicode value round trip failed'
  Assert ($read.values[2][1] -eq 7) 'Numeric value round trip failed'
  Invoke 'patch_cells' @{sheetName='测试';address='B4';formulas=@(@('=SUM(B2:B3)'))} | Out-Null
  $formula=Invoke 'read_range' @{sheetName='测试';address='B4';includeFormulas=$true}
  Assert ($formula.formulas[0][0] -eq '=SUM(B2:B3)') 'Formula round trip failed'
  Invoke 'format_cells' @{sheetName='测试';address='A1:B1';bold=$true;numberFormat='0.00';backgroundColor='#E9EEFC'} | Out-Null
  $styles=Invoke 'get_range_styles' @{sheetName='测试';address='A1';mode='cells'}
  Assert ($styles.cells[0].bold) 'Bold formatting not applied'
  Invoke 'add_conditional_formatting' @{sheetName='测试';address='B2:B3';ruleType='cell_value';operator='less_than';formula1=5;backgroundColor='#FEE2E2'} | Out-Null
  Assert ($sheet.Range('B2:B3').FormatConditions.Count -eq 1) 'Conditional rule missing'
  $chart=Invoke 'add_chart' @{sheetName='测试';dataRange='A1:B3';chartType='column_clustered';position=[pscustomobject]@{leftCell='D2';width=360;height=220};hasLegend=$false;replaceExisting=$false;title='中文图表'}
  $charts=Invoke 'get_charts' @{sheetName='测试';shapeName=$chart.shapeName;detail=$true}
  Assert ($charts.count -eq 1) 'Chart lookup failed'
  Invoke 'create_sheet' @{sheetName='透视'} | Out-Null
  Invoke 'create_pivot_table' @{sourceSheetName='测试';sourceRange='A1:B3';destSheetName='透视';destCell='A1';rowFields=@('项目');dataFields=@([pscustomobject]@{fieldName='数值';summaryFunction='sum'})} | Out-Null
  Invoke 'set_filter_and_sort' @{sheetName='测试';range='A1:B3';enableAutoFilter=$true;sortRules=@([pscustomobject]@{colIndex=2;order='desc'})} | Out-Null
  Assert ($sheet.Range('B2').Value2 -eq 7) 'Sorting failed'
  Invoke 'set_data_validation' @{sheetName='测试';address='C2:C3';validationType='list';listItems=@('通过','待确认')} | Out-Null
  Assert ($sheet.Range('C2').Validation.Type -eq 3) 'Validation missing'
  $preview=Invoke 'capture_sheet_preview' @{sheetName='测试';address='A1:F10';outputPath=(Join-Path $folder 'preview.png')}
  Assert ($preview.imageBase64.Length -gt 100) 'Preview failed'
  Invoke 'rollback_cells' @{sheetName='测试';address='A1:B3';snapshot=$patch.beforeSnapshot} | Out-Null
  Assert ($null -eq $sheet.Range('A1').Value2) 'Rollback failed'
  $script:wb.Save()
  Write-Output "Office acceptance PASSED. Disposable artifacts: $folder"
} finally {
  $script:wb.Close($false)
  $script:testExcel.Quit()
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:wb) | Out-Null
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:testExcel) | Out-Null
}
