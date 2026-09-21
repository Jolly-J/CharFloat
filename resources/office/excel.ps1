# Structured Excel COM adapter. All entry data arrives as JSON over stdin.
function Get-Matrix($range, $property) {
  $raw=$range.$property; $rows=New-Object System.Collections.Generic.List[object]
  for($r=1;$r -le $range.Rows.Count;$r++) {
    $row=New-Object System.Collections.Generic.List[object]
    for($c=1;$c -le $range.Columns.Count;$c++) {
      if($raw -is [Array] -and $raw.Rank -eq 2) { $row.Add($raw.GetValue($r,$c)) }
      else { $row.Add($raw) }
    }
    $rows.Add($row.ToArray())
  }
  return ,$rows.ToArray()
}
function Set-Matrix($range,$property,$rows) {
  if($rows.Count -ne $range.Rows.Count) { throw 'Row count must match the target range' }
  $matrix=New-Object 'object[,]' $range.Rows.Count,$range.Columns.Count
  for($r=0;$r -lt $rows.Count;$r++) {
    if($rows[$r].Count -ne $range.Columns.Count) { throw 'Column count must match the target range' }
    for($c=0;$c -lt $rows[$r].Count;$c++) { $matrix[$r,$c]=$rows[$r][$c] }
  }
  $range.$property=$matrix
}
function Get-Snapshot($sheet,$range) {
  return @{workbookName=$sheet.Parent.Name;sheetName=$sheet.Name;address=$range.Address();rowCount=$range.Rows.Count;columnCount=$range.Columns.Count;values=(Get-Matrix $range 'Value2');formulas=(Get-Matrix $range 'Formula')}
}
function Get-Color($hex) {
  if($hex -notmatch '^#[0-9a-fA-F]{6}$') { throw 'Color must be #RRGGBB' }
  return [Convert]::ToInt32($hex.Substring(1,2),16) + 256*[Convert]::ToInt32($hex.Substring(3,2),16) + 65536*[Convert]::ToInt32($hex.Substring(5,2),16)
}
function Format-Color($color) {
  if($null -eq $color) { return $null }
  $n=[int]$color
  return '#{0:X2}{1:X2}{2:X2}' -f ($n -band 255),(($n -shr 8) -band 255),(($n -shr 16) -band 255)
}
function Get-Sheet($wb,$name) { if($name) { return $wb.Worksheets.Item($name) }; return $wb.ActiveSheet }
function Check-Range($range,$max=10000) { if($range.Areas.Count -ne 1 -or $range.Cells.CountLarge -gt $max) { throw "Use one rectangular range with at most $max cells" } }
function Invoke-ExcelTool($method,$p) {
  $app=Get-OfficeApp 'excel'
  if($method -eq 'get_workspace_summary') {
    if($app.Workbooks.Count -eq 0) { return @{hasOpenWorkbook=$false;openWorkbooks=@()} }
    $wb=Get-Target $app 'excel' $p.workbookName
    return @{hasOpenWorkbook=$true;workbookName=$wb.Name;fullName=$wb.FullName;activeSheetName=$wb.ActiveSheet.Name;sheetCount=$wb.Worksheets.Count;sheets=@($wb.Worksheets | ForEach-Object {@{index=$_.Index;name=$_.Name;visible=($_.Visible -eq -1)}});openWorkbooks=@($app.Workbooks | ForEach-Object {@{name=$_.Name;fullName=$_.FullName}})}
  }
  $wb=Get-Target $app 'excel' $p.workbookName
  if($method -eq 'create_sheet') {
    if(@($wb.Worksheets | Where-Object {$_.Name -ceq $p.sheetName}).Count) { throw 'Sheet already exists' }
    $sheet=$wb.Worksheets.Add(); $sheet.Name=$p.sheetName
    return @{success=$true;workbookName=$wb.Name;sheetName=$sheet.Name}
  }
  $sheet=Get-Sheet $wb $p.sheetName
  $range=$null
  if($p.address) { $range=$sheet.Range($p.address); Check-Range $range }
  switch($method) {
    'get_sheet_outline' {
      $used=$sheet.UsedRange
      $sample=$used.Resize([Math]::Min(3,$used.Rows.Count),[Math]::Min(50,$used.Columns.Count))
      return @{sheetName=$sheet.Name;isEmpty=($app.WorksheetFunction.CountA($used) -eq 0);usedRangeAddress=$used.Address();startRow=$used.Row;startColumn=$used.Column;rowCount=$used.Rows.Count;columnCount=$used.Columns.Count;headerPreview=(Get-Matrix $sample 'Value2')}
    }
    'read_range' {
      $result=@{sheetName=$sheet.Name;address=$range.Address();rowCount=$range.Rows.Count;columnCount=$range.Columns.Count;values=(Get-Matrix $range 'Value2');formulas=$null;numberFormats=$null}
      if($p.includeFormulas) {$result.formulas=Get-Matrix $range 'Formula'}
      if($p.includeNumberFormats) {
        $formats=New-Object System.Collections.Generic.List[object]
        for($r=1;$r -le $range.Rows.Count;$r++){ $row=@();for($c=1;$c -le $range.Columns.Count;$c++){$row+=,$range.Cells.Item($r,$c).NumberFormat};$formats.Add($row) }
        $result.numberFormats=$formats.ToArray()
      }
      return $result
    }
    'get_range_styles' {
      $include=$p.include;if(!$include){$include=@('fontName','fontSize','bold','fontColor','backgroundColor','numberFormat','horizontalAlignment','verticalAlignment','wrapText','rowHeight','columnWidth','merged','mergeArea')}
      $base=@{workbookName=$wb.Name;sheetName=$sheet.Name;address=$range.Address();rowCount=$range.Rows.Count;columnCount=$range.Columns.Count;mode=$p.mode}
      if($p.mode -ne 'cells') {
        $styles=Read-ExcelStyle $range $include
        $base.styles=$styles;$base.mixedOrUnavailableFields=@($styles.Keys | Where-Object {$null -eq $styles[$_] -and !($_ -eq 'mergeArea' -and $styles.merged -eq $false)})
      } else {
        $max=100;if($p.maxCells){$max=[Math]::Max(1,[Math]::Min(500,[int]$p.maxCells))};$cells=@()
        foreach($cell in $range.Cells){$style=Read-ExcelStyle $cell $include;$style.address=$cell.Address();$cells+=,$style;if($cells.Count -ge $max){break}}
        $base.cells=$cells;$base.totalCells=$range.Cells.CountLarge;$base.returnedCells=$cells.Count;$base.truncated=$cells.Count -lt $range.Cells.CountLarge
      }
      return $base
    }
    'patch_cells' {
      $before=Get-Snapshot $sheet $range
      # Validate both matrices before the first mutation.
      foreach($prop in @('values','formulas')) {if($null -ne $p.$prop){if($p.$prop.Count -ne $range.Rows.Count){throw 'Row count mismatch'};foreach($row in $p.$prop){if($row.Count -ne $range.Columns.Count){throw 'Column count mismatch'}}}}
      try { if($null -ne $p.values){Set-Matrix $range 'Value2' $p.values};if($null -ne $p.formulas){Set-Matrix $range 'Formula' $p.formulas} }
      catch { $original=$_.Exception.Message;try{Set-Matrix $range 'Formula' $before.formulas}catch{throw "Partial write; restoration failed. $original"};throw "Write failed; original formulas restored. $original" }
      $after=Get-Snapshot $sheet $range;$diff=@()
      for($r=0;$r -lt $range.Rows.Count;$r++){for($c=0;$c -lt $range.Columns.Count;$c++){
        if($before.values[$r][$c] -cne $after.values[$r][$c] -or $before.formulas[$r][$c] -cne $after.formulas[$r][$c]){$diff+=@{cell=$range.Cells.Item($r+1,$c+1).Address();oldValue=$before.values[$r][$c];newValue=$after.values[$r][$c];oldFormula=$before.formulas[$r][$c];newFormula=$after.formulas[$r][$c]}}
      }}
      return @{workbookName=$wb.Name;sheetName=$sheet.Name;address=$range.Address();modifiedCount=$diff.Count;diff=$diff;beforeSnapshot=$before;afterSnapshot=$after}
    }
    'rollback_cells' { Set-Matrix $range 'Formula' $p.snapshot.formulas;return @{success=$true;restoredRows=$range.Rows.Count;restoredCols=$range.Columns.Count} }
    'clear_range' { $range.Clear();return @{success=$true} }
    'delete_sheet' {
      if($wb.Worksheets.Count -le 1){throw 'Cannot delete the last worksheet'}
      $alerts=$app.DisplayAlerts;try{$app.DisplayAlerts=$false;$sheet.Delete()}finally{$app.DisplayAlerts=$alerts};return @{success=$true}
    }
    'duplicate_sheet' {
      if(@($wb.Worksheets | Where-Object {$_.Name -ceq $p.newSheetName}).Count){throw 'Destination sheet already exists'}
      $src=Get-Sheet $wb $p.sourceSheetName
      if($p.position -eq 'before'){$src.Copy($src)}elseif($p.position -eq 'end'){$src.Copy([Type]::Missing,$wb.Worksheets.Item($wb.Worksheets.Count))}else{$src.Copy([Type]::Missing,$src)}
      $wb.ActiveSheet.Name=$p.newSheetName;return @{success=$true;sheetName=$p.newSheetName}
    }
    'format_cells' {
      if($p.merge -and $app.WorksheetFunction.CountA($range) -gt 1){throw 'Merge would discard values; clear or relocate content first'}
      if($p.fontName){$range.Font.Name=$p.fontName};if($p.fontSize){$range.Font.Size=$p.fontSize};if($null -ne $p.bold){$range.Font.Bold=$p.bold}
      if($p.fontColor){$range.Font.Color=Get-Color $p.fontColor};if($p.backgroundColor){$range.Interior.Color=Get-Color $p.backgroundColor}
      if($p.horizontalAlignment){$range.HorizontalAlignment=@{left=-4131;center=-4108;right=-4152}[$p.horizontalAlignment]}
      if($p.verticalAlignment){$range.VerticalAlignment=@{top=-4160;center=-4108;bottom=-4107}[$p.verticalAlignment]}
      if($p.numberFormat){$range.NumberFormat=$p.numberFormat};if($p.rowHeight){$range.RowHeight=$p.rowHeight};if($null -ne $p.wrapText){$range.WrapText=$p.wrapText}
      if($p.borders){$color='#CBD5E1';if($p.borders -is [string]){$color=$p.borders};foreach($i in @(7,8,9,10,11,12)){$b=$range.Borders.Item($i);$b.LineStyle=1;$b.Weight=2;$b.Color=Get-Color $color}}
      if($p.merge){$range.Merge()};if($p.unmerge){$range.UnMerge()}
      return @{success=$true;address=$range.Address()}
    }
    'auto_fit_columns' {
      if($p.columnRules){foreach($rule in $p.columnRules){$col=$sheet.Columns.Item($rule.colIndex);$col.AutoFit();if($rule.minWidth -and $col.ColumnWidth -lt $rule.minWidth){$col.ColumnWidth=$rule.minWidth};if($rule.maxWidth -and $col.ColumnWidth -gt $rule.maxWidth){$col.ColumnWidth=$rule.maxWidth;$col.WrapText=($rule.wrapText -ne $false)}}}else{$sheet.UsedRange.Columns.AutoFit()};return @{success=$true}
    }
    'freeze_panes' {
      $wb.Activate();$sheet.Activate();$window=$app.ActiveWindow;$window.FreezePanes=$false
      if(!$p.unfreeze){$window.SplitRow=[int]$p.freezeRowIndex;$window.SplitColumn=[int]$p.freezeColumnIndex;$window.FreezePanes=$true};return @{success=$true}
    }
    {$_ -in @('modify_rows_columns','manage_rows_and_columns','insert_dimension')} {
      $kind=$p.targetType;$action=$p.action;if($method -eq 'insert_dimension'){$kind=$p.type;$action='insert'}
      $count=1;if($p.count){$count=[int]$p.count};if($count -lt 1 -or $count -gt 1000){throw 'count must be 1..1000'}
      if($kind -eq 'row'){$index=[int]$p.index;$target=$sheet.Rows.Item($index).Resize($count)}else{$index=$p.index;if($index -match '^\d+$'){$index=[int]$index};$target=$sheet.Columns.Item($index).Resize([Type]::Missing,$count)}
      switch($action){'insert'{$target.Insert() | Out-Null}'delete'{$target.Delete() | Out-Null}'hide'{$target.Hidden=$true}'unhide'{$target.Hidden=$false}'set_size'{if($kind -eq 'row'){$target.RowHeight=$p.size}else{$target.ColumnWidth=$p.size}}default{throw 'Unknown dimension action'}}
      return @{success=$true}
    }
    'manage_sheet' {
      switch($p.action){'rename'{$sheet.Name=$p.newName}'tab_color'{$sheet.Tab.Color=Get-Color $p.color}'protect'{$sheet.Protect($p.password)}'unprotect'{$sheet.Unprotect($p.password)}'move'{if($p.targetIndex -lt 1 -or $p.targetIndex -gt $wb.Worksheets.Count){throw 'Sheet index out of range'};$to=$wb.Worksheets.Item($p.targetIndex);if($sheet.Index -lt $p.targetIndex){$sheet.Move([Type]::Missing,$to)}else{$sheet.Move($to)}}default{throw 'Unknown sheet action'}};return @{success=$true;sheetName=$sheet.Name}
    }
    {$_ -in @('search_cells','find_and_replace')} {
      $query=[string]$p.query;if($method -eq 'find_and_replace'){$query=[string]$p.searchQuery}
      $scope=$sheet.UsedRange;if($p.searchRange){$scope=$sheet.Range($p.searchRange)};Check-Range $scope 100000
      $limit=50;if($p.maxResults){$limit=[Math]::Min(500,[int]$p.maxResults)};$matches=@();$replaced=0
      $values=Get-Matrix $scope 'Value2'
      for($r=0;$r -lt $scope.Rows.Count;$r++){for($c=0;$c -lt $scope.Columns.Count;$c++){
        $text=[string]$values[$r][$c];$comparison=[StringComparison]::OrdinalIgnoreCase;if($p.matchCase){$comparison=[StringComparison]::Ordinal}
        $found=$text.IndexOf($query,$comparison) -ge 0;if($p.matchEntireCell){$found=[string]::Equals($text,$query,$comparison)}
        if($found){$cell=$scope.Cells.Item($r+1,$c+1);$matches+=@{address=$cell.Address();row=$cell.Row;column=$cell.Column;value=$values[$r][$c]}
          if($null -ne $p.replaceText){$options=[Text.RegularExpressions.RegexOptions]::IgnoreCase;if($p.matchCase){$options=[Text.RegularExpressions.RegexOptions]::None};$replacement=[string]$p.replaceText;if($p.matchEntireCell){$cell.Value2=$replacement}else{$regex=New-Object Text.RegularExpressions.Regex([regex]::Escape($query),$options);$cell.Value2=$regex.Replace($text,[Text.RegularExpressions.MatchEvaluator]{param($m) $replacement})};$replaced++}
          if($matches.Count -ge $limit){break}
        }
      };if($matches.Count -ge $limit){break}}
      return @{sheetName=$sheet.Name;query=$query;totalFound=$matches.Count;matches=$matches;replacedCount=$replaced;truncated=($matches.Count -ge $limit)}
    }
    'set_filter_and_sort' {
      $target=$sheet.Range($p.range)
      if($null -ne $p.enableAutoFilter){if($p.enableAutoFilter -and !$sheet.AutoFilterMode){$target.AutoFilter() | Out-Null};if(!$p.enableAutoFilter -and $sheet.AutoFilterMode){$sheet.AutoFilterMode=$false}}
      if($p.sortRules){$sort=$sheet.Sort;$sort.SortFields.Clear();foreach($rule in $p.sortRules){$order=1;if($rule.order -eq 'desc'){$order=2};$sort.SortFields.Add($target.Columns.Item([int]$rule.colIndex),0,$order) | Out-Null};$sort.SetRange($target);$sort.Header=1;$sort.Apply()};return @{success=$true}
    }
    'set_data_validation' {
      $operator=@{between=1;equal=3;greater_than=5;less_than=6}[$p.operator];if(!$operator){$operator=1}
      if($p.validationType -eq 'list'){$separator=$app.International(5);$formula=$p.listItems -join $separator;if(!$formula -or $formula.Length -gt 255){throw 'Validation list must be 1..255 characters'};$range.Validation.Delete();$range.Validation.Add(3,1,1,$formula);$range.Validation.InCellDropdown=$true}
      else{$range.Validation.Delete();$range.Validation.Add(2,1,$operator,$p.minVal,$p.maxVal)}
      if($p.promptTitle){$range.Validation.InputTitle=$p.promptTitle};if($p.promptMessage){$range.Validation.InputMessage=$p.promptMessage;$range.Validation.ShowInput=$true};if($p.errorTitle){$range.Validation.ErrorTitle=$p.errorTitle};if($p.errorMessage){$range.Validation.ErrorMessage=$p.errorMessage};$range.Validation.ShowError=$true;return @{success=$true}
    }
    'add_conditional_formatting' {
      if($p.clearExisting){$range.FormatConditions.Delete()}
      switch($p.ruleType){
        'data_bar'{$rule=$range.FormatConditions.AddDatabar();if($p.barColor){$rule.BarColor.Color=Get-Color $p.barColor}}
        'color_scale'{$rule=$range.FormatConditions.AddColorScale(2);if($p.colorScaleMin){$rule.ColorScaleCriteria.Item(1).FormatColor.Color=Get-Color $p.colorScaleMin};if($p.colorScaleMax){$rule.ColorScaleCriteria.Item(2).FormatColor.Color=Get-Color $p.colorScaleMax}}
        default{$operator=@{between=1;equal=3;greater_than=5;less_than=6}[$p.operator];if(!$operator){$operator=6};$rule=$range.FormatConditions.Add(1,$operator,$p.formula1,$p.formula2);if($p.backgroundColor){$rule.Interior.Color=Get-Color $p.backgroundColor};if($p.fontColor){$rule.Font.Color=Get-Color $p.fontColor}}
      };return @{success=$true}
    }
    'manage_cell_comments' {
      if($p.action -eq 'clear_all'){$sheet.Cells.ClearComments();return @{success=$true}}
      if(!$range){throw 'address is required'};$comments=@()
      foreach($cell in $range.Cells){switch($p.action){'read'{if($cell.Comment){$comments+=@{address=$cell.Address();text=$cell.Comment.Text();author=$cell.Comment.Author}}}'add'{if($cell.Comment){$cell.Comment.Text([string]$p.text) | Out-Null}else{$cell.AddComment([string]$p.text) | Out-Null}}'delete'{if($cell.Comment){$cell.Comment.Delete()}}default{throw 'Unknown comment action'}}};return @{success=$true;comments=$comments}
    }
    'create_pivot_table' {
      $source=Get-Sheet $wb $p.sourceSheetName;$dest=Get-Sheet $wb $p.destSheetName
      $cache=$wb.PivotCaches().Create(1,$source.Range($p.sourceRange).Address($true,$true,1,$true))
      $pivot=$cache.CreatePivotTable($dest.Range($p.destCell),'BridgePivot_'+[Guid]::NewGuid().ToString('N').Substring(0,8))
      foreach($field in $p.rowFields){$pivot.PivotFields($field).Orientation=1};foreach($field in $p.columnFields){$pivot.PivotFields($field).Orientation=2}
      foreach($field in $p.dataFields){$fn=@{sum=-4157;count=-4112;average=-4106;max=-4136;min=-4139}[$field.summaryFunction];if(!$fn){$fn=-4157};$caption=$field.caption;if(!$caption){$caption='Summary '+$field.fieldName};$pivot.AddDataField($pivot.PivotFields($field.fieldName),$caption,$fn) | Out-Null};return @{success=$true;name=$pivot.Name}
    }
    'get_charts' { return Get-ExcelCharts $sheet $p }
    'delete_chart' {
      $matches=Select-ExcelCharts $sheet $p
      if(!$p.clearAll -and !$p.shapeName -and !$p.chartTitle -and !$p.chartIndex -and !$p.leftCell){throw 'Specify a chart selector or clearAll'}
      if(!$p.clearAll -and $matches.Count -gt 1){throw 'Chart selector is ambiguous'}
      $deleted=@();foreach($chart in $matches){$deleted+=$chart.Name;$chart.Delete()};return @{success=$true;deletedNames=$deleted}
    }
    'add_chart' { return Add-ExcelChart $sheet $p }
    'capture_sheet_preview' {
      if(!$range){$range=$sheet.UsedRange;Check-Range $range};$wb.Activate();$sheet.Activate();$range.CopyPicture(1,-4147)
      $temp=$null
      try{$temp=$sheet.ChartObjects().Add(0,0,[Math]::Max(100,$range.Width),[Math]::Max(100,$range.Height));$temp.Activate();$temp.Chart.Paste();$ok=$temp.Chart.Export($p.outputPath,'PNG');if(!$ok){throw 'Excel did not export preview'}}finally{if($temp){$temp.Delete()}}
      return @{success=$true;workbookName=$wb.Name;sheetName=$sheet.Name;address=$range.Address();imageBase64=[Convert]::ToBase64String([IO.File]::ReadAllBytes($p.outputPath));imagePath=$p.outputPath;imageMimeType='image/png'}
    }
    default {throw "Unsupported Excel method: $method"}
  }
}
function Select-ExcelCharts($sheet,$p) {
  $matches=@();$charts=$sheet.ChartObjects()
  for($i=1;$i -le $charts.Count;$i++){$obj=$charts.Item($i);$match=$true
    if($p.shapeName){$match=$obj.Name -ceq $p.shapeName}
    elseif($p.chartIndex){$match=$i -eq $p.chartIndex}
    elseif($p.chartTitle){$match=$obj.Chart.HasTitle -and $obj.Chart.ChartTitle.Text -ceq $p.chartTitle}
    elseif($p.leftCell){$cell=$sheet.Range($p.leftCell);$match=[Math]::Abs($obj.Left-$cell.Left) -lt 1 -and [Math]::Abs($obj.Top-$cell.Top) -lt 1}
    if($match){$matches+=,$obj}
  };return ,$matches
}
function Get-ExcelCharts($sheet,$p) {
  $items=@();foreach($obj in (Select-ExcelCharts $sheet $p)){
    $chart=$obj.Chart;$title='';if($chart.HasTitle){$title=$chart.ChartTitle.Text}
    $item=@{chartIndex=$obj.Index;leftCell=$obj.TopLeftCell.Address();seriesCount=$chart.SeriesCollection().Count;shapeName=$obj.Name;title=$title;chartType=$chart.ChartType;left=$obj.Left;top=$obj.Top;width=$obj.Width;height=$obj.Height;hasLegend=$chart.HasLegend}
    if($p.detail){$series=@();foreach($s in $chart.SeriesCollection()){$series+=@{seriesIndex=$s.Index;name=$s.Name;formula=$s.Formula;values=@($s.Values);xValues=@($s.XValues);smooth=$s.Smooth;lineColor=(Format-Color $s.Format.Line.ForeColor.RGB);fillColor=(Format-Color $s.Format.Fill.ForeColor.RGB)}};$item.series=$series;try{$axis=$chart.Axes(2,1);$axisTitle='';if($axis.HasTitle){$axisTitle=$axis.AxisTitle.Text};$item.yAxis=@{minimumScale=$axis.MinimumScale;maximumScale=$axis.MaximumScale;majorUnit=$axis.MajorUnit;numberFormat=$axis.TickLabels.NumberFormat;title=$axisTitle}}catch{$item.yAxis=$null}}
    $items+=$item
  };return @{workbookName=$sheet.Parent.Name;sheetName=$sheet.Name;charts=$items;count=$items.Count;chartCount=$items.Count;totalChartCount=$sheet.ChartObjects().Count;detail=[bool]$p.detail}
}
function Add-ExcelChart($sheet,$p) {
  $types=@{line=4;column_clustered=51;bar_clustered=57;pie=5;doughnut=-4120;pareto=122}
  $type=$types[$p.chartType];if(!$type){throw 'Unsupported chart type'}
  $address=$p.dataRange;if(!$address){$address=$p.dataRanges -join ','};$source=$sheet.Range($address)
  $anchor='G2';$width=480;$height=280;if($p.position){$anchor=$p.position.leftCell;if($p.position.width){$width=$p.position.width};if($p.position.height){$height=$p.position.height}}
  $cell=$sheet.Range($anchor);$previous=@();if($p.replaceExisting){$previous=Select-ExcelCharts $sheet ([pscustomobject]@{leftCell=$anchor})}
  $obj=$sheet.ChartObjects().Add($cell.Left,$cell.Top,$width,$height)
  try {
    $chart=$obj.Chart;$chart.ChartType=$type;$chart.SetSourceData($source);$chart.HasLegend=$p.hasLegend
    if($p.title){$chart.HasTitle=$true;$chart.ChartTitle.Text=$p.title}
    for($i=1;$i -le $chart.SeriesCollection().Count;$i++){$s=$chart.SeriesCollection($i);$s.HasDataLabels=$p.hasDataLabels;if($p.smoothLine -and $p.chartType -eq 'line'){$s.Smooth=$true};if($p.seriesColors -and $i -le $p.seriesColors.Count){$color=Get-Color $p.seriesColors[$i-1];$s.Format.Line.ForeColor.RGB=$color;$s.Format.Fill.ForeColor.RGB=$color}}
    foreach($setting in $p.seriesSettings){$s=$chart.SeriesCollection([int]$setting.seriesIndex);if($setting.color){$color=Get-Color $setting.color;$s.Format.Line.ForeColor.RGB=$color;$s.Format.Fill.ForeColor.RGB=$color};if($null -ne $setting.smooth){$s.Smooth=$setting.smooth}}
    if($p.yAxis){$axis=$chart.Axes(2);if($null -ne $p.yAxis.min){$axis.MinimumScale=$p.yAxis.min};if($null -ne $p.yAxis.max){$axis.MaximumScale=$p.yAxis.max};if($null -ne $p.yAxis.step){$axis.MajorUnit=$p.yAxis.step};if($p.yAxis.numberFormat){$axis.TickLabels.NumberFormat=$p.yAxis.numberFormat};if($p.yAxis.title){$axis.HasTitle=$true;$axis.AxisTitle.Text=$p.yAxis.title}}
  }catch{$obj.Delete();throw}
  foreach($old in $previous){$old.Delete()}
  return @{success=$true;shapeName=$obj.Name;left=$obj.Left;top=$obj.Top;width=$obj.Width;height=$obj.Height}
}

function Read-ExcelStyle($range,$include) {
  $style=@{}
  foreach($field in $include) {
    try {
      $value=switch($field){
        'fontName'{$range.Font.Name} 'fontSize'{$range.Font.Size} 'bold'{$range.Font.Bold}
        'fontColor'{Format-Color $range.Font.Color} 'backgroundColor'{Format-Color $range.Interior.Color}
        'numberFormat'{$range.NumberFormat} 'rowHeight'{$range.RowHeight} 'columnWidth'{$range.ColumnWidth}
        'wrapText'{$range.WrapText} 'merged'{$range.MergeCells}
        'mergeArea'{if($range.MergeCells -eq $true){$range.Cells.Item(1,1).MergeArea.Address()}else{$null}}
        'horizontalAlignment'{$v=$range.HorizontalAlignment;$map=@{'-4131'='left';'-4108'='center';'-4152'='right';'1'='general'};if($map.ContainsKey([string]$v)){$map[[string]$v]}else{$v}}
        'verticalAlignment'{$v=$range.VerticalAlignment;$map=@{'-4160'='top';'-4108'='center';'-4107'='bottom'};if($map.ContainsKey([string]$v)){$map[[string]$v]}else{$v}}
        'borders'{$borders=@{};foreach($pair in @(@('left',7),@('top',8),@('bottom',9),@('right',10),@('insideHorizontal',11),@('insideVertical',12))){$b=$range.Borders.Item($pair[1]);$borders[$pair[0]]=@{lineStyle=$b.LineStyle;weight=$b.Weight;color=(Format-Color $b.Color)}};$borders}
      }
      $style[$field]=$value
    }catch{$style[$field]=$null}
  }
  return $style
}
