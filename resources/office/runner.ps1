$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$ProgressPreference = 'SilentlyContinue'
. (Join-Path $PSScriptRoot 'common.ps1')

try {
  $payload = [Console]::In.ReadToEnd() | ConvertFrom-Json
  if ($payload.action -eq 'status') {
    $result = @{runningComponents=@{};openDocuments=@{};errors=@{}}
    foreach ($component in @('excel','word','ppt')) {
      try {
        $app=Get-OfficeApp $component
        $collection=$null
        switch($component){'excel'{$collection=$app.Workbooks}'word'{$collection=$app.Documents}'ppt'{$collection=$app.Presentations}}
        $result.runningComponents[$component]=$true
        $result.openDocuments[$component]=@($collection | ForEach-Object {$_.Name})
      } catch { $result.runningComponents[$component]=$false; $result.openDocuments[$component]=@(); $result.errors[$component]=$_.Exception.Message }
    }
  } elseif ($payload.action -eq 'clipboard') {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $img=[Windows.Forms.Clipboard]::GetImage()
    if(!$img){throw 'Clipboard has no image'}
    $ms=New-Object IO.MemoryStream
    try{$img.Save($ms,[Drawing.Imaging.ImageFormat]::Png);$result=@{imageBase64=[Convert]::ToBase64String($ms.ToArray())}}finally{$ms.Dispose();$img.Dispose()}
  } elseif ($payload.action -eq 'excel') {
    . (Join-Path $PSScriptRoot 'excel.ps1')
    $result=Invoke-ExcelTool $payload.method $payload.params
  } else {
    $app=Get-OfficeApp $payload.component
    $target=Get-Target $app $payload.component $payload.targetName
    $doc=$target; $wb=$target; $pres=$target; $params=$payload.params
    if ($payload.action -eq 'export') {
      if ($payload.slideIndex -lt 1 -or $payload.slideIndex -gt $pres.Slides.Count) { throw 'Slide index out of range' }
      $pres.Slides.Item($payload.slideIndex).Export($payload.outputPath,'PNG',1280,720)
      $result=@{imagePath=$payload.outputPath}
    } elseif ($payload.action -eq 'script') {
      $result=& ([ScriptBlock]::Create($payload.script))
    } else { throw 'Unknown action' }
  }
  ConvertTo-Json -InputObject @{success=$true;data=$result} -Depth 70 -Compress
} catch {
  ConvertTo-Json -InputObject @{success=$false;error=$_.Exception.Message} -Depth 8 -Compress

}
