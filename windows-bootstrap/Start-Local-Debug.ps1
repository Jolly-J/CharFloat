# Windows PowerShell 5.1; explicit user action launches only the verified local-debug EXE.
# No elevation, global environment changes, execution-policy changes, network or document access.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$runId = [Guid]::NewGuid().ToString()
$base = [Environment]::GetFolderPath('LocalApplicationData')
$profilePath = [Environment]::GetFolderPath('UserProfile')
$runDir = Join-Path (Join-Path $base 'CharFloatStartupDiagnostics') $runId
$reportPath = Join-Path $runDir 'CharFloat-Debug-Report.json'
$prefixes = @($env:LOCALAPPDATA, $env:APPDATA, $env:USERPROFILE, $env:TEMP, $env:TMP) | Where-Object { $_ } | Sort-Object Length -Descending
function Protect-Text([string]$text) {
  foreach ($prefix in $prefixes) { $text = [regex]::Replace($text, [regex]::Escape($prefix), '<user-path>', 'IgnoreCase') }
  $text = [regex]::Replace($text, '(?i)\b(Bearer\s+)[a-z0-9._~+/-]+', '$1<redacted>')
  return [regex]::Replace($text, '(?i)\b(token|password|cookie|authorization|api[_-]?key)\s*[:=]\s*[^\s,;]+', '$1=<redacted>')
}
function Assert-SafeDirectory([string]$target) {
  $current = [IO.Path]::GetFullPath($target)
  for ($depth = 0; $depth -lt 64; $depth++) {
    if ([IO.Directory]::Exists($current) -or [IO.File]::Exists($current)) {
      $attributes = [IO.File]::GetAttributes($current)
      if (($attributes -band [IO.FileAttributes]::ReparsePoint) -or -not ($attributes -band [IO.FileAttributes]::Directory)) { throw 'Diagnostic path is redirected or is not a directory; no launch performed.' }
    }
    $parent = [IO.Path]::GetDirectoryName($current)
    if (-not $parent -or $parent -eq $current) { return }
    $current = $parent
  }
  throw 'Diagnostic directory depth exceeded.'
}
$report = [ordered]@{ runId=$runId; startedAt=[DateTime]::UtcNow.ToString('o'); status='preparing'; launchesVerifiedInstaller=$true; launchAttempted=$false; childStarted=$false;
  process64Bit=[Environment]::Is64BitProcess; os64Bit=[Environment]::Is64BitOperatingSystem;
  coverage=[ordered]@{ nativeNSIS='Process path/hash/pid/exit and available stdout/stderr only; native MessageBox and Win32 calls are not traced.'; application='Bootstrap lifecycle only; installed application and Office documents are not monitored.'; networkUpload=$false; screenRecording=$false } }
$utf8 = New-Object Text.UTF8Encoding($false)
function Save-Report {
  if ($outText) { $report['stdout']=Protect-Text $outText.ToString() }
  if ($errText) { $report['stderr']=Protect-Text $errText.ToString() }
  $report['updatedAt'] = [DateTime]::UtcNow.ToString('o')
  $report['bootstrapLogPresent'] = [IO.File]::Exists((Join-Path $runDir 'bootstrap.jsonl'))
  $events = New-Object Collections.Generic.List[object]
  if ($report['bootstrapLogPresent']) {
    try {
      foreach ($line in [IO.File]::ReadAllLines((Join-Path $runDir 'bootstrap.jsonl'))) {
        if ($events.Count -ge 10000) { break }
        try { $item = $line | ConvertFrom-Json; if ($item.runId -eq $runId) { $events.Add($item) } } catch { }
      }
    } catch { $report['logReadWarning'] = Protect-Text $_.Exception.Message }
  }
  $report['bootstrapEvents'] = @($events.ToArray())
  $report['coverage']['bootstrapLifecycle'] = $(if ($events.Count -gt 0) { 'observed-events-see-report' } else { 'not-observed-check-stderr-and-native-coverage-gap' })
  $finalFile = Join-Path $runDir 'bootstrap-final.json'
  if ([IO.File]::Exists($finalFile)) { try { $report['bootstrapFinal'] = [IO.File]::ReadAllText($finalFile) | ConvertFrom-Json } catch { } }
  $json = $report | ConvertTo-Json -Depth 14
  [IO.File]::WriteAllText($reportPath, $json, $utf8)
}
function Append-Bounded($builder, [char[]]$buffer, [int]$count) {
  $remaining = 65536 - $builder.Length
  if ($remaining -gt 0) { [void]$builder.Append($buffer, 0, [Math]::Min($remaining,$count)) }
}
$child = $null
$outText = $null; $errText = $null
try {
  $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = New-Object Security.Principal.WindowsPrincipal($identity)
  $report['elevated'] = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if ($report['elevated']) { throw 'Please close this window and run the diagnostic normally, without administrator privileges.' }
  Assert-SafeDirectory $runDir
  [void][IO.Directory]::CreateDirectory($runDir)
  Write-Host ('Report (also updated while running): ' + $reportPath)
  $report['temp'] = [ordered]@{ tmp=Protect-Text ([string]$env:TMP); temp=Protect-Text ([string]$env:TEMP); dotNet=Protect-Text ([IO.Path]::GetTempPath()) }
  $package = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'debug-package.json')) | ConvertFrom-Json
  if ($package.executable -ne 'CharFloat-2.2.3-win-x64-local-debug.exe') { throw 'Unexpected executable in package manifest.' }
  $exe = Join-Path $PSScriptRoot $package.executable
  $info = Get-Item -LiteralPath $exe
  if ($info.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Executable is a redirected link; no launch performed.' }
  $hash = (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash.ToLowerInvariant()
  $report['installer'] = [ordered]@{ path=Protect-Text $exe; bytes=$info.Length; sha256=$hash; expectedSha256=$package.sha256; version=[Diagnostics.FileVersionInfo]::GetVersionInfo($exe).FileVersion; pathCharacters=$exe.Length; nonAscii=($exe -match '[^\x00-\x7f]'); attributes=[string]$info.Attributes }
  if ($hash -ne $package.sha256 -or $info.Length -ne $package.bytes) { throw 'EXE hash/size does not match this debug package; no launch performed.' }
  $report['installer']['zoneId'] = $null
  try { foreach ($line in [IO.File]::ReadAllLines($exe + ':Zone.Identifier')) { if ($line -match '^ZoneId=(\d+)$') { $report['installer']['zoneId']=[int]$Matches[1];break } } } catch { }
  $report['status']='verified'; Save-Report
  $start = New-Object Diagnostics.ProcessStartInfo
  $start.FileName=$exe; $start.WorkingDirectory=$PSScriptRoot; $start.UseShellExecute=$false; $start.CreateNoWindow=$true
  $start.RedirectStandardOutput=$true; $start.RedirectStandardError=$true
  # These variables exist only in this child process, never in the user/system environment.
  $start.EnvironmentVariables['CHARFLOAT_DIAGNOSTIC_RUN_ID']=$runId
  $start.EnvironmentVariables['CHARFLOAT_DIAGNOSTIC_DIR']=$runDir
  $child=New-Object Diagnostics.Process; $child.StartInfo=$start
  $report['launchAttempted']=$true
  if (-not $child.Start()) { throw 'Process creation returned false.' }
  $report['childStarted']=$true
  $report['pid']=$child.Id; $report['childStartedAt']=[DateTime]::UtcNow.ToString('o'); $report['status']='running'
  Write-Host 'The verified setup is running. Use its window. Do not start a second copy.'
  $outText=New-Object Text.StringBuilder; $errText=New-Object Text.StringBuilder
  $outBuffer=New-Object char[] 4096; $errBuffer=New-Object char[] 4096
  $outTask=$child.StandardOutput.ReadAsync($outBuffer,0,$outBuffer.Length); $errTask=$child.StandardError.ReadAsync($errBuffer,0,$errBuffer.Length)
  $outDone=$false; $errDone=$false; $exitAt=$null; $saved=[DateTime]::MinValue
  while ($true) {
    if (-not $outDone -and $outTask.IsCompleted) { $n=$outTask.GetAwaiter().GetResult(); if ($n -eq 0) { $outDone=$true } else { Append-Bounded $outText $outBuffer $n; $outTask=$child.StandardOutput.ReadAsync($outBuffer,0,$outBuffer.Length) } }
    if (-not $errDone -and $errTask.IsCompleted) { $n=$errTask.GetAwaiter().GetResult(); if ($n -eq 0) { $errDone=$true } else { Append-Bounded $errText $errBuffer $n; $errTask=$child.StandardError.ReadAsync($errBuffer,0,$errBuffer.Length) } }
    if ($child.HasExited -and -not $exitAt) { $exitAt=[DateTime]::UtcNow; $report['exitCode']=$child.ExitCode; $report['childExitedAt']=$exitAt.ToString('o'); $report['status']='exited' }
    if (([DateTime]::UtcNow-$saved).TotalSeconds -ge 2) { Save-Report; $saved=[DateTime]::UtcNow }
    if ($exitAt -and (($outDone -and $errDone) -or ([DateTime]::UtcNow-$exitAt).TotalSeconds -ge 3)) { break }
    Start-Sleep -Milliseconds 100
  }
  $report['stdout']=Protect-Text $outText.ToString(); $report['stderr']=Protect-Text $errText.ToString()
  $report['captureCapped']=($outText.Length -ge 65536 -or $errText.Length -ge 65536); $report['pipeDrainIncomplete']=(-not $outDone -or -not $errDone)
} catch {
  $report['status']='diagnostic-error'; $report['error']=Protect-Text $_.Exception.Message
  if ($_.Exception -is [ComponentModel.Win32Exception]) { $report['nativeErrorCode']=$_.Exception.NativeErrorCode }
  Write-Host $report['error'] -ForegroundColor Red
} finally {
  if ($outText) { $report['stdout']=Protect-Text $outText.ToString() }
  if ($errText) { $report['stderr']=Protect-Text $errText.ToString() }
  if ($child -and $report['childStarted']) { try { $report['childStillRunning']=(-not $child.HasExited) } catch { } }
  if ([IO.Directory]::Exists($runDir)) {
    try { Save-Report; Write-Host ('Send this one report: ' + $reportPath) -ForegroundColor Cyan } catch { Write-Host ('Report save failed: ' + $_.Exception.Message) -ForegroundColor Red }
  }
  if ($child) { $child.Dispose() }
}
if ($report['status'] -eq 'diagnostic-error') { exit 1 }
exit 0
