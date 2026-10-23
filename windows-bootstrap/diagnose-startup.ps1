# 只诊断 NSIS 启动前的 Win32 临时目录 API；不安装、不启动应用、不改系统环境或策略。
param([string]$InstallerPath = (Join-Path $PSScriptRoot 'CharFloat-2.2.3-win-x64.exe'))
$ErrorActionPreference = 'Stop'
$expectedHash = '5be48ab24c2a2f0f7accc2235182f67dbe1427519c730e03922b9740d5b5b88a'
$originalTemp = [Environment]::GetEnvironmentVariable('TEMP', 'Process')
$originalTmp = [Environment]::GetEnvironmentVariable('TMP', 'Process')
$local = [Environment]::GetFolderPath('LocalApplicationData')
$profile = [Environment]::GetFolderPath('UserProfile')
$windows = [Environment]::GetFolderPath('Windows')
$report = [ordered]@{
  schema = 1; diagnosticOnly = $true; installsOrLaunchesApplication = $false
  utc = [DateTime]::UtcNow.ToString('o'); process64Bit = [Environment]::Is64BitProcess
  os64Bit = [Environment]::Is64BitOperatingSystem
  tempEnvironment = $originalTemp; tmpEnvironment = $originalTmp
  tempCharacters = ([string]$originalTemp).Length; tmpCharacters = ([string]$originalTmp).Length
  tempHasNonAscii = ([string]$originalTemp -match '[^\x00-\x7f]'); tmpHasNonAscii = ([string]$originalTmp -match '[^\x00-\x7f]')
  win32TempPriority = 'TMP, TEMP, USERPROFILE, Windows directory'
  expectedInstallerSha256 = $expectedHash
  expectedInstallerBytes = 199039041
}
$run = $null
$nativeSource = @'
using System;
using System.Text;
using System.Runtime.InteropServices;
namespace CharFloatStartup {
  public static class Native {
    [DllImport("kernel32.dll", EntryPoint="GetTempPathW", ExactSpelling=true, CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern uint GetTempPath(uint count, StringBuilder buffer);
    [DllImport("kernel32.dll", EntryPoint="GetTempFileNameW", ExactSpelling=true, CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern uint GetTempFileName(string path, string prefix, uint unique, StringBuilder buffer);
    [DllImport("kernel32.dll", EntryPoint="GetFileAttributesW", ExactSpelling=true, CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern uint GetFileAttributes(string path);
    [DllImport("kernel32.dll", EntryPoint="DeleteFileW", ExactSpelling=true, CharSet=CharSet.Unicode, SetLastError=true)]
    [return: MarshalAs(UnmanagedType.Bool)] public static extern bool DeleteFile(string path);
    [DllImport("kernel32.dll", ExactSpelling=true)] public static extern uint GetTickCount();
    [DllImport("kernel32.dll", ExactSpelling=true)] public static extern void SetLastError(uint error);
  }
}
'@
function ErrorInfo($errorRecord) {
  $exception = $errorRecord.Exception
  while ($exception.InnerException) { $exception = $exception.InnerException }
  $unsigned = ([int64]$exception.HResult) -band 4294967295L
  $result = @{ message = $exception.Message; hresult = ('0x{0:X8}' -f $unsigned) }
  if ((($unsigned -shr 16) -band 8191) -eq 7) { $result.win32Error = $unsigned -band 65535 }
  return $result
}
function IsUnder([string]$child, [string]$parent) {
  if ([string]::IsNullOrWhiteSpace($child) -or [string]::IsNullOrWhiteSpace($parent)) { return $false }
  try {
    $c = [IO.Path]::GetFullPath($child).TrimEnd('\') + '\'
    $p = [IO.Path]::GetFullPath($parent).TrimEnd('\') + '\'
    return $c.StartsWith($p, [StringComparison]::OrdinalIgnoreCase)
  } catch { return $false }
}
function WriteRisk([string]$candidate) {
  if (!(IsUnder $candidate $local) -and !(IsUnder $candidate $profile)) { return 'Outside confirmed user directories; no write.' }
  if (IsUnder $candidate $windows) { return 'System directory; no write.' }
  try {
    $current = [IO.Path]::GetFullPath($candidate)
    if ($current.StartsWith('\\')) { return 'Network/extended namespace path not verified; no write.' }
    for ($depth = 0; $current -and $depth -lt 128; $depth++) {
      try {
        $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { return ('Reparse point in path; no write: ' + $current) }
      } catch { if ($_.CategoryInfo.Category -ne 'ObjectNotFound') { return ('Path metadata unreadable; no write: ' + $current) } }
      $parent = [IO.Directory]::GetParent($current)
      if (!$parent) { return $null }; $current = $parent.FullName
    }
    if ($current) { return 'Path ancestor budget exceeded; no write.' }
  } catch { return 'Path cannot be verified; no write.' }
  return $null
}
function ProbeTemp([string]$candidate) {
  $result = [ordered]@{ path = $candidate; characters = $candidate.Length; exceedsLegacyMaxPathMinus14 = ($candidate.Length -gt 246) }
  $risk = WriteRisk $candidate
  if ($risk) { $result.skipped = $risk; return $result }
  [CharFloatStartup.Native]::SetLastError(0)
  $attr = [CharFloatStartup.Native]::GetFileAttributes($candidate)
  $attrError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
  $result.attributes = ('0x{0:X8}' -f $attr); $result.attributesError = $attrError
  $failures = @{}
  # 与 NSIS 3.04 的 ns[a-z] 前缀和最多100次 GetTempFileNameW 保持一致。
  for ($attempt = 1; $attempt -le 100; $attempt++) {
    $prefix = 'ns' + [char](97 + ([CharFloatStartup.Native]::GetTickCount() % 26))
    $buffer = New-Object System.Text.StringBuilder 1024
    [CharFloatStartup.Native]::SetLastError(0)
    $value = [CharFloatStartup.Native]::GetTempFileName($candidate, $prefix, 0, $buffer)
    $nativeError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
    if ($value -ne 0) {
      $result.success = $true; $result.attempts = $attempt
      # API的uUnique=0只返回本次新创建文件；只删除该文件。
      [CharFloatStartup.Native]::SetLastError(0)
      $deleted = [CharFloatStartup.Native]::DeleteFile($buffer.ToString())
      $deleteError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
      $result.ownProbeFileDeleted = $deleted; $result.deleteError = $deleteError
      return $result
    }
    $key = [string]$nativeError
    if (!$failures.ContainsKey($key)) { $failures[$key] = 0 }; $failures[$key]++
  }
  $result.success = $false; $result.attempts = 100; $result.win32ErrorCounts = $failures
  return $result
}
try {
  $report.installer = @{ exists = (Test-Path -LiteralPath $InstallerPath -PathType Leaf) }
  if ($report.installer.exists) {
    $report.installer.bytes = (Get-Item -LiteralPath $InstallerPath).Length
    $report.installer.sizeMatchesExpected = ($report.installer.bytes -eq 199039041)
    $actualHash = (Get-FileHash -LiteralPath $InstallerPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $report.installer.sha256 = $actualHash; $report.installer.matchesExpected = ($actualHash -eq $expectedHash)
  }
  $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
  try {
    $principal = New-Object System.Security.Principal.WindowsPrincipal -ArgumentList $identity
    $report.elevated = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  } finally { $identity.Dispose() }
  if ($report.elevated) {
    $report.writeProbesSkipped = 'Elevated process. Close this window and use normal non-administrator PowerShell.'
    throw 'Elevated run detected; all write probes skipped. Use a normal non-administrator window.'
  }
  # 在编译互操作类型前先收集原始环境及.NET包装的临时API结果；失败也保留证据。
  try {
    $managed = [IO.Path]::GetTempPath()
    $report.managedGetTempPath = @{ path = $managed; characters = $managed.Length }
    $risk = WriteRisk $managed
    if (!$risk) {
      try {
        $probeFile = [IO.Path]::GetTempFileName()
        $report.managedGetTempFileName = @{ success = $true }
        [IO.File]::Delete($probeFile)
      } catch { $report.managedGetTempFileName = ErrorInfo $_ }
    } else { $report.managedGetTempFileName = @{ skipped = $risk } }
  } catch { $report.managedGetTempPathError = ErrorInfo $_ }
  if ([string]::IsNullOrWhiteSpace($local) -or !(Test-Path -LiteralPath $local -PathType Container) -or (IsUnder $local $windows)) { throw 'Cannot confirm an existing current-user LocalApplicationData directory.' }
  $base = Join-Path $local 'CharFloat Startup Diagnostics'
  $risk = WriteRisk $base
  if ($risk) { $report.writeProbesSkipped = $risk; throw $risk }
  if (Test-Path -LiteralPath $base) {
    $item = Get-Item -LiteralPath $base -Force
    if (!$item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Diagnostic directory is not a normal directory.' }
    $marker = Join-Path $base '.charfloat-diagnostic-owner'
    if (!(Test-Path -LiteralPath $marker -PathType Leaf)) { throw 'Existing diagnostic directory has no ownership marker; left unchanged.' }
    $markerItem = Get-Item -LiteralPath $marker -Force
    if (($markerItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -or $markerItem.Length -gt 128 -or [IO.File]::ReadAllText($marker) -ne 'charfloat-startup-diagnostic-v1') { throw 'Existing diagnostic directory is not owned by this tool; left unchanged.' }
  } else {
    [void][IO.Directory]::CreateDirectory($base)
    [IO.File]::WriteAllText((Join-Path $base '.charfloat-diagnostic-owner'), 'charfloat-startup-diagnostic-v1')
  }
  $run = Join-Path $base ([Guid]::NewGuid().ToString('N')); [void][IO.Directory]::CreateDirectory($run)
  $risk = WriteRisk $run
  if ($risk) { $run = $null; throw $risk }
  $report.reportDirectory = $run
  # Windows PowerShell 5.1 支持显式编译临时目录；不修改任何进程/用户/系统环境变量。
  if (!('CharFloatStartup.Native' -as [type])) {
    if (!(Get-Command Add-Type).Parameters.ContainsKey('CompilerParameters')) { throw 'Native probe needs the built-in Windows PowerShell 5.1 Add-Type CompilerParameters; no fallback or policy bypass attempted.' }
    $parameters = New-Object System.CodeDom.Compiler.CompilerParameters
    $parameters.GenerateInMemory = $true
    $parameters.TempFiles = New-Object System.CodeDom.Compiler.TempFileCollection -ArgumentList $run, $false
    [void]$parameters.ReferencedAssemblies.Add('System.dll')
    Add-Type -TypeDefinition $nativeSource -CompilerParameters $parameters -ErrorAction Stop
  }
  $buffer = New-Object System.Text.StringBuilder 1024
  [CharFloatStartup.Native]::SetLastError(0)
  $length = [CharFloatStartup.Native]::GetTempPath(1024, $buffer)
  $getPathError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
  $report.getTempPathW = @{ returnedLength = $length; path = $buffer.ToString(); win32Error = $getPathError; bufferCharacters = 1024 }
  $report.configuredTempProbe = ProbeTemp ($buffer.ToString())
  $report.privateUserDirectoryProbe = ProbeTemp $run
  # Windows\Temp 只读取名称，不创建目录或文件。
  $report.systemFallbackNotProbed = (Join-Path $windows 'Temp')
} catch { $report.diagnosticError = ErrorInfo $_ }
$json = $report | ConvertTo-Json -Depth 8
foreach ($pair in @(@($local, '<LOCALAPPDATA>'), @($profile, '<USERPROFILE>'))) {
  if (![string]::IsNullOrWhiteSpace($pair[0])) {
    $escaped = ConvertTo-Json -InputObject ([string]$pair[0]) -Compress
    $json = [regex]::Replace($json, [regex]::Escape($escaped.Substring(1, $escaped.Length - 2)), $pair[1], [Text.RegularExpressions.RegexOptions]::IgnoreCase)
  }
}
if ($run) {
  try {
    $risk = WriteRisk $run
    if ($risk) { throw $risk }
    $output = Join-Path $run 'startup-report.json'
    [IO.File]::WriteAllText($output, $json, (New-Object System.Text.UTF8Encoding $false))
    Write-Host ('Report saved: ' + $output)
  } catch { Write-Host ('Report file could not be saved: ' + $_.Exception.Message) }
}
Write-Output $json
Write-Host 'Diagnostic only. No installer/application was started. No user/machine environment or policy was changed.'
if ($report.Contains('diagnosticError')) { exit 2 }
