# Pure snapshot reader; does not launch a process, modify policy or write log files.
function Read-SharedDiagnosticText([string]$Path, [long]$MaximumBytes = 2097152) {
  $stream = $null; $reader = $null
  try {
    $share = [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, $share)
    if ($stream.Length -gt $MaximumBytes) { throw 'Diagnostic snapshot exceeds size limit.' }
    $reader = [IO.StreamReader]::new($stream, [Text.Encoding]::UTF8, $true, 4096, $false)
    return $reader.ReadToEnd()
  } finally {
    if ($reader) { $reader.Dispose() } elseif ($stream) { $stream.Dispose() }
  }
}
function New-DiagnosticSnapshotState {
  return @{ cache=@{}; currentWarning=$null; readErrors=0; recoveredReadErrors=0; incompleteTail=$false; invalidCompleteLines=0; eventLimitReached=$false; lastReadErrorAt=$null; lastRecoveryAt=$null }
}
function Update-DiagnosticSnapshot([string]$Path, [string]$RunId, $State) {
  try {
    $text = [string](Read-SharedDiagnosticText $Path)
    # Re-read the bounded snapshot next time. Never parse an unfinished UTF-8/JSON line.
    $lastNewline = $text.LastIndexOf("`n")
    $State.incompleteTail = ($lastNewline -lt $text.Length - 1)
    $State.invalidCompleteLines = 0
    if ($lastNewline -ge 0) {
      foreach ($line in $text.Substring(0,$lastNewline+1).Split([char[]]"`n", [StringSplitOptions]::RemoveEmptyEntries)) {
        try {
          $item = $line | ConvertFrom-Json
          if ($item.sequence -isnot [int] -and $item.sequence -isnot [long]) { $State.invalidCompleteLines++; continue }
          $sequence = [long]$item.sequence
          if ($item.runId -ne $RunId -or $sequence -lt 1) { continue }
          $key = [string]$sequence
          if ($State.cache.ContainsKey($key) -or $State.cache.Count -lt 10000) { $State.cache[$key]=$item }
          else { $State.eventLimitReached=$true }
        } catch { $State.invalidCompleteLines++ }
      }
    }
    if ($State.currentWarning) { $State.recoveredReadErrors++; $State.lastRecoveryAt=[DateTime]::UtcNow.ToString('o') }
    $State.currentWarning=$null
  } catch {
    $State.readErrors++; $State.currentWarning=$_.Exception.Message; $State.lastReadErrorAt=[DateTime]::UtcNow.ToString('o')
  }
}
