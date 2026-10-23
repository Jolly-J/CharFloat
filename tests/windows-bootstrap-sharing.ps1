param([string]$LogPath,[string]$RunId,[string]$ReadyPath,[string]$GatePath,[string]$ResultPath)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../windows-bootstrap/Read-DiagnosticSnapshot.ps1')
$legacyBlocked=$false
try { [void][IO.File]::ReadAllLines($LogPath) } catch [IO.IOException] { $legacyBlocked=$true }
$state=New-DiagnosticSnapshotState
Update-DiagnosticSnapshot $LogPath $RunId $state
$firstCount=$state.cache.Count; $firstIncomplete=$state.incompleteTail
[IO.File]::WriteAllText($ReadyPath,'ready')
$deadline=[DateTime]::UtcNow.AddSeconds(10)
while (-not [IO.File]::Exists($GatePath)) { if ([DateTime]::UtcNow -gt $deadline) { throw 'Sharing test gate timeout' }; Start-Sleep -Milliseconds 20 }
Update-DiagnosticSnapshot $LogPath $RunId $state
$secondCount=$state.cache.Count; $secondIncomplete=$state.incompleteTail
Update-DiagnosticSnapshot ($LogPath+'.missing') $RunId $state
$failurePreserved=($state.cache.Count -eq $secondCount -and $null -ne $state.currentWarning)
Update-DiagnosticSnapshot $LogPath $RunId $state
$report=@{legacyBlocked=$legacyBlocked;firstCount=$firstCount;firstIncomplete=$firstIncomplete;secondCount=$secondCount;secondIncomplete=$secondIncomplete;failurePreserved=$failurePreserved;recovered=$state.recoveredReadErrors;warningCleared=($null -eq $state.currentWarning);message=$state.cache['2'].message}
[IO.File]::WriteAllText($ResultPath,($report | ConvertTo-Json),[Text.UTF8Encoding]::new($false))
