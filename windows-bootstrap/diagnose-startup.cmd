@echo off
setlocal EnableExtensions DisableDelayedExpansion
rem Match the 32-bit NSIS stub when 32-bit PowerShell is available. No policy bypass.
set "CF_DIAGNOSTIC_PS=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%CF_DIAGNOSTIC_PS%" set "CF_DIAGNOSTIC_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%CF_DIAGNOSTIC_PS%" -NoLogo -NoProfile -File "%~dp0diagnose-startup.ps1" -InstallerPath "%~dp0CharFloat-2.2.3-win-x64.exe"
set "CF_DIAGNOSTIC_EXIT=%ERRORLEVEL%"
echo.
echo If your organization blocks this diagnostic, stop here and report that message. Do not change policy or run as administrator.
pause
exit /b %CF_DIAGNOSTIC_EXIT%
