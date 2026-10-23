@echo off
setlocal
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -File "%~dp0Start-Local-Debug.ps1"
set "CHARFLOAT_DEBUG_EXIT=%ERRORLEVEL%"
echo.
echo Diagnostic session ended. Keep the report shown above.
pause
exit /b %CHARFLOAT_DEBUG_EXIT%
