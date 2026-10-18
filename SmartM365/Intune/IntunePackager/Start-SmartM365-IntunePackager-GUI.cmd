@echo off
setlocal
rem Smart IntunePackager GUI launcher. Requires PowerShell 7.
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" (
  echo PowerShell 7 is required: %PWSH%
  pause
  exit /b 1
)
start "Smart IntunePackager" "%PWSH%" -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0SmartM365-IntunePackager-GUI.ps1" %*
exit /b 0
