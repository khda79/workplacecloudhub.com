@echo off
setlocal
set "DemoHost=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%DemoHost%" (
  echo Required PowerShell host not found: %DemoHost%
  pause
  exit /b 1
)
start "" "%DemoHost%" -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0SmartM365-ExchangeUserNotifications-GUI.ps1" -Demo %*
