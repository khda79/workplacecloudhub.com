@echo off
setlocal
set "DemoHost=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%DemoHost%" (
  echo Required PowerShell host not found: %DemoHost%
  pause
  exit /b 1
)
start "" "%DemoHost%" -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0Scripts\SmartM365-Windows11Upgrade-LotLauncher-GUI.ps1" -Demo %*
