@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "SCRIPT=%SCRIPT_DIR%SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1"

start "" powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%SCRIPT%" %*
exit /b 0
