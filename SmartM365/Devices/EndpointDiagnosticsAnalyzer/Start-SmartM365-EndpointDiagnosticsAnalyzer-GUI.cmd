@echo off
setlocal

set "UNC_WORK_DIR=%~dp0."

set "SCRIPT_DIR=%UNC_WORK_DIR%\"
set "SCRIPT=%SCRIPT_DIR%SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1"
start "SmartM365 Endpoint Diagnostics Analyzer" powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%SCRIPT%" %*
exit /b 0
