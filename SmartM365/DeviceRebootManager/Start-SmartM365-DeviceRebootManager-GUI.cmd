@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "POWERSHELL5=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if exist "%POWERSHELL5%" (
    start "" /min "%POWERSHELL5%" -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-DeviceRebootManager-GUI.ps1" %*
) else (
    start "" /min powershell.exe -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-DeviceRebootManager-GUI.ps1" %*
)
