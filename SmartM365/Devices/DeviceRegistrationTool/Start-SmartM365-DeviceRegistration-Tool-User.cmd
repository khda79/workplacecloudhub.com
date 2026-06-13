@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%SmartM365-DeviceRegistration-Tool.ps1"
set "POWERSHELL5=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if exist "%POWERSHELL5%" (
    start "" /min "%POWERSHELL5%" -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -Mode User %*
) else (
    start "" /min powershell.exe -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -Mode User %*
)
