@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%SmartM365-DeviceRegistration-Tool.ps1"
set "POWERSHELL5=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if exist "%POWERSHELL5%" (
    start "" /min "%POWERSHELL5%" -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -Cli -Mode User -JsonOutput -SupportBundle %*
) else (
    start "" /min powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -Cli -Mode User -JsonOutput -SupportBundle %*
)
exit /b 0
