@echo off
setlocal

set "UNC_WORK_DIR=%~dp0."

set "SCRIPT_DIR=%UNC_WORK_DIR%\"
set "SCRIPT_PATH=%SCRIPT_DIR%SmartM365-DeviceRegistration-Tool.ps1"
set "POWERSHELL5=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if exist "%POWERSHELL5%" (
"%POWERSHELL5%" -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -Cli -Mode User -JsonOutput -SupportBundle %*
) else (
powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -Cli -Mode User -JsonOutput -SupportBundle %*
)
set "EXIT_CODE=0"
exit /b %EXIT_CODE%
