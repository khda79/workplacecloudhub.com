@echo off
setlocal

set "UNC_WORK_DIR=%~dp0."

set "SCRIPT_DIR=%UNC_WORK_DIR%\"
set "POWERSHELL5=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if exist "%POWERSHELL5%" (
start "SmartM365 Device Reboot Manager" "%POWERSHELL5%" -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-DeviceRebootManager-GUI.ps1" %*
) else (
start "SmartM365 Device Reboot Manager" powershell.exe -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-DeviceRebootManager-GUI.ps1" %*
)
exit /b 0
