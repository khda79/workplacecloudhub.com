@echo off
setlocal

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "SCRIPT_DIR=%CD%\"
set "POWERSHELL5=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if exist "%POWERSHELL5%" (
start "SmartM365 Device Reboot Manager Test" "%POWERSHELL5%" -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-DeviceRebootManager-GUI.ps1" -TestRequiredRestart -PreviewOnly %*
) else (
start "SmartM365 Device Reboot Manager Test" powershell.exe -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-DeviceRebootManager-GUI.ps1" -TestRequiredRestart -PreviewOnly %*
)
popd
exit /b 0
