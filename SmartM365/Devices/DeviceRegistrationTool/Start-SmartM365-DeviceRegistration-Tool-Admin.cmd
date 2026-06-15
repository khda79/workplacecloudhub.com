@echo off
setlocal

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "SCRIPT_DIR=%CD%\"
set "SCRIPT_PATH=%SCRIPT_DIR%SmartM365-DeviceRegistration-Tool.ps1"
set "POWERSHELL5=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if exist "%POWERSHELL5%" (
"%POWERSHELL5%" -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -Mode Admin %*
) else (
powershell.exe -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -Mode Admin %*
)
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
