@echo off
setlocal

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "SCRIPT_DIR=%CD%\"
set "SCRIPT=%SCRIPT_DIR%SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1"
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%SCRIPT%" %*
set "EXIT_CODE=0"
popd
exit /b %EXIT_CODE%
