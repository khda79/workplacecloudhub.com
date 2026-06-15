@echo off
setlocal EnableExtensions

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "ROOT_DIR=%CD%\"
set "SCRIPT=%ROOT_DIR%Scripts\SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: GUI script not found:
    echo   %SCRIPT%
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT%" %*
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
