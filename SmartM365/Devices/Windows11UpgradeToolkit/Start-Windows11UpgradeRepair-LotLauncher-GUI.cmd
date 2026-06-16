@echo off
setlocal EnableExtensions

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "ROOT_DIR=%CD%\"
set "SCRIPT=%ROOT_DIR%Scripts\SmartM365-Windows11Upgrade-LotLauncher-GUI.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: GUI script not found:
    echo   %SCRIPT%
    exit /b 1
)

start "SmartM365 Windows 11 Upgrade LOT Launcher" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "%SCRIPT%" %*
popd
exit /b 0
