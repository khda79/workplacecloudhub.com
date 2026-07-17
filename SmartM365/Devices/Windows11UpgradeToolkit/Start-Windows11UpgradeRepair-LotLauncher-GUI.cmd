@echo off
setlocal EnableExtensions

set "UNC_WORK_DIR=%~dp0."

set "ROOT_DIR=%UNC_WORK_DIR%\"
set "SCRIPT=%ROOT_DIR%Scripts\SmartM365-Windows11Upgrade-LotLauncher-GUI.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: GUI script not found:
    echo   %SCRIPT%
    exit /b 1
)

start "SmartM365 Windows 11 Upgrade LOT Launcher" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "%SCRIPT%" %*
exit /b 0
