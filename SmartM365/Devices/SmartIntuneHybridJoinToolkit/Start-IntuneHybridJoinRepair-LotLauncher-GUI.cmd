@echo off
setlocal EnableExtensions

set "ROOT_DIR=%~dp0"
set "SCRIPT=%ROOT_DIR%Scripts\SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: GUI script not found:
    echo   %SCRIPT%
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT%" %*
exit /b %ERRORLEVEL%
