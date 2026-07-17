@echo off
setlocal EnableExtensions

set "UNC_WORK_DIR=%~dp0."

set "ROOT_DIR=%UNC_WORK_DIR%\"
set "SCRIPT=%ROOT_DIR%Scripts\SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1"
call :PrintStartupInfo

if not exist "%SCRIPT%" (
    echo ERROR: GUI script not found:
    echo   %SCRIPT%
    exit /b 1
)

start "SmartM365 LOT Launcher" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "%SCRIPT%" %*
exit /b 0

:PrintStartupInfo
echo.
echo SmartM365 Intune Hybrid Join Toolkit - GUI launcher
echo Started : %DATE% %TIME%
echo Root    : %ROOT_DIR%
echo Script  : %SCRIPT%
echo Mode    : Detached WPF GUI, hidden PowerShell host
echo.
exit /b 0
