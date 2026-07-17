@echo off
setlocal

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "SCRIPT=%CD%\SmartM365-ExchangeUserNotifications-GUI.ps1"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"

if not exist "%SCRIPT%" (
    echo ERROR: GUI script not found:
    echo   %SCRIPT%
    pause
    popd
    exit /b 1
)

if exist "%PWSH%" (
    start "SmartM365 Exchange User Notifications" "%PWSH%" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "%SCRIPT%" %*
) else (
    start "SmartM365 Exchange User Notifications" pwsh.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "%SCRIPT%" %*
)

popd
exit /b 0
