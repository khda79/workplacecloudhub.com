@echo off
setlocal

set "SCRIPT=%~dp0SmartM365-ExchangeUserNotifications-GUI.ps1"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"

if not exist "%SCRIPT%" (
    echo ERROR: GUI script not found:
    echo   %SCRIPT%
    pause
    exit /b 1
)

if exist "%PWSH%" (
    start "SmartM365 Exchange User Notifications" "%PWSH%" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "%SCRIPT%" %*
) else (
    start "SmartM365 Exchange User Notifications" pwsh.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "%SCRIPT%" %*
)

exit /b 0
