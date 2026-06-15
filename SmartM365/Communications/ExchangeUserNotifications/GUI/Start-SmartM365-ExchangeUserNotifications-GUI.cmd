@echo off
setlocal

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

pwsh -NoProfile -ExecutionPolicy Bypass -STA -File "%CD%\SmartM365-ExchangeUserNotifications-GUI.ps1"
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
