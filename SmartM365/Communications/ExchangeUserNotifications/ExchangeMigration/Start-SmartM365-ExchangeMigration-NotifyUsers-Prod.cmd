@echo off
setlocal

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

pwsh -NoProfile -ExecutionPolicy Bypass -File "%CD%\SmartM365-ExchangeMigration-NotifyUsers.ps1" -Tenant prod -WhatIf
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
