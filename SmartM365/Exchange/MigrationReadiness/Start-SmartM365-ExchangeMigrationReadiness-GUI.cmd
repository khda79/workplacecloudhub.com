@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PWSH=C:\Program Files\PowerShell\7\pwsh.exe"

if not exist "%PWSH%" (
    echo PowerShell 7 was not found at "%PWSH%".
    pause
    exit /b 1
)

start "" "%PWSH%" -NoLogo -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-ExchangeMigrationReadiness-GUI.ps1"
exit /b 0
