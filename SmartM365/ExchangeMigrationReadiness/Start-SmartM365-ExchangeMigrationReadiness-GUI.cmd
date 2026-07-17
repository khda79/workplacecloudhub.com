@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PWSH=C:\Program Files\PowerShell\7\pwsh.exe"

if not exist "%PWSH%" (
    echo PowerShell 7 was not found at "%PWSH%".
    echo Install PowerShell 7, then run this launcher again.
    pause
    exit /b 1
)

"%PWSH%" -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-ExchangeMigrationReadiness-GUI.ps1"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo Smart Exchange Migration Readiness exited with code %EXIT_CODE%.
    pause
)

exit /b %EXIT_CODE%
