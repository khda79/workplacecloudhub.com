@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%SmartM365-SharePointMigration-GUI.ps1"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"

if exist "%PWSH%" (
    "%PWSH%" -NoProfile -STA -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*
) else (
    powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*
)

set "EXIT_CODE=%ERRORLEVEL%"
if %EXIT_CODE% NEQ 0 (
    echo.
    echo Exit code: %EXIT_CODE%
    pause
)
exit /b %EXIT_CODE%
