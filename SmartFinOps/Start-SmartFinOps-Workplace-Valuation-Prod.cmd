@echo off
setlocal EnableExtensions
rem Runs the V2 valuation engine for the prod profile. Exit code: 0 success, 3 completed with warnings, 1 failure.
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" (
    echo [%date% %time%] PowerShell 7 x64 is required.
    exit /b 1
)
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0SmartFinOps-Workplace-Valuation.ps1" -Tenant prod %*
exit /b %ERRORLEVEL%
