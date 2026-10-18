@echo off
setlocal EnableExtensions
rem Builds the V2 executive report from the latest prod valuation. Exit code: 0 success, 3 completed with warnings, 1 failure.
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" (
    echo [%date% %time%] PowerShell 7 x64 is required.
    exit /b 1
)
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0SmartFinOps-Workplace-Report.ps1" -Tenant prod %*
exit /b %ERRORLEVEL%
