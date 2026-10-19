@echo off
setlocal EnableExtensions
rem Runs the V2 valuation, then the report only when the valuation succeeded (0) or completed with warnings (3),
rem so that a report is never built on an older valuation. Exit code: the most severe of the two (1 > 3 > 0).
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" (
    echo [%date% %time%] PowerShell 7 x64 is required.
    exit /b 1
)
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0SmartFinOps-Workplace-Valuation.ps1" -Tenant prod
set "VALUATION_EXIT=%ERRORLEVEL%"
if not "%VALUATION_EXIT%"=="0" if not "%VALUATION_EXIT%"=="3" (
    echo [%date% %time%] Valuation failed with exit code %VALUATION_EXIT%: the report is not run, so no report is built on an older valuation.
    exit /b 1
)
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0SmartFinOps-Workplace-Report.ps1" -Tenant prod
set "REPORT_EXIT=%ERRORLEVEL%"
if not "%REPORT_EXIT%"=="0" if not "%REPORT_EXIT%"=="3" (
    echo [%date% %time%] Report failed with exit code %REPORT_EXIT%.
    exit /b 1
)
if "%VALUATION_EXIT%"=="3" exit /b 3
if "%REPORT_EXIT%"=="3" exit /b 3
exit /b 0
