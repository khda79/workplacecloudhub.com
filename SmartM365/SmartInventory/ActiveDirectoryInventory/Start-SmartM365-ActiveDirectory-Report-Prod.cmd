@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" set "PWSH=pwsh"
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-ActiveDirectory-Report.ps1" -Tenant prod
exit /b %ERRORLEVEL%
