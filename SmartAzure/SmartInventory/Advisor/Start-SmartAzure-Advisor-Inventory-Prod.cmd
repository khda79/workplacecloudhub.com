@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" set "PWSH=pwsh"
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartAzure-Advisor-Inventory.ps1" -Tenant prod -Connect
exit /b %ERRORLEVEL%
