@echo off
setlocal

set "UNC_WORK_DIR=%~dp0."

set "SCRIPT_DIR=%UNC_WORK_DIR%\"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" set "PWSH=pwsh"
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartAzure-NetworkExposure-Inventory.ps1" -Tenant prod -Connect
set "EXIT_CODE=%ERRORLEVEL%"
exit /b %EXIT_CODE%
