@echo off
setlocal EnableExtensions
set "SCRIPT_DIR=%~dp0"
set "PWSH=C:\Program Files\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" set "PWSH=pwsh.exe"
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-ExchangeSE-ADPreCheck-Inventory.ps1" -Tenant test %*
exit /b %ERRORLEVEL%
