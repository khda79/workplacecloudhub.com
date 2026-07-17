@echo off
setlocal EnableExtensions EnableDelayedExpansion
set "UNC_WORK_DIR=%~dp0..\..\M365Inventory\Teams\."
set "SCRIPT_DIR=%UNC_WORK_DIR%\"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" for /f "delims=" %%I in ('where pwsh.exe 2^>nul') do if not defined PWSH set "PWSH=%%I"
if not exist "%PWSH%" (
  echo PowerShell 7 ^(pwsh.exe^) was not found.
  pause
  exit /b 1
)
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-Teams-Inventory.ps1" -Tenant prod
set "EXIT_CODE=%ERRORLEVEL%"
exit /b %EXIT_CODE%
