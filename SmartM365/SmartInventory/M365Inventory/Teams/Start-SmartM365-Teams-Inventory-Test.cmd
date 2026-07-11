@echo off
setlocal EnableExtensions EnableDelayedExpansion
pushd "%~dp0" >nul 2>&1
if errorlevel 1 exit /b 1
set "SCRIPT_DIR=%CD%\"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" for /f "delims=" %%I in ('where pwsh.exe 2^>nul') do if not defined PWSH set "PWSH=%%I"
if not exist "%PWSH%" (
  echo PowerShell 7 ^(pwsh.exe^) was not found.
  pause
  popd
  exit /b 1
)
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-Teams-Inventory.ps1" -Tenant test -MaxItems 25
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%