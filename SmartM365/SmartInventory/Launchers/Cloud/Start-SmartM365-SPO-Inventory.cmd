@echo off
setlocal EnableExtensions DisableDelayedExpansion
pushd "%~dp0..\..\M365Inventory\SharePoint\" >nul 2>&1
if errorlevel 1 (
  echo Failed to access the launcher directory.
  exit /b 1
)
set "SCRIPT_DIR=%CD%\"
set "PWSH=C:\Program Files\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" (
  echo PowerShell 7 was not found at "%PWSH%".
  echo Install PowerShell 7 or update this launcher.
  pause
  popd
  exit /b 1
)
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-SPO-Inventory.ps1" -Tenant prod %*
set "EXITCODE=%ERRORLEVEL%"
pause
popd
exit /b %EXITCODE%
