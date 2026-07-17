@echo off
setlocal EnableExtensions DisableDelayedExpansion
set "UNC_WORK_DIR=%~dp0..\..\M365Inventory\SharePoint\."
set "SCRIPT_DIR=%UNC_WORK_DIR%\"
set "PWSH=C:\Program Files\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" (
  echo PowerShell 7 was not found at "%PWSH%".
  echo Install PowerShell 7 or update this launcher.
  pause
  exit /b 1
)
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-SPO-Inventory.ps1" -Tenant prod %*
set "EXITCODE=%ERRORLEVEL%"
exit /b %EXITCODE%
