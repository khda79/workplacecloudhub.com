@echo off
setlocal

set "UNC_WORK_DIR=%~dp0."

set "SCRIPT_DIR=%UNC_WORK_DIR%\"
where pwsh.exe >nul 2>nul
if %ERRORLEVEL%==0 (
  pwsh.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartAzureVirtualDesktop-SessionHostHealth-Inventory.ps1" -Tenant test -Connect %*
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartAzureVirtualDesktop-SessionHostHealth-Inventory.ps1" -Tenant test -Connect %*
)
set "EXIT_CODE=%ERRORLEVEL%"
exit /b %EXIT_CODE%
