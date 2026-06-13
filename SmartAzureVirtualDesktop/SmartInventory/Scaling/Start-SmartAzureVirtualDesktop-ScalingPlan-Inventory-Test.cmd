@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
where pwsh.exe >nul 2>nul
if %ERRORLEVEL%==0 (
  pwsh.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartAzureVirtualDesktop-ScalingPlan-Inventory.ps1" -Tenant test -Connect %*
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartAzureVirtualDesktop-ScalingPlan-Inventory.ps1" -Tenant test -Connect %*
)
