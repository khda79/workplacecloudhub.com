@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
pwsh -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-M365UserActivity-Inventory.ps1" -Tenant test -Period D180 -Reports All -Connect %*
endlocal
