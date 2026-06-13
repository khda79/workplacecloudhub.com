@echo off
set "SCRIPT_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartCitrix-OnPrem-CVADHostingPower-Inventory.ps1" -Tenant test %*
