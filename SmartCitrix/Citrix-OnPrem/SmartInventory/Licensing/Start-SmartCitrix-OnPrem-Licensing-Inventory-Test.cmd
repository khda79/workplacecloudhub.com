@echo off
set "SCRIPT_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartCitrix-OnPrem-Licensing-Inventory.ps1" -Tenant test %*
