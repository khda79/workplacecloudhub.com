@echo off
setlocal

set "UNC_WORK_DIR=%~dp0."

set "SCRIPT_DIR=%UNC_WORK_DIR%\"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartCitrix-OnPrem-StoreFront-Inventory.ps1" -Tenant prod %*
set "EXIT_CODE=%ERRORLEVEL%"
exit /b %EXIT_CODE%
