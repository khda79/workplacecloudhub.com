@echo off
setlocal

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "SCRIPT_DIR=%CD%\"
where pwsh.exe >nul 2>nul
if %ERRORLEVEL%==0 (
  pwsh.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartAzureVirtualDesktop-SessionHostHealth-Inventory.ps1" -Tenant test -Connect %*
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartAzureVirtualDesktop-SessionHostHealth-Inventory.ps1" -Tenant test -Connect %*
)
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
