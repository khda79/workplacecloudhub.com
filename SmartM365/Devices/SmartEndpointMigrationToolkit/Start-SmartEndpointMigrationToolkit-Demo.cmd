@echo off
rem Smart Endpoint Migration Toolkit - offline demonstration with fictional Contoso data.
setlocal DisableDelayedExpansion
set "SEMT_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%SEMT_PS%" (
    echo Required PowerShell host not found: %SEMT_PS%
    pause
    exit /b 1
)
start "" "%SEMT_PS%" -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0Scripts\SmartM365-EndpointMigration-GUI.ps1" -Demo %*
exit /b 0
