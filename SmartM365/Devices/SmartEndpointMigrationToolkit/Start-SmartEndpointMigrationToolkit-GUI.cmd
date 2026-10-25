@echo off
rem Smart Endpoint Migration Toolkit - GUI launcher (Windows PowerShell 5.1, STA).
setlocal DisableDelayedExpansion
set "SEMT_GUI=%~dp0Scripts\SmartM365-EndpointMigration-GUI.ps1"
set "SEMT_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%SEMT_GUI%" (
    echo ERROR: GUI script not found: %SEMT_GUI%
    pause
    exit /b 1
)
start "Smart Endpoint Migration Toolkit" "%SEMT_PS%" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "%SEMT_GUI%" %*
exit /b 0
