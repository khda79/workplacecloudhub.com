@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%SmartM365-SharePointMigration-GUI.ps1"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
set "POWERSHELLW=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershellw.exe"

if exist "%PWSH%" (
    start "SmartM365 SharePoint Migration" "%PWSH%" -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*
) else if exist "%POWERSHELLW%" (
    start "SmartM365 SharePoint Migration" "%POWERSHELLW%" -NoProfile -STA -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*
) else (
    start "SmartM365 SharePoint Migration" powershell.exe -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*
)
exit /b 0
