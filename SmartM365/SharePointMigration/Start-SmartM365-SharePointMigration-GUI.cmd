@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%SmartM365-SharePointMigration-GUI.ps1"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" set "PWSH=pwsh.exe"

"%PWSH%" -NoProfile -Command "if ($PSVersionTable.PSVersion -ge [version]'7.4') { exit 0 } else { exit 1 }" >nul 2>&1
if errorlevel 1 (
    echo PowerShell 7.4 or later is required to run the GUI.
    pause
    exit /b 1
)

start "SmartM365 SharePoint Migration" "%PWSH%" -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*
exit /b 0
