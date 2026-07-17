@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..\..\..") do set "MIGRATION_NAME=%%~nxI"

set "UNC_WORK_DIR=%SCRIPT_DIR%..\..\..\..\..\."

set "PS_SCRIPT=%UNC_WORK_DIR%\Scripts\Launchers\Generic\SmartM365-SharePointMigration-Launcher.ps1"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"

if exist "%PWSH%" (
    "%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -MigrationName "%MIGRATION_NAME%" -Action ScanSourceFiles %*
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -MigrationName "%MIGRATION_NAME%" -Action ScanSourceFiles %*
)

set "EXIT_CODE=%ERRORLEVEL%"
exit /b %EXIT_CODE%
