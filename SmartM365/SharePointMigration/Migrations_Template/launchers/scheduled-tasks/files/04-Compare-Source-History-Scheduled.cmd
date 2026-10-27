@echo off
setlocal
echo Script  : %~nx0 v1.0.1

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..\..\..") do set "MIGRATION_NAME=%%~nxI"

set "UNC_WORK_DIR=%SCRIPT_DIR%..\..\..\..\..\."
rem Migrations stored outside the toolkit: toolkit.root.txt in the migrations root names the toolkit.
if exist "%SCRIPT_DIR%..\..\..\..\toolkit.root.txt" set /p UNC_WORK_DIR=<"%SCRIPT_DIR%..\..\..\..\toolkit.root.txt"

set "PS_SCRIPT=%UNC_WORK_DIR%\Scripts\Launchers\Generic\SmartM365-SharePointMigration-Launcher.ps1"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"

if exist "%PWSH%" (
    "%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -MigrationName "%MIGRATION_NAME%" -Action CompareSourceHistory %*
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -MigrationName "%MIGRATION_NAME%" -Action CompareSourceHistory %*
)

set "EXIT_CODE=%ERRORLEVEL%"
exit /b %EXIT_CODE%
