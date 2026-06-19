@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..\..\..") do set "MIGRATION_NAME=%%~nxI"

pushd "%SCRIPT_DIR%..\..\..\..\.." || (
    echo Failed to access project root from: %SCRIPT_DIR%
    pause
    exit /b 1
)

set "PS_SCRIPT=%CD%\Scripts\Launchers\Generic\SmartM365-SharePointMigration-Launcher.ps1"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"

if exist "%PWSH%" (
    "%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -MigrationName "%MIGRATION_NAME%" -Action ScanTargetFiles %*
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -MigrationName "%MIGRATION_NAME%" -Action ScanTargetFiles %*
)

set "EXIT_CODE=%ERRORLEVEL%"
popd
echo.
echo Exit code: %EXIT_CODE%
pause
exit /b %EXIT_CODE%
