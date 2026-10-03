@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%..\..\..\..\.." || (
    echo Failed to access project root from: %SCRIPT_DIR%
    exit /b 1
)

set "MIGRATION_ROOT=%SCRIPT_DIR%..\..\.."
set "PS_SCRIPT=%CD%\Scripts\Launchers\Generic\SmartM365-SharePointMigration-OperationLauncher.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -MigrationRoot "%MIGRATION_ROOT%" -Operation DisablePageComments %*

set "EXIT_CODE=%ERRORLEVEL%"
popd
echo.
echo Exit code: %EXIT_CODE%
pause
exit /b %EXIT_CODE%
