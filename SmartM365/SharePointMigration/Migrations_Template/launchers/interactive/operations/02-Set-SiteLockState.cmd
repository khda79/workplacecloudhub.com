@echo off
setlocal
echo Script  : %~nx0 v1.0.1

set "SCRIPT_DIR=%~dp0"
set "PROJECT_ROOT=%SCRIPT_DIR%..\..\..\..\.."
rem Migrations stored outside the toolkit: toolkit.root.txt in the migrations root names the toolkit.
if exist "%SCRIPT_DIR%..\..\..\..\toolkit.root.txt" set /p PROJECT_ROOT=<"%SCRIPT_DIR%..\..\..\..\toolkit.root.txt"
pushd "%PROJECT_ROOT%" || (
    echo Failed to access project root from: %SCRIPT_DIR%
    exit /b 1
)

set "MIGRATION_ROOT=%SCRIPT_DIR%..\..\.."
set "PS_SCRIPT=%CD%\Scripts\Launchers\Generic\SmartM365-SharePointMigration-OperationLauncher.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -MigrationRoot "%MIGRATION_ROOT%" -Operation SetSiteLockState %*

set "EXIT_CODE=%ERRORLEVEL%"
popd
echo.
echo Exit code: %EXIT_CODE%
pause
exit /b %EXIT_CODE%
