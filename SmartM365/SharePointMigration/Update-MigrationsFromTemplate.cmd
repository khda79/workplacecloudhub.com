@echo off
setlocal
echo Script  : %~nx0 v1.0.1

set "UNC_WORK_DIR=%~dp0."

set "PROJECT_ROOT=%UNC_WORK_DIR%"
set "MIGRATIONS_ROOT=%PROJECT_ROOT%\Migrations"
set "SCRIPT_PATH=%PROJECT_ROOT%\Scripts\Operations\SmartM365-SharePointMigration-UpdateFromTemplate.ps1"

if not exist "%SCRIPT_PATH%" (
    echo Update script not found:
    echo %SCRIPT_PATH%
    pause
    exit /b 1
)

where pwsh.exe >nul 2>&1
if errorlevel 1 (
    set "POWERSHELL_EXE=powershell.exe"
) else (
    set "POWERSHELL_EXE=pwsh.exe"
)

echo Using PowerShell: %POWERSHELL_EXE%
echo Migrations root: %MIGRATIONS_ROOT%
echo.

"%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -MigrationsRoot "%MIGRATIONS_ROOT%" %*
set "EXIT_CODE=%ERRORLEVEL%"

echo.
echo Exit code: %EXIT_CODE%
pause
exit /b %EXIT_CODE%
