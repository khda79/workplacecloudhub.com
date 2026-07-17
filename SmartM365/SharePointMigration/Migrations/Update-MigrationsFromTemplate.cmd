@echo off
setlocal

set "UNC_WORK_DIR=%~dp0."

set "MIGRATIONS_ROOT=%UNC_WORK_DIR%"
for %%I in ("%MIGRATIONS_ROOT%\..") do set "PROJECT_ROOT=%%~fI"
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
