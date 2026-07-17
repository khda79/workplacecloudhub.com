@echo off
setlocal

set "UNC_WORK_DIR=%~dp0."

set "SCRIPT_DIR=%UNC_WORK_DIR%"
for %%I in ("%SCRIPT_DIR%\..") do set "PROJECT_ROOT=%%~fI"
set "PS_SCRIPT=%PROJECT_ROOT%\Scripts\SmartM365-SharePointMigration-InstallPortablePython.ps1"

if not exist "%PS_SCRIPT%" (
    echo Portable Python installer not found:
    echo %PS_SCRIPT%
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
echo Script: %PS_SCRIPT%
echo.

"%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*
set "EXIT_CODE=%ERRORLEVEL%"

echo.
echo Exit code: %EXIT_CODE%
pause
exit /b %EXIT_CODE%
