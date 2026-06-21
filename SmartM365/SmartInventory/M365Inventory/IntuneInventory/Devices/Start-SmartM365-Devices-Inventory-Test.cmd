@echo off
setlocal

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "SCRIPT_DIR=%CD%\"
set "PWSH="
if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not defined PWSH if not "%ProgramFiles(x86)%"=="" if exist "%ProgramFiles(x86)%\PowerShell\7\pwsh.exe" set "PWSH=%ProgramFiles(x86)%\PowerShell\7\pwsh.exe"
if not defined PWSH for /f "delims=" %%I in ('where pwsh.exe 2^>nul') do if not defined PWSH set "PWSH=%%I"
if not defined PWSH (
    echo PowerShell 7 ^(pwsh.exe^) was not found.
    echo Install PowerShell 7 or add pwsh.exe to PATH.
    echo Checked:
    echo   %ProgramFiles%\PowerShell\7\pwsh.exe
    if not "%ProgramFiles(x86)%"=="" echo   %ProgramFiles(x86)%\PowerShell\7\pwsh.exe
    echo   PATH
    pause
    popd
    exit /b 1
)
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-Devices-Inventory.ps1" -Tenant test -Connect
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
