@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "UNC_WORK_DIR=%~dp0..\..\M365Inventory\Security\SecureScore\."
set "SCRIPT_PATH=%UNC_WORK_DIR%\SmartM365-SecureScore-Inventory.ps1"
set "PWSH="
set "PWSH_X64=%ProgramFiles%\PowerShell\7\pwsh.exe"
set "PWSH_X86="
if not "%ProgramFiles(x86)%"=="" set "PWSH_X86=%ProgramFiles(x86)%\PowerShell\7\pwsh.exe"
if exist "%PWSH_X64%" set "PWSH=%PWSH_X64%"
if not defined PWSH if defined PWSH_X86 if exist "%PWSH_X86%" set "PWSH=%PWSH_X86%"
if not defined PWSH for /f "delims=" %%I in ('where pwsh.exe 2^>nul') do if not defined PWSH set "PWSH=%%I"
if not defined PWSH (
    call :Log "PowerShell 7 (pwsh.exe) was not found."
    call :Log "Install PowerShell 7 or add pwsh.exe to PATH."
    exit /b 1
)
if not exist "%SCRIPT_PATH%" (
    call :Log "Script not found: %SCRIPT_PATH%"
    exit /b 2
)

"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -Tenant prod -EnableConfiguredExternalActions %*
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" call :Log "Secure Score inventory failed with exit code %EXIT_CODE%."
exit /b %EXIT_CODE%

:Log
set "STAMP="
for /f "delims=" %%T in ('powershell.exe -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd HH:mm:ss'"') do set "STAMP=%%T"
echo [!STAMP!] %~1
exit /b 0
