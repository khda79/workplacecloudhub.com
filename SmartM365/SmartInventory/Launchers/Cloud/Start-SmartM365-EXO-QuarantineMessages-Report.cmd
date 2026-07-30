@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "UNC_WORK_DIR=%~dp0..\..\ExchangeInventory\Quarantine\."
set "SCRIPT_DIR=%UNC_WORK_DIR%\"
set "PWSH="
set "PWSH_X64=%ProgramFiles%\PowerShell\7\pwsh.exe"
set "PWSH_X86="
if not "%ProgramFiles(x86)%"=="" set "PWSH_X86=%ProgramFiles(x86)%\PowerShell\7\pwsh.exe"
if exist "%PWSH_X64%" set "PWSH=%PWSH_X64%"
if not defined PWSH if defined PWSH_X86 if exist "%PWSH_X86%" set "PWSH=%PWSH_X86%"
if not defined PWSH for /f "delims=" %%I in ('where pwsh.exe 2^>nul') do if not defined PWSH set "PWSH=%%I"
if not defined PWSH (
    set "STAMP="
    for /f "delims=" %%T in ('powershell.exe -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd HH:mm:ss'"') do set "STAMP=%%T"
    echo [!STAMP!] PowerShell 7 ^(pwsh.exe^) was not found.
    echo [!STAMP!] Install PowerShell 7 or add pwsh.exe to PATH.
    echo [!STAMP!] Checked:
    echo [!STAMP!]   !PWSH_X64!
    if defined PWSH_X86 echo [!STAMP!]   !PWSH_X86!
    echo [!STAMP!]   PATH
    pause
    exit /b 1
)

"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-EXO-QuarantineMessages-Report.ps1" -Tenant prod -Connect
set "EXIT_CODE=%ERRORLEVEL%"
exit /b %EXIT_CODE%
