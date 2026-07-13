@echo off
setlocal EnableExtensions EnableDelayedExpansion

pushd "%~dp0..\..\Orchestrator\" >nul 2>&1
if errorlevel 1 exit /b 1

set "SCRIPT_DIR=%CD%\"
set "PWSH="
set "PWSH_X64=%ProgramFiles%\PowerShell\7\pwsh.exe"
set "PWSH_X86="
if not "%ProgramFiles(x86)%"=="" set "PWSH_X86=%ProgramFiles(x86)%\PowerShell\7\pwsh.exe"
if exist "%PWSH_X64%" set "PWSH=%PWSH_X64%"
if not defined PWSH if defined PWSH_X86 if exist "%PWSH_X86%" set "PWSH=%PWSH_X86%"
if not defined PWSH for /f "delims=" %%I in ('where pwsh.exe 2^>nul') do if not defined PWSH set "PWSH=%%I"
if not defined PWSH (
    popd
    exit /b 1
)

"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-Inventory-Orchestrator.ps1" -Tenant prod -SendExecutionSummary
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
