@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "UNC_WORK_DIR=%~dp0..\..\Orchestrator\."

set "SCRIPT_DIR=%UNC_WORK_DIR%\"
set "PWSH="
set "PWSH_X64=%ProgramFiles%\PowerShell\7\pwsh.exe"
set "PWSH_X86="
if not "%ProgramFiles(x86)%"=="" set "PWSH_X86=%ProgramFiles(x86)%\PowerShell\7\pwsh.exe"
if exist "%PWSH_X64%" set "PWSH=%PWSH_X64%"
if not defined PWSH if defined PWSH_X86 if exist "%PWSH_X86%" set "PWSH=%PWSH_X86%"
if not defined PWSH for /f "delims=" %%I in ('where pwsh.exe 2^>nul') do if not defined PWSH set "PWSH=%%I"
if not defined PWSH (
    exit /b 1
)

"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%SmartM365-Inventory-Orchestrator.ps1" -Tenant prod -SendExecutionSummary
set "EXIT_CODE=%ERRORLEVEL%"
exit /b %EXIT_CODE%
