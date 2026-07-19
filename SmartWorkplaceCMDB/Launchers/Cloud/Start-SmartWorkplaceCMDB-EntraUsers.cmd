@echo off
setlocal EnableExtensions EnableDelayedExpansion
set "ORCHESTRATOR=%~dp0..\..\Orchestration\SmartWorkplaceCMDB-Orchestrator.ps1"
set "PWSH="
set "PWSH_X64=%ProgramFiles%\PowerShell\7\pwsh.exe"
set "PWSH_X86="
if not "%ProgramFiles(x86)%"=="" set "PWSH_X86=%ProgramFiles(x86)%\PowerShell\7\pwsh.exe"
if exist "%PWSH_X64%" set "PWSH=%PWSH_X64%"
if not defined PWSH if defined PWSH_X86 if exist "%PWSH_X86%" set "PWSH=%PWSH_X86%"
if not defined PWSH for /f "delims=" %%I in ('where pwsh.exe 2^>nul') do if not defined PWSH set "PWSH=%%I"
if not defined PWSH (
    echo PowerShell 7 ^(pwsh.exe^) was not found.
    exit /b 1
)
"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%ORCHESTRATOR%" -Tenant prod -Collect -Pipeline EntraUsers %*
exit /b %ERRORLEVEL%