@echo off
setlocal EnableExtensions DisableDelayedExpansion
rem Version 0.1.0

set "INSTALLER=%~dp0..\..\Orchestration\Install-SmartWorkplaceCMDB-OrchestratorScheduledTask.ps1"
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

"%PWSH%" -NoProfile -ExecutionPolicy Bypass -File "%INSTALLER%" -Interactive %*
exit /b %ERRORLEVEL%
