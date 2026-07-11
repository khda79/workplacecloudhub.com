@echo off
setlocal EnableExtensions DisableDelayedExpansion
rem Version 1.1.1

set "SOURCE_DIRECTORY=%~dp0"
set "INSTALLER_PATH=%SOURCE_DIRECTORY%Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1"
set "WINDOWS_POWERSHELL=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

pushd "%SOURCE_DIRECTORY%" >nul 2>&1
if errorlevel 1 (
    echo Failed to access the launcher directory: %SOURCE_DIRECTORY%
    exit /b 1
)

set "PS_ENGINE="
set "PWSH_X64=%ProgramFiles%\PowerShell\7\pwsh.exe"
set "PWSH_X86="
if not "%ProgramFiles(x86)%"=="" set "PWSH_X86=%ProgramFiles(x86)%\PowerShell\7\pwsh.exe"
if exist "%PWSH_X64%" set "PS_ENGINE=%PWSH_X64%"
if not defined PS_ENGINE if defined PWSH_X86 if exist "%PWSH_X86%" set "PS_ENGINE=%PWSH_X86%"
if not defined PS_ENGINE set "PS_ENGINE=%WINDOWS_POWERSHELL%"

"%PS_ENGINE%" -NoProfile -ExecutionPolicy Bypass -File "%INSTALLER_PATH%" -Interactive
set "EXIT_CODE=%ERRORLEVEL%"

popd
exit /b %EXIT_CODE%
