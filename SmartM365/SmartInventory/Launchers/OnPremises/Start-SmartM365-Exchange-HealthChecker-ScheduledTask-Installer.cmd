@echo off
setlocal EnableExtensions DisableDelayedExpansion
rem Version 1.0.1

set "SOURCE_DIRECTORY=%~dp0..\..\ExchangeInventory\OnPremises\ServersAndStorage\"
set "INSTALLER_PATH=%SOURCE_DIRECTORY%Install-SmartM365-Exchange-HealthCheckerScheduledTask.ps1"
rem Windows PowerShell 5.1 only: under PowerShell 7 the ScheduledTasks module cannot receive CIM sessions.
set "WINDOWS_POWERSHELL=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

"%WINDOWS_POWERSHELL%" -NoProfile -ExecutionPolicy Bypass -File "%INSTALLER_PATH%" -Interactive
set "EXIT_CODE=%ERRORLEVEL%"
exit /b %EXIT_CODE%
