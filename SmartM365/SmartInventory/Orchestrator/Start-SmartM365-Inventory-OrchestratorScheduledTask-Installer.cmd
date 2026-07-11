@echo off
setlocal EnableExtensions DisableDelayedExpansion
rem Version 1.1.0

set "INSTALLER_PATH=%~dp0Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1"
set "WINDOWS_POWERSHELL=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

"%WINDOWS_POWERSHELL%" -NoProfile -ExecutionPolicy Bypass -File "%INSTALLER_PATH%" -Interactive
exit /b %ERRORLEVEL%
