@echo off
setlocal EnableExtensions
set "POWERSHELL5=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "POWERSHELL5=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%POWERSHELL5%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\..\Orchestrator\Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1" -UninstallInteractive %*
exit /b %ERRORLEVEL%
