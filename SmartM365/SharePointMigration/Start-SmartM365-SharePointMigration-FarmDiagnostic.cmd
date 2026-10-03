@echo off
setlocal
rem Starts only the farm diagnostic launcher in 64-bit Windows PowerShell 5.1.
set "PS51=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS51=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
set "DIAG_LAUNCHER=%~dp0Scripts\Diagnostics\SmartM365-SharePointMigration-FarmDiagnosticLauncher.ps1"
"%PS51%" -NoProfile -ExecutionPolicy Bypass -File "%DIAG_LAUNCHER%" %*
exit /b %ERRORLEVEL%
