@echo off
setlocal
rem Starts the site lock launcher in 64-bit Windows PowerShell 5.1.
set "PS51=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS51=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
set "LOCK_LAUNCHER=%~dp0Scripts\Launchers\Generic\SmartM365-SharePointMigration-SiteLockStateLauncher.ps1"
"%PS51%" -NoProfile -ExecutionPolicy Bypass -File "%LOCK_LAUNCHER%" %*
exit /b %ERRORLEVEL%
