@echo off
setlocal
rem Starts the page comments launcher in 64-bit Windows PowerShell 5.1.
set "PS51=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS51=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
set "PAGE_COMMENTS_LAUNCHER=%~dp0Scripts\Launchers\Generic\SmartM365-SharePointMigration-PageCommentsDisableLauncher.ps1"
"%PS51%" -NoProfile -ExecutionPolicy Bypass -File "%PAGE_COMMENTS_LAUNCHER%" %*
exit /b %ERRORLEVEL%
