@echo off
setlocal

set "PROJECT_ROOT=%~dp0."
set "BATCH_SCRIPT=%~dp0Scripts\Launchers\Generic\SmartM365-SharePointMigration-SourceScanBatch.ps1"
set "PS51=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS51=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%BATCH_SCRIPT%" (
    "%PS51%" -NoProfile -Command "Write-Host ('[{0}] Source batch script not found.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))"
    exit /b 1
)

"%PS51%" -NoProfile -ExecutionPolicy Bypass -File "%BATCH_SCRIPT%" -ProjectRoot "%PROJECT_ROOT%" %*
exit /b %ERRORLEVEL%
