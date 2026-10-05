@echo off
setlocal

set "PROJECT_ROOT=%~dp0."
set "BATCH_SCRIPT=%~dp0Scripts\Launchers\Generic\SmartM365-SharePointMigration-ComparisonBatch.ps1"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" set "PWSH=pwsh.exe"

if not exist "%BATCH_SCRIPT%" (
    powershell.exe -NoProfile -Command "Write-Host ('[{0}] Comparison batch script not found.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))"
    exit /b 1
)

"%PWSH%" -NoProfile -Command "if ($PSVersionTable.PSVersion -ge [version]'7.4') { exit 0 } else { exit 1 }" >nul 2>&1
if errorlevel 1 (
    powershell.exe -NoProfile -Command "Write-Host ('[{0}] PowerShell 7.4 or later is required for comparisons.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))"
    exit /b 1
)

"%PWSH%" -NoProfile -File "%BATCH_SCRIPT%" -ProjectRoot "%PROJECT_ROOT%" %*
exit /b %ERRORLEVEL%
