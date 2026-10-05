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

rem Stage the signed entry script locally before executing it from a network toolkit.
rem The batch itself copies and verifies inventory scripts and dependencies locally.
set "SOURCE_BATCH_LOCAL=%TEMP%\SmartM365-SourceBatchLauncher-%RANDOM%-%RANDOM%.ps1"
"%PS51%" -NoProfile -Command "$env:PSModulePath=(Join-Path $PSHOME 'Modules')+[IO.Path]::PathSeparator+$env:PSModulePath; $ErrorActionPreference='Stop'; try { Copy-Item -LiteralPath $env:BATCH_SCRIPT -Destination $env:SOURCE_BATCH_LOCAL -ErrorAction Stop; if ((Get-FileHash -LiteralPath $env:BATCH_SCRIPT -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $env:SOURCE_BATCH_LOCAL -Algorithm SHA256).Hash) { throw 'Local source batch hash differs from the shared script.' }; $signature=Get-AuthenticodeSignature -LiteralPath $env:SOURCE_BATCH_LOCAL; if ($signature.Status -ne 'Valid') { throw ('Local source batch signature is not trusted: '+$signature.Status+'; '+$signature.StatusMessage) }; Write-Host ('[{0}] Local source batch verified: {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$env:SOURCE_BATCH_LOCAL) } catch { Write-Host ('[{0}] Source batch staging failed: {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$_.Exception.Message); exit 1 }"
if errorlevel 1 exit /b 1

"%PS51%" -NoProfile -ExecutionPolicy Bypass -File "%SOURCE_BATCH_LOCAL%" -ProjectRoot "%PROJECT_ROOT%" %*
set "SOURCE_BATCH_EXIT=%ERRORLEVEL%"
"%PS51%" -NoProfile -Command "if (Test-Path -LiteralPath $env:SOURCE_BATCH_LOCAL) { Remove-Item -LiteralPath $env:SOURCE_BATCH_LOCAL -Force -ErrorAction SilentlyContinue }"
exit /b %SOURCE_BATCH_EXIT%
