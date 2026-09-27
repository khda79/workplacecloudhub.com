@echo off
setlocal EnableExtensions
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" (
    echo [%date% %time%] PowerShell 7 x64 is required.
    exit /b 1
)
pushd "%~dp0..\..\PreparedEvidence" || exit /b 1
"%PWSH%" -NoProfile -File "SmartM365-WorkplaceEvidence-Prepare.ps1" -Tenant prod -RepairLegacyHistory %*
set "RESULT=%ERRORLEVEL%"
popd
exit /b %RESULT%
