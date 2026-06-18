@echo off
setlocal

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo Failed to switch to the launcher directory.
    pause
    exit /b 1
)

set "SCRIPT_DIR=%CD%\"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"

if exist "%PWSH%" (
    start "SmartM365 Intune Remediation" "%PWSH%" -STA -NoProfile -WindowStyle Hidden -File "%SCRIPT_DIR%GUI\SmartM365-IntuneRemediation-GUI.ps1" %*
) else (
    start "SmartM365 Intune Remediation" powershell.exe -STA -NoProfile -WindowStyle Hidden -File "%SCRIPT_DIR%GUI\SmartM365-IntuneRemediation-GUI.ps1" %*
)

popd
exit /b 0
