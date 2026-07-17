@echo off
setlocal

set "UNC_WORK_DIR=%~dp0."

set "SCRIPT_DIR=%UNC_WORK_DIR%\"
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"

if exist "%PWSH%" (
    start "SmartM365 Intune Remediation" "%PWSH%" -STA -NoProfile -WindowStyle Hidden -File "%SCRIPT_DIR%GUI\SmartM365-IntuneRemediation-GUI.ps1" %*
) else (
    start "SmartM365 Intune Remediation" powershell.exe -STA -NoProfile -WindowStyle Hidden -File "%SCRIPT_DIR%GUI\SmartM365-IntuneRemediation-GUI.ps1" %*
)
exit /b 0
